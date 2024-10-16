const express = require('express');
const xlsx = require('xlsx');
const bodyParser = require('body-parser');
const cors = require('cors');
const fs = require('fs');
const { google } = require('googleapis');
const app = express();
const port = 3000;

// Usar credenciales de OAuth2
const SCOPES = ['https://www.googleapis.com/auth/drive.file']; // Permisos para Google Drive
const credentials = require('./credentials.json'); // El archivo JSON que descargaste de Google Cloud
const { client_secret, client_id, redirect_uris } = credentials.installed;
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

// Configurar token (deberías haber generado un token de acceso al iniciar sesión en Google)
oAuth2Client.setCredentials(require('./token.json'));

// Crear el servicio de Google Drive
const drive = google.drive({ version: 'v3', auth: oAuth2Client });

app.use(cors()); // Permite el CORS para peticiones desde el frontend
app.use(bodyParser.json()); // Para poder recibir JSON en el body
app.use(express.static('public')); // Para servir archivos estáticos
app.use('/images', express.static('images')); // Para mostrar imágenes

const rutaExcel = './Lesiones FCB24-25.xlsx'; // Ruta del archivo Excel local (temporal)

// Ruta para manejar solicitudes GET en la raíz
app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html'); // Enviar el archivo HTML
});

// Ruta para agregar nuevos datos al Excel y subirlo a Google Drive
app.post('/actualizar-excel', async (req, res) => {
    const {
        nombre, equipo, parteCuerpo, incidencia, origen, circumstancias, lesion,
        metodoLesional, metodoEspecifico, tipoEspecifico, observaciones,
        fecha_lesion, fecha_alta, diagnostico
    } = req.body;

    // Cargar el archivo Excel existente o crear uno nuevo si no existe
    let wb;
    let ws;

    try {
        wb = xlsx.readFile(rutaExcel);
        ws = wb.Sheets[wb.SheetNames[0]]; // Obtener la primera hoja
    } catch (error) {
        // Si no se puede leer el archivo, creamos uno nuevo
        wb = xlsx.utils.book_new(); // Crear un nuevo libro
        ws = xlsx.utils.aoa_to_sheet([
            ["Nombre", "Equipo", "Cuerpo", "Incidencia", "Origen", "Circumstancias", "Tipo de Lesión", "Método Lesional", "Método Específico", "Tipo Específico", "Observaciones", "Fecha de Lesion", "Fecha del alta", "Diagnóstico"]
        ]);
        xlsx.utils.book_append_sheet(wb, ws, "Lesiones");
    }

    // Obtener las filas existentes y agregar los nuevos datos
    const datosExistentes = xlsx.utils.sheet_to_json(ws, { header: 1 });
    datosExistentes.push([nombre, equipo, parteCuerpo, incidencia, origen, circumstancias, lesion, metodoLesional, metodoEspecifico, tipoEspecifico, observaciones, fecha_lesion, fecha_alta, diagnostico]);

    // Actualizar la hoja con los nuevos datos
    const wsActualizado = xlsx.utils.aoa_to_sheet(datosExistentes);
    wb.Sheets[wb.SheetNames[0]] = wsActualizado;

    // Guardar el archivo Excel localmente (temporalmente)
    xlsx.writeFile(wb, rutaExcel);

    try {
        // Subir o actualizar el archivo en Google Drive
        const fileMetadata = {
            'name': 'Lesiones FCB24-25.xlsx', // Nombre del archivo en Google Drive
            'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        };
        const media = {
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            body: fs.createReadStream(rutaExcel) // Leer el archivo temporalmente guardado
        };

        // Subir o actualizar el archivo en Drive
        const response = await drive.files.create({
            resource: fileMetadata,
            media: media,
            fields: 'id'
        });

        console.log('Archivo subido a Google Drive con ID:', response.data.id);
        res.json({ message: 'Excel actualizado y subido a Google Drive exitosamente' });
    } catch (error) {
        console.error('Error al subir el archivo a Google Drive:', error);
        res.status(500).json({ message: 'Error al subir el archivo a Google Drive' });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Servidor corriendo en el puerto http://localhost:${PORT}`);
});
