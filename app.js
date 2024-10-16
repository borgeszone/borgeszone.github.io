const express = require('express');
const xlsx = require('xlsx');
const bodyParser = require('body-parser');
const cors = require('cors');
const fs = require('fs');
const { google } = require('googleapis');
const app = express();
const port = 3000;

app.use(cors());
app.use(bodyParser.json());
app.use(express.static('public'));

// Configura las credenciales desde las variables de entorno
const oAuth2Client = new google.auth.OAuth2(
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET,
  process.env.REDIRECT_URI
);

// Usar el token de actualización para obtener tokens de acceso automáticamente
oAuth2Client.setCredentials({
  refresh_token: process.env.REFRESH_TOKEN
});

// Crear el servicio de Google Drive
const drive = google.drive({ version: 'v3', auth: oAuth2Client });

const rutaExcel = './Lesiones FCB24-25.xlsx';

// Ruta para manejar solicitudes GET en la raíz
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html');
});

// Ruta para agregar nuevos datos al Excel y subirlo a Google Drive
app.post('/actualizar-excel', async (req, res) => {
  const {
    nombre, equipo, parteCuerpo, incidencia, origen, circumstancias, lesion,
    metodoLesional, metodoEspecifico, tipoEspecifico, observaciones,
    fecha_lesion, fecha_alta, diagnostico
  } = req.body;

  let wb;
  let ws;

  try {
    wb = xlsx.readFile(rutaExcel);
    ws = wb.Sheets[wb.SheetNames[0]];
  } catch (error) {
    // Si no se puede leer el archivo, creamos uno nuevo
    wb = xlsx.utils.book_new();
    ws = xlsx.utils.aoa_to_sheet([
      ["Nombre", "Equipo", "Cuerpo", "Incidencia", "Origen", "Circumstancias", "Tipo de Lesión", "Método Lesional", "Método Específico", "Tipo Específico", "Observaciones", "Fecha de Lesion", "Fecha del alta", "Diagnóstico"]
    ]);
    xlsx.utils.book_append_sheet(wb, ws, "Lesiones");
  }

  // Agregar los nuevos datos
  const datosExistentes = xlsx.utils.sheet_to_json(ws, { header: 1 });
  datosExistentes.push([nombre, equipo, parteCuerpo, incidencia, origen, circumstancias, lesion, metodoLesional, metodoEspecifico, tipoEspecifico, observaciones, fecha_lesion, fecha_alta, diagnostico]);

  const wsActualizado = xlsx.utils.aoa_to_sheet(datosExistentes);
  wb.Sheets[wb.SheetNames[0]] = wsActualizado;

  // Guardar el archivo Excel localmente
  xlsx.writeFile(wb, rutaExcel);

  try {
    // Subir el archivo a Google Drive
    const fileMetadata = {
      'name': 'Lesiones FCB24-25.xlsx',
      'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    };
    const media = {
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      body: fs.createReadStream(rutaExcel)
    };

    // Subir el archivo
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
