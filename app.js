const express = require('express');
const xlsx = require('xlsx');
const bodyParser = require('body-parser');
const cors = require('cors');
const app = express();
const port = 3000;

app.use(cors()); // Permite el CORS para peticiones desde el frontend
app.use(bodyParser.json()); // Para poder recibir JSON en el body
app.use(express.static('public')); // Para servir archivos estáticos
app.use('/images', express.static('images')); // Para mostrar imágenes

const rutaExcel = './Lesiones FCB24-25.xlsx'; // Ruta del archivo Excel

// Ruta para manejar solicitudes GET en la raíz
app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html'); // Enviar el archivo HTML
});

// Ruta para agregar nuevos datos al Excel
app.post('/actualizar-excel', (req, res) => {
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
        xlsx.writeFile(wb, rutaExcel);
    }

    // Obtener las filas existentes
    const datosExistentes = xlsx.utils.sheet_to_json(ws, { header: 1 });

    // Agregar los nuevos datos (incluir parte del cuerpo)
    datosExistentes.push([nombre, equipo, parteCuerpo, incidencia, origen, circumstancias, lesion, metodoLesional, metodoEspecifico, tipoEspecifico, observaciones, fecha_lesion, fecha_alta, diagnostico]);

    // Actualizar la hoja con los nuevos datos
    const wsActualizado = xlsx.utils.aoa_to_sheet(datosExistentes);

    // Definir el ancho de las columnas
    wsActualizado['!cols'] = [
        { wch: 30 },  // Ancho de la columna para "Nombre"
        { wch: 20 },  // Ancho de la columna para "Equipo"
        { wch: 20 },  // Ancho de la columna para "Cuerpo"
        { wch: 20 },  // Ancho de la columna para "Incidencia"
        { wch: 20 },  // Ancho de la columna para "Origen"
        { wch: 25 },  // Ancho de la columna para "Circumstancias"
        { wch: 20 },  // Ancho de la columna para "Tipo de Lesión"
        { wch: 20 },  // Ancho de la columna para "Método Lesional"
        { wch: 20 },  // Ancho de la columna para "Método Específico"
        { wch: 25 },  // Ancho de la columna para "Tipo Específico"
        { wch: 10 },  // Ancho de la columna para "Observaciones"
        { wch: 15 },  // Ancho de la columna para "Fecha de Lesión"
        { wch: 15 },  // Ancho de la columna para "Fecha del alta"
        { wch: 25 }   // Ancho de la columna para "Diagnóstico"
    ];

    wb.Sheets[wb.SheetNames[0]] = wsActualizado;

    // Guardar el archivo Excel actualizado
    xlsx.writeFile(wb, rutaExcel);

    res.json({ message: 'Excel actualizado exitosamente' });
});

app.listen(port, () => {
    console.log(`Servidor corriendo en http://localhost:${port}`);
});
