const express = require('express');
const xlsx = require('xlsx');
const app = express();

// Ruta para manejar los datos del formulario y generar el archivo Excel
app.post('/generar-excel', (req, res) => {
    // Obtener los datos del formulario desde req.body
    const datosFormulario = req.body;

    // Crear un nuevo libro de Excel
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.json_to_sheet(datosFormulario);

    // Agregar la hoja al libro
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Datos de Formulario');

    // Escribir el libro en un buffer
    const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    // Configurar la respuesta para descargar el archivo Excel
    res.setHeader('Content-Disposition', 'attachment; filename=formulario.xlsx');
    res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);

});

app.listen(9000, () => {
    console.log('Servidor escuchando en el puerto 9000');
});