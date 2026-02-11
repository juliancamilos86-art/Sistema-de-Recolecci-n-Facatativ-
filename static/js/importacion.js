// Importación de Datos UNIAGRARIA 2026
function validarArchivo(archivo) {
    const extensiones = ['.xlsx', '.xls', '.csv'];
    const nombre = archivo.name.toLowerCase();
    
    for (let ext of extensiones) {
        if (nombre.endsWith(ext)) {
            return true;
        }
    }
    return false;
}

function previewDatos(archivo) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        // Mostrar preview de primeras filas
        if (archivo.name.endsWith('.csv')) {
            const lines = e.target.result.split('\n').slice(0, 5);
            mostrarPreviewCSV(lines);
        }
    };
    
    if (archivo.name.endsWith('.csv')) {
        reader.readAsText(archivo);
    }
}