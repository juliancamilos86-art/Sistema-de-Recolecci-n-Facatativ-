// Reportes UNIAGRARIA 2026
function descargarReporte(tipo) {
    const params = new URLSearchParams({
        periodo: document.getElementById('periodoSelect').value,
        formato: tipo
    });
    
    window.location.href = `/api/reportes/exportar/${tipo}?${params.toString()}`;
}

function imprimirReporte() {
    window.print();
}

function compartirReporte() {
    const url = window.location.href;
    
    if (navigator.share) {
        navigator.share({
            title: 'Reporte UNIAGRARIA 2026',
            text: 'Reporte de recolección de datos',
            url: url
        });
    } else {
        navigator.clipboard.writeText(url);
        toastr.success('URL copiada al portapapeles');
    }
}