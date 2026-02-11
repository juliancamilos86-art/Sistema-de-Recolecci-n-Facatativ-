// Gestión de Ferias UNIAGRARIA 2026
let feriaActual = null;

function inicializarDropzone(elementId, feriaId) {
    const dropzone = new Dropzone(`#${elementId}`, {
        url: `/api/ferias/${feriaId}/imagenes`,
        headers: {
            'X-CSRFToken': getCookie('session')
        },
        maxFilesize: 10,
        acceptedFiles: 'image/*',
        dictDefaultMessage: 'Arrastra las imágenes aquí o haz clic para subir',
        thumbnailWidth: 300,
        thumbnailHeight: 300
    });
    
    dropzone.on('success', function(file, response) {
        toastr.success('Imagen subida exitosamente');
        cargarGalería(feriaId);
    });
    
    dropzone.on('error', function(file, errorMessage) {
        toastr.error(`Error: ${errorMessage}`);
    });
    
    return dropzone;
}

function cargarGalería(feriaId) {
    fetch(`/api/ferias/${feriaId}/imagenes`)
        .then(response => response.json())
        .then(data => {
            const grid = document.getElementById('galeriaGrid');
            grid.innerHTML = '';
            
            data.imagenes.forEach(img => {
                grid.appendChild(crearTarjetaImagen(img));
            });
        });
}

function crearTarjetaImagen(imagen) {
    const col = document.createElement('div');
    col.className = 'col-md-4 mb-3';
    col.innerHTML = `
        <div class="card feria-card">
            <img src="${imagen.url}" class="card-img-top" style="height: 200px; object-fit: cover;">
            <div class="card-body">
                <p class="small text-muted">
                    <i class="fas fa-user me-1"></i>${imagen.usuario || 'N/A'}<br>
                    <i class="fas fa-clock me-1"></i>${new Date(imagen.fecha).toLocaleDateString()}
                </p>
                <button class="btn btn-sm btn-danger" onclick="eliminarImagen('${imagen.public_id}')">
                    <i class="fas fa-trash"></i>
                </button>
                <button class="btn btn-sm btn-primary" onclick="copiarURL('${imagen.url}')">
                    <i class="fas fa-link"></i>
                </button>
            </div>
        </div>
    `;
    return col;
}

function copiarURL(url) {
    navigator.clipboard.writeText(url);
    toastr.success('URL copiada al portapapeles');
}