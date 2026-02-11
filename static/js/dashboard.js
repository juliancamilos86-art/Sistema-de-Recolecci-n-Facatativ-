// Dashboard UNIAGRARIA 2026
document.addEventListener('DOMContentLoaded', function() {
    // Auto-refresh cada 5 minutos
    setTimeout(function() {
        location.reload();
    }, 300000);
});

function actualizarDashboard() {
    fetch('/api/dashboard/stats')
        .then(response => response.json())
        .then(data => {
            // Actualizar números
            document.querySelectorAll('.estadistica-numero').forEach(el => {
                // Animación de conteo
                const target = parseInt(el.textContent);
                animateValue(el, 0, target, 1000);
            });
        });
}

function animateValue(element, start, end, duration) {
    const range = end - start;
    const increment = range / (duration / 10);
    let current = start;
    
    const timer = setInterval(() => {
        current += increment;
        if (current >= end) {
            current = end;
            clearInterval(timer);
        }
        element.textContent = Math.round(current);
    }, 10);
}