document.addEventListener('DOMContentLoaded', (event) => {
    const input = document.getElementById('numero_parte');
    if (input) {
        input.focus();
    }

    const form = document.querySelector('form');
    if (form) {
        form.addEventListener('submit', () => {
            setTimeout(() => {
                input.focus();
            }, 10);
        });
    }
});
