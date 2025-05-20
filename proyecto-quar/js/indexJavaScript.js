window.addEventListener('scroll', function() {
    const header = document.querySelector('header');
    if (window.scrollY > 0) {
        header.classList.add('scrolled');
    } else {
        header.classList.remove('scrolled');
    }
});

document.addEventListener('DOMContentLoaded', function() {
    const titulos = document.querySelectorAll('.titulo-noticia');
    titulos.forEach(function(titulo) {
        titulo.addEventListener('click', function() {
            const contenido = this.nextElementSibling;
            if (window.getComputedStyle(contenido).display === 'none') {
                contenido.style.display = 'block';
                this.querySelector('.toggle').textContent = '–';
                this.classList.add('active'); // Agregar clase active
            } else {
                contenido.style.display = 'none';
                this.querySelector('.toggle').textContent = '+';
                this.classList.remove('active'); // Quitar clase active
            }
        });
    });
});




