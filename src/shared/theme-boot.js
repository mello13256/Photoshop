// Aplica o tema antes da primeira pintura (script clássico, sem módulos).
(function () {
  try {
    var p = localStorage.getItem('lumina.theme') || 'system';
    document.documentElement.dataset.theme = p === 'system' ? (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light') : p;
  } catch (e) { /* armazenamento indisponível */ }
})();
