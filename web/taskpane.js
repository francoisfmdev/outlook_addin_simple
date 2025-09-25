Office.onReady(() => {
  const btn = document.getElementById('noop');
  const status = document.getElementById('status');
  btn.addEventListener('click', () => {
    // volontairement vide (il ne fait rien)
    status.textContent = 'Bouton cliqué (aucune action).';
  });
});
