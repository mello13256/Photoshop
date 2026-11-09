// Detecta o sistema, destaca o download certo e mostra a versão mais recente do GitHub.
(() => {
  const REPO = 'mello13256/Photoshop';
  const BASE = `https://github.com/${REPO}/releases/latest/download/`;
  const PLATFORMS = {
    windows: { label: 'Baixar para Windows', sub: 'Instalador · 64 bits', file: 'LuminaStudio-Setup.exe' },
    mac: { label: 'Baixar para Mac', sub: 'Apple Silicon (M1 ou mais novo)', file: 'LuminaStudio-macOS.dmg' },
    linux: { label: 'Baixar para Linux', sub: 'AppImage · 64 bits', file: 'LuminaStudio-Linux.AppImage' },
  };

  const ua = (navigator.userAgentData?.platform || navigator.platform || navigator.userAgent || '').toLowerCase();
  const os = /win/.test(ua) ? 'windows' : /mac|iphone|ipad/.test(ua) ? 'mac' : /linux|x11|cros/.test(ua) && !/android/.test(ua) ? 'linux' : null;
  const mobile = /android|iphone|ipad|mobile/.test((navigator.userAgent || '').toLowerCase());

  const main = document.getElementById('main-download');
  if (os && !mobile) {
    const p = PLATFORMS[os];
    main.href = BASE + p.file;
    document.getElementById('main-label').textContent = p.label;
    document.getElementById('main-sub').textContent = p.sub;
    document.querySelector(`.dl[data-os="${os}"]`)?.classList.add('recommended');
  } else if (mobile) {
    main.href = '#baixar';
    document.getElementById('main-label').textContent = 'Ver downloads';
    document.getElementById('main-sub').textContent = 'O Lumina Studio é para computador';
  }

  // Vitrine de telas
  const img = document.getElementById('showcase-img');
  const names = { docs: 'Lumina Texto', sheets: 'Lumina Planilhas', slides: 'Lumina Apresentações', photo: 'Lumina Foto' };
  for (const tab of document.querySelectorAll('.tab')) {
    tab.addEventListener('click', () => {
      for (const t of document.querySelectorAll('.tab')) { t.classList.toggle('on', t === tab); t.setAttribute('aria-selected', String(t === tab)); }
      img.src = `img/${tab.dataset.shot}.webp`;
      img.alt = names[tab.dataset.shot];
    });
  }

  // Versão e tamanhos do release mais recente
  const mb = (n) => `${(n / 1048576).toFixed(0)} MB`;
  fetch(`https://api.github.com/repos/${REPO}/releases/latest`, { headers: { Accept: 'application/vnd.github+json' } })
    .then((r) => (r.ok ? r.json() : Promise.reject(new Error(r.status))))
    .then((rel) => {
      const date = new Date(rel.published_at).toLocaleDateString('pt-BR', { day: 'numeric', month: 'long', year: 'numeric' });
      document.getElementById('version').textContent = `Versão ${rel.tag_name.replace(/^v/, '')} · publicada em ${date}`;
      for (const a of rel.assets || []) {
        const el = document.querySelector(`[data-size="${a.name}"]`);
        if (el) el.textContent = mb(a.size);
      }
    })
    .catch(() => { /* sem rede ou limite da API: os links continuam funcionando */ });
})();
