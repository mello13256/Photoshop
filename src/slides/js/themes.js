// Temas, layouts e modelos de apresentação.

export const W = 1920;
export const H = 1080;

export const THEMES = {
  claro: { name: 'Claro', bg: { type: 'color', color: '#ffffff' }, title: '#1d2129', text: '#4a5061', accent: '#3b6cf6', accent2: '#e8edff', head: '"Segoe UI Variable", "Segoe UI", Inter, system-ui, sans-serif', body: '"Segoe UI", Inter, system-ui, sans-serif' },
  aurora: { name: 'Aurora', bg: { type: 'gradient', from: '#6a5cff', to: '#ff5ea8', angle: 135 }, title: '#ffffff', text: '#f3eeff', accent: '#ffe066', accent2: 'rgba(255,255,255,.16)', head: '"Segoe UI Variable", "Segoe UI", Inter, sans-serif', body: '"Segoe UI", Inter, sans-serif' },
  midnight: { name: 'Meia-noite', bg: { type: 'gradient', from: '#0f172a', to: '#1e293b', angle: 160 }, title: '#f8fafc', text: '#cbd5e1', accent: '#38bdf8', accent2: '#1e3a5f', head: '"Segoe UI Variable", "Segoe UI", Inter, sans-serif', body: '"Segoe UI", Inter, sans-serif' },
  coral: { name: 'Coral', bg: { type: 'gradient', from: '#ff9a76', to: '#ff6a88', angle: 135 }, title: '#ffffff', text: '#fff5f0', accent: '#3d1a3a', accent2: 'rgba(255,255,255,.22)', head: 'Georgia, "Times New Roman", serif', body: '"Segoe UI", Inter, sans-serif' },
  forest: { name: 'Floresta', bg: { type: 'gradient', from: '#0b3d2e', to: '#1f6f4a', angle: 160 }, title: '#ecfdf5', text: '#d1fae5', accent: '#a7f3d0', accent2: 'rgba(167,243,208,.15)', head: '"Segoe UI Variable", "Segoe UI", Inter, sans-serif', body: '"Segoe UI", Inter, sans-serif' },
  pitch: { name: 'Pitch', bg: { type: 'color', color: '#0b0b0f' }, title: '#ffffff', text: '#b4b4c0', accent: '#7c5cff', accent2: '#1d1a2e', head: '"Segoe UI Variable", "Segoe UI", Inter, sans-serif', body: '"Segoe UI", Inter, sans-serif' },
  papel: { name: 'Papel', bg: { type: 'color', color: '#f6f1e7' }, title: '#2b2118', text: '#4b3f33', accent: '#c2410c', accent2: '#eadfcb', head: 'Georgia, "Times New Roman", serif', body: 'Georgia, serif' },
  oceano: { name: 'Oceano', bg: { type: 'gradient', from: '#0077b6', to: '#00b4d8', angle: 135 }, title: '#ffffff', text: '#e6f7ff', accent: '#ffd166', accent2: 'rgba(255,255,255,.18)', head: '"Segoe UI Variable", "Segoe UI", Inter, sans-serif', body: '"Segoe UI", Inter, sans-serif' },
};

export const LAYOUTS = {
  title: 'Título',
  titleContent: 'Título e conteúdo',
  twoCols: 'Duas colunas',
  section: 'Cabeçalho de seção',
  imageText: 'Imagem e texto',
  quote: 'Citação',
  bigNumber: 'Número em destaque',
  titleOnly: 'Somente título',
  blank: 'Em branco',
};

let seq = 0;
export const uid = (p = 'e') => `${p}${Date.now().toString(36)}${(seq++).toString(36)}${Math.random().toString(36).slice(2, 5)}`;

const text = (role, x, y, w, h, html, extra = {}) => ({
  id: uid(), type: 'text', role, x, y, w, h, rot: 0, opacity: 1,
  text: { html, size: { title: 88, subtitle: 44, body: 38, caption: 30, big: 220, quote: 64 }[role] || 38, color: null, align: 'left', valign: 'top', bold: role === 'title' || role === 'big', font: null, line: 1.25 },
  ...extra,
});

const shape = (kind, x, y, w, h, extra = {}) => ({
  id: uid(), type: 'shape', x, y, w, h, rot: 0, opacity: 1,
  shape: { kind, fill: 'accent', stroke: null, strokeW: 0, radius: 24 },
  text: { html: '', size: 36, color: '#ffffff', align: 'center', valign: 'middle', bold: true, font: null, line: 1.2 },
  ...extra,
});

export function layoutElements(layout) {
  switch (layout) {
    case 'title':
      return [
        text('title', 160, 360, 1600, 200, '', { text: { ...text('title').text, align: 'center', valign: 'bottom', size: 104 }, ph: 'Título da apresentação' }),
        text('subtitle', 260, 590, 1400, 120, '', { text: { ...text('subtitle').text, align: 'center' }, ph: 'Subtítulo ou nome do autor' }),
      ];
    case 'titleContent':
      return [
        text('title', 120, 90, 1680, 150, '', { ph: 'Título do slide' }),
        text('body', 120, 280, 1680, 700, '', { ph: 'Clique para adicionar o texto', list: true }),
      ];
    case 'twoCols':
      return [
        text('title', 120, 90, 1680, 150, '', { ph: 'Título do slide' }),
        text('body', 120, 280, 800, 700, '', { ph: 'Primeira coluna', list: true }),
        text('body', 1000, 280, 800, 700, '', { ph: 'Segunda coluna', list: true }),
      ];
    case 'section':
      return [
        shape('rect', 160, 420, 16, 240, { shape: { kind: 'rect', fill: 'accent', stroke: null, strokeW: 0, radius: 8 } }),
        text('title', 220, 400, 1500, 180, '', { text: { ...text('title').text, size: 110, valign: 'middle' }, ph: 'Nome da seção' }),
        text('subtitle', 224, 590, 1400, 90, '', { ph: 'Descrição curta' }),
      ];
    case 'imageText':
      return [
        text('title', 120, 120, 820, 200, '', { ph: 'Título' }),
        text('body', 120, 340, 820, 620, '', { ph: 'Explique a ideia aqui', list: true }),
        { id: uid(), type: 'image', x: 1020, y: 0, w: 900, h: 1080, rot: 0, opacity: 1, image: { asset: null, fit: 'cover', radius: 0 }, ph: 'Clique duas vezes para escolher uma imagem' },
      ];
    case 'quote':
      return [
        text('quote', 220, 260, 1480, 420, '', { text: { ...text('quote').text, align: 'center', valign: 'middle', font: 'head' }, ph: '“Uma frase marcante vai aqui.”' }),
        text('caption', 460, 720, 1000, 80, '', { text: { ...text('caption').text, align: 'center' }, ph: '— Autor da citação' }),
      ];
    case 'bigNumber':
      return [
        text('big', 160, 220, 1600, 360, '', { text: { ...text('big').text, align: 'center', valign: 'middle', color: 'accent' }, ph: '87%' }),
        text('subtitle', 260, 620, 1400, 160, '', { text: { ...text('subtitle').text, align: 'center' }, ph: 'O que esse número significa' }),
      ];
    case 'titleOnly':
      return [text('title', 120, 90, 1680, 150, '', { ph: 'Título do slide' })];
    default:
      return [];
  }
}

export function newSlide(layout = 'titleContent') {
  return { id: uid('s'), layout, bg: null, elements: layoutElements(layout), notes: '', transition: 'fade' };
}

export function newPresentation(themeId = 'claro') {
  return { title: 'Apresentação sem título', theme: themeId, slides: [newSlide('title')], assets: {} };
}

// ---- Modelo: pitch de startup ---------------------------------------------------------

function fill(slide, values) {
  slide.elements.forEach((e, i) => {
    const v = values[i];
    if (v === undefined || !e.text) return;
    e.text.html = Array.isArray(v) ? `<ul>${v.map((x) => `<li>${x}</li>`).join('')}</ul>` : `<p>${v}</p>`;
  });
  return slide;
}

export function pitchPresentation() {
  const p = newPresentation('pitch');
  p.title = 'Pitch de startup';
  p.slides = [
    fill(newSlide('title'), ['NomeDaStartup', 'Resolvendo um grande problema com tecnologia · 2026']),
    fill(newSlide('titleContent'), ['O problema', ['Pequenas empresas perdem 12 horas por semana com tarefas manuais', 'Ferramentas atuais são caras e difíceis de usar', '68% dos donos dizem não ter tempo para crescer']]),
    fill(newSlide('titleContent'), ['Nossa solução', ['Plataforma simples que automatiza o dia a dia', 'Configuração em 5 minutos, sem código', 'Integra com WhatsApp, e-mail e planilhas']]),
    fill(newSlide('bigNumber'), ['R$ 48 bi', 'Mercado endereçável no Brasil (SAM), crescendo 21% ao ano']),
    fill(newSlide('twoCols'), ['Modelo de negócio', ['Assinatura mensal a partir de R$ 89', 'Planos por número de usuários'], ['CAC: R$ 210', 'LTV: R$ 3.400', 'Payback: 4 meses']]),
    fill(newSlide('titleContent'), ['Tração', ['1.200 clientes pagantes', 'Crescimento de 18% ao mês nos últimos 6 meses', 'NPS 71']]),
    fill(newSlide('twoCols'), ['Equipe', ['<b>Ana Souza</b> — CEO, ex-líder de produto', '<b>Bruno Lima</b> — CTO, 12 anos em SaaS'], ['<b>Carla Mendes</b> — Vendas, ex-diretora comercial', '<b>Diego Alves</b> — Design']]),
    fill(newSlide('section'), ['Vamos conversar?', 'Captando R$ 5 milhões · contato@startup.com.br']),
  ];
  return p;
}

export function themedPresentation(themeId) {
  if (themeId === 'pitch') return pitchPresentation();
  const p = newPresentation(THEMES[themeId] ? themeId : 'claro');
  return p;
}
