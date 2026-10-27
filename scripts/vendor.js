// Copia as bibliotecas de navegador de node_modules para src/vendor (executado no postinstall).
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const out = path.join(root, 'src', 'vendor');
const FILES = {
  'exceljs.min.js': 'exceljs/dist/exceljs.min.js',
  'mammoth.browser.min.js': 'mammoth/mammoth.browser.min.js',
  'docx.iife.js': 'docx/dist/index.iife.js',
  'pptxgen.bundle.js': 'pptxgenjs/dist/pptxgen.bundle.js',
  'chart.umd.min.js': 'chart.js/dist/chart.umd.min.js',
  'purify.min.js': 'dompurify/dist/purify.min.js',
};

fs.mkdirSync(out, { recursive: true });
let missing = 0;
for (const [dest, src] of Object.entries(FILES)) {
  const from = path.join(root, 'node_modules', src);
  if (!fs.existsSync(from)) {
    console.warn(`[vendor] não encontrado: ${src}`);
    missing++;
    continue;
  }
  fs.copyFileSync(from, path.join(out, dest));
}
console.log(`[vendor] ${Object.keys(FILES).length - missing} bibliotecas copiadas para src/vendor`);
