// Gera dist/junta-planilhas.html: um único arquivo com o SheetJS embutido (funciona offline).
const fs = require('fs');
const path = require('path');

const ler = (p) => fs.readFileSync(path.join(__dirname, p), 'utf8');
// "</script" dentro de um <script> encerraria a tag antes da hora
const seguro = (codigo) => codigo.replace(/<\/script/gi, '<\\/script');

const partes = {
  '/*__SHEETJS__*/': seguro(ler('node_modules/xlsx/dist/xlsx.full.min.js')),
  '/*__NUCLEO__*/': seguro(ler('src/nucleo.js')),
};

let html = ler('src/index.template.html');
for (const [marcador, codigo] of Object.entries(partes)) {
  if (!html.includes(marcador)) throw new Error(`Marcador ${marcador} não encontrado no modelo.`);
  html = html.split(marcador).join(codigo);
}

fs.mkdirSync(path.join(__dirname, 'dist'), { recursive: true });
const destino = path.join(__dirname, 'dist', 'junta-planilhas.html');
fs.writeFileSync(destino, html);
console.log(`Gerado ${path.relative(__dirname, destino)} (${(Buffer.byteLength(html) / 1048576).toFixed(2)} MB)`);
