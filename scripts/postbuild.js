const fs = require('fs');
const path = require('path');

const repoRoot = path.join(__dirname, '..');
const distDir = path.join(repoRoot, 'dist');
const publicLib = path.join(repoRoot, 'public', 'libs', 'pdfobject.min.js');
const outPdfObject = path.join(distDir, 'pdfobject.min.js');

const cdnUrl = 'https://cdnjs.cloudflare.com/ajax/libs/pdfobject/2.1.1/pdfobject.min.js';
const integrityRegex = / integrity="sha512-[^"]+" crossorigin="anonymous"/g;

function walk(dir) {
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const e of entries) {
    const full = path.join(dir, e.name);
    if (e.isDirectory()) walk(full);
    else if (/\.js$|\.html$/.test(e.name)) replaceInFile(full);
  }
}

function replaceInFile(filePath) {
  let content = fs.readFileSync(filePath, 'utf8');
  if (content.indexOf(cdnUrl) === -1 && !integrityRegex.test(content)) return;
  content = content.split(cdnUrl).join('./pdfobject.min.js');
  content = content.replace(integrityRegex, '');
  fs.writeFileSync(filePath, content, 'utf8');
  console.log('Patched:', filePath);
}

function copyPdfObject() {
  if (!fs.existsSync(publicLib)) {
    console.warn('pdfobject stub not found at', publicLib);
    return;
  }
  fs.copyFileSync(publicLib, outPdfObject);
  console.log('Copied pdfobject to dist:', outPdfObject);
}

if (!fs.existsSync(distDir)) {
  console.error('dist/ not found. Run build first.');
  process.exit(1);
}

walk(distDir);
copyPdfObject();
console.log('postbuild: finished.');
