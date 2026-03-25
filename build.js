const fs = require('fs');
const path = require('path');

const SLIDES_DIR = path.join(__dirname, 'slides');
const CSS_FILE = path.join(__dirname, 'css', 'styles.css');
const JS_FILE = path.join(__dirname, 'js', 'main.js');
const OUTPUT = path.join(__dirname, 'index.html');

function build() {
  // Read all slide files in order (00-icons, 01-40 slides, 99-nav)
  const slideFiles = fs.readdirSync(SLIDES_DIR)
    .filter(f => f.endsWith('.html'))
    .sort();

  const icons = fs.readFileSync(path.join(SLIDES_DIR, slideFiles.find(f => f.startsWith('00-'))), 'utf8');
  const nav = fs.readFileSync(path.join(SLIDES_DIR, slideFiles.find(f => f.startsWith('99-'))), 'utf8');
  const slides = slideFiles
    .filter(f => !f.startsWith('00-') && !f.startsWith('99-'))
    .map(f => fs.readFileSync(path.join(SLIDES_DIR, f), 'utf8'))
    .join('\n\n');

  const css = fs.readFileSync(CSS_FILE, 'utf8');
  const js = fs.readFileSync(JS_FILE, 'utf8');

  const html = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Repelente Natural de Eucalipto — Proyecto Final</title>
<script src="https://cdn.jsdelivr.net/npm/gsap@3.12.5/dist/gsap.min.js"><\/script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,600;0,700;1,400;1,600&family=Outfit:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
${css}</style>
</head>
<body>

<!-- SVG ICON LIBRARY -->
${icons}

<div class="presentation" id="presentation">

${slides}

</div><!-- /presentation -->

${nav}

<script>
${js}</script>

</body>
</html>`;

  fs.writeFileSync(OUTPUT, html);
  console.log(`[build] index.html generated (${html.split('\n').length} lines)`);
}

build();

module.exports = build;
