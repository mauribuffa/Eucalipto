const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');
const os = require('os');

(async () => {
  const browser = await puppeteer.launch({
    headless: 'new',
    executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  });
  const page = await browser.newPage();

  // Set viewport to 1920x1080 (presentation aspect ratio)
  await page.setViewport({ width: 1920, height: 1080 });

  const filePath = path.resolve(__dirname, 'index.html');
  await page.goto(`file://${filePath}`, { waitUntil: 'networkidle0', timeout: 30000 });

  // Wait for fonts to load
  await new Promise(r => setTimeout(r, 2000));

  // Disable ALL animations and transitions so every element is in its final state
  await page.evaluate(() => {
    const style = document.createElement('style');
    style.textContent = `
      *, *::before, *::after {
        animation: none !important;
        animation-delay: 0s !important;
        transition: none !important;
        opacity: 1 !important;
      }
      .slide .card,
      .slide .stat-card,
      .slide .table-wrapper,
      .slide .objective-list li,
      .slide .flow-step,
      .slide .highlight-box {
        opacity: 1 !important;
        transform: none !important;
      }
      svg animate { }
    `;
    document.head.appendChild(style);

    // Force all SVG animations to their end state
    document.querySelectorAll('svg animate').forEach(anim => {
      try { anim.endElement(); } catch(e) {}
    });

    // Kill GSAP
    if (window.gsap) {
      gsap.globalTimeline.clear();
      gsap.set(document.querySelectorAll('*'), { clearProps: 'all' });
    }
  });

  await new Promise(r => setTimeout(r, 500));

  // Get total number of slides
  const totalSlides = await page.evaluate(() => {
    return document.querySelectorAll('.slide').length;
  });

  console.log(`Found ${totalSlides} slides`);

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'slides-'));
  console.log(`Temp dir: ${tmpDir}`);

  for (let i = 0; i < totalSlides; i++) {
    // Activate the current slide
    await page.evaluate((index) => {
      const slides = document.querySelectorAll('.slide');
      slides.forEach((s, idx) => {
        s.classList.remove('active', 'exit');
        if (idx === index) {
          s.classList.add('active');
          s.style.opacity = '1';
          s.style.visibility = 'visible';
          s.style.transform = 'none';
          s.style.position = 'absolute';
          s.style.inset = '0';
        } else {
          s.style.opacity = '0';
          s.style.visibility = 'hidden';
        }
      });
      // Ensure all children inside active slide are visible
      const active = slides[index];
      active.querySelectorAll('*').forEach(el => {
        const cs = getComputedStyle(el);
        if (cs.opacity === '0') el.style.opacity = '1';
      });
    }, i);

    await new Promise(r => setTimeout(r, 200));

    const imgPath = path.join(tmpDir, `slide-${String(i).padStart(3, '0')}.png`);
    await page.screenshot({ path: imgPath, type: 'png' });
    console.log(`Captured slide ${i + 1}/${totalSlides}`);
  }

  // Build PDF from saved images using file:// URLs
  const page2 = await browser.newPage();
  await page2.setViewport({ width: 1920, height: 1080 });

  const imagesHtml = Array.from({ length: totalSlides }, (_, i) => {
    const imgPath = path.join(tmpDir, `slide-${String(i).padStart(3, '0')}.png`);
    return `<div class="page"><img src="file://${imgPath}" /></div>`;
  }).join('\n');

  await page2.goto(`data:text/html,`, { waitUntil: 'domcontentloaded' });
  await page2.setContent(`
    <html>
    <style>
      * { margin: 0; padding: 0; }
      @page { size: 1920px 1080px; margin: 0; }
      .page { width: 1920px; height: 1080px; page-break-after: always; overflow: hidden; }
      .page:last-child { page-break-after: auto; }
      img { width: 100%; height: 100%; display: block; object-fit: contain; }
    </style>
    <body>${imagesHtml}</body>
    </html>
  `, { waitUntil: 'load', timeout: 120000 });

  await page2.pdf({
    path: path.resolve(__dirname, 'presentacion-eucalipto.pdf'),
    width: '1920px',
    height: '1080px',
    printBackground: true,
    margin: { top: 0, right: 0, bottom: 0, left: 0 },
  });

  console.log('PDF saved as presentacion-eucalipto.pdf');
  await browser.close();

  // Clean up temp files
  fs.rmSync(tmpDir, { recursive: true, force: true });
  console.log('Cleaned up temp files');
})();
