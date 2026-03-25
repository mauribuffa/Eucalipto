let current = 1;
const total = document.querySelectorAll('.slide').length;
document.getElementById('totalSlides').textContent = total;

function goToSlide(n) {
  if (n < 1 || n > total) return;
  const slides = document.querySelectorAll('.slide');
  slides.forEach(s => {
    s.classList.remove('active', 'exit');
  });
  const prev = slides[current - 1];
  const next = slides[n - 1];
  if (n > current) {
    prev.classList.add('exit');
  }
  next.classList.add('active');
  current = n;
  document.getElementById('currentSlide').textContent = current;
  document.getElementById('navProgress').style.width = ((current / total) * 100) + '%';
}

function nextSlide() { goToSlide(current + 1); }
function prevSlide() { goToSlide(current - 1); }

document.addEventListener('keydown', (e) => {
  if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'ArrowDown') {
    e.preventDefault();
    nextSlide();
  } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
    e.preventDefault();
    prevSlide();
  } else if (e.key === 'Home') {
    e.preventDefault();
    goToSlide(1);
  } else if (e.key === 'End') {
    e.preventDefault();
    goToSlide(total);
  }
});

// Touch support
let touchStartX = 0;
document.addEventListener('touchstart', (e) => { touchStartX = e.touches[0].clientX; });
document.addEventListener('touchend', (e) => {
  const dx = e.changedTouches[0].clientX - touchStartX;
  if (Math.abs(dx) > 50) {
    dx < 0 ? nextSlide() : prevSlide();
  }
});

// Initial progress
document.getElementById('navProgress').style.width = ((1 / total) * 100) + '%';

// Auto-hide nav bar
const navBar = document.querySelector('.nav-bar');
let navTimeout;
function showNav() {
  navBar.classList.remove('hidden');
  clearTimeout(navTimeout);
  navTimeout = setTimeout(() => navBar.classList.add('hidden'), 3000);
}
document.addEventListener('mousemove', showNav);
document.addEventListener('keydown', showNav);
document.addEventListener('touchstart', showNav);
showNav();

// ═══════════════════ GSAP ANIMATIONS ═══════════════════

// --- Slide 1: Floating leaf + entrance animations ---
function animateSlide1() {
  const leaf = document.querySelector('.portada-leaf');
  const title = document.querySelector('.portada-title');
  const sub = document.querySelector('.portada-sub');
  const meta = document.querySelector('.portada-meta');
  const dividers = document.querySelectorAll('.portada .portada-divider');

  // Entrance timeline
  const tl = gsap.timeline();
  tl.from(leaf, { y: -60, opacity: 0, scale: 0.3, duration: 1, ease: 'back.out(1.7)' })
    .from(title, { y: 40, opacity: 0, duration: 0.8, ease: 'power3.out' }, '-=0.3')
    .from(sub, { y: 30, opacity: 0, duration: 0.6, ease: 'power2.out' }, '-=0.3')
    .from(dividers, { scaleX: 0, duration: 0.5, ease: 'power2.inOut' }, '-=0.2')
    .from(meta, { y: 20, opacity: 0, duration: 0.6, ease: 'power2.out' }, '-=0.2');

  // Continuous floating leaf
  gsap.to(leaf, {
    y: -12,
    rotation: 5,
    duration: 2.5,
    ease: 'sine.inOut',
    yoyo: true,
    repeat: -1
  });
}

// --- Slide 14: Eucalyptus tree grow + content stagger ---
function animateSlide14() {
  const tl = gsap.timeline();

  // List items stagger in
  tl.from('.slide14-list li', {
    x: -30, opacity: 0, duration: 0.4, stagger: 0.12, ease: 'power2.out'
  })
  // Table rows stagger in
  .from('.slide14-table tr', {
    x: 30, opacity: 0, duration: 0.4, stagger: 0.1, ease: 'power2.out'
  }, '-=0.3');

  // Tree grow animation
  const treeTl = gsap.timeline();

  // Trunk grows up
  treeTl.from('#tree-trunk', {
    scaleY: 0, transformOrigin: 'bottom center', duration: 0.8, ease: 'power2.out'
  })
  // Branches extend
  .from('#tree-branches path', {
    strokeDasharray: 100, strokeDashoffset: 100, duration: 0.5, stagger: 0.08, ease: 'power2.out'
  }, '-=0.2')
  // Canopy blooms from bottom to top
  .from('.canopy-layer', {
    scale: 0, transformOrigin: 'center center', opacity: 0, duration: 0.4, stagger: 0.1, ease: 'back.out(1.4)'
  }, '-=0.3');

  // Falling leaves (continuous loop after tree grows)
  treeTl.add(() => {
    document.querySelectorAll('.leaf-particle').forEach((leaf, i) => {
      gsap.to(leaf, {
        opacity: 0.7,
        y: '+=' + (60 + Math.random() * 40),
        x: '+=' + (Math.random() * 40 - 20),
        rotation: Math.random() * 360,
        duration: 2.5 + Math.random() * 1.5,
        ease: 'sine.inOut',
        repeat: -1,
        delay: i * 0.8,
        yoyo: false,
        onRepeat: function() {
          gsap.set(leaf, { opacity: 0 });
          gsap.to(leaf, { opacity: 0.7, duration: 0.3 });
        }
      });
    });
  });

  // Gentle canopy sway (continuous)
  gsap.to('.canopy-layer', {
    x: '+=' + 3,
    duration: 2,
    ease: 'sine.inOut',
    yoyo: true,
    repeat: -1,
    stagger: 0.2
  });
}

// --- Slide 17: Donut composition chart animation ---
function animateSlide17() {
  // Donut segments grow in
  const segments = document.querySelectorAll('.donut-seg');
  segments.forEach((seg, i) => {
    const finalDash = seg.getAttribute('stroke-dasharray');
    const finalOffset = seg.getAttribute('stroke-dashoffset');
    seg.setAttribute('stroke-dasharray', '0 502.65');
    gsap.to(seg, {
      attr: { 'stroke-dasharray': finalDash },
      duration: 0.8,
      delay: 0.2 + i * 0.15,
      ease: 'power2.out'
    });
  });

  // Center text fade in
  gsap.from('#composition-chart circle:last-of-type', {
    scale: 0, transformOrigin: '130px 125px', duration: 0.6, ease: 'back.out(1.4)'
  });
  gsap.from('#composition-chart text', {
    opacity: 0, duration: 0.5, delay: 0.6, stagger: 0.05
  });
}

// Track which slides have been animated
const animatedSlides = {};

// Hook into slide navigation
const originalGoToSlide = goToSlide;
goToSlide = function(n) {
  originalGoToSlide(n);
  const slides = document.querySelectorAll('.slide');
  const slideEl = slides[n - 1];
  if (!animatedSlides[n]) {
    animatedSlides[n] = true;
    // Title entrance
    const title = slideEl?.querySelector('.slide-title, .sep-title');
    if (title) gsap.from(title, { y: 30, opacity: 0, duration: 0.6, ease: 'power3.out', delay: 0.1 });
    // Section tag
    const tag = slideEl?.querySelector('.section-tag');
    if (tag) gsap.from(tag, { x: -20, opacity: 0, duration: 0.4, ease: 'power2.out' });
    // Separator icon animation
    if (slideEl?.classList.contains('slide--separator')) {
      const icon = slideEl.querySelector('.sep-icon');
      if (icon) gsap.from(icon, { scale: 0, rotation: -180, duration: 0.8, ease: 'back.out(1.7)' });
      const line = slideEl.querySelector('.sep-line');
      if (line) gsap.from(line, { scaleX: 0, duration: 0.5, delay: 0.4, ease: 'power2.inOut' });
    }
    // Specific slide animations
    if (n === 1) animateSlide1();
    const dataSlide = slideEl?.getAttribute('data-slide');
    if (dataSlide === '14') animateSlide14();
    if (dataSlide === '17') animateSlide17();
    if (dataSlide === '9') animateChartDemanda();
    if (dataSlide === '10') animateChartOferta();
    if (dataSlide === '38') animateDonutCharts(slideEl);
    if (dataSlide === '42') animateBreakeven();
    if (dataSlide === '43') animateChartTIR();
    if (dataSlide === '33') animateBalance();
  }
};

// Trigger slide 1 animation on load
animateSlide1();
animatedSlides[1] = true;

// ═══════════════════ CHART ANIMATIONS ═══════════════════

// Slide 9: Bar chart — bars grow up from baseline
function animateChartDemanda() {
  const svg = document.getElementById('chart-demanda');
  if (!svg) return;
  svg.querySelectorAll('.demand-bar').forEach((bar, i) => {
    const targetY = parseFloat(bar.getAttribute('data-target-y'));
    const targetH = parseFloat(bar.getAttribute('data-target-h'));
    gsap.to(bar, {
      attr: { y: targetY, height: targetH },
      duration: 0.7,
      delay: 0.15 + i * 0.07,
      ease: 'power2.out'
    });
  });
  gsap.to(svg.querySelectorAll('.demand-label'), {
    opacity: 1, duration: 0.4, delay: 1, stagger: 0.08
  });
}

// Slide 10: Line chart — lines draw in, dots pop, areas fade
function animateChartOferta() {
  const svg = document.getElementById('chart-oferta');
  if (!svg) return;
  // Animate polylines drawing in
  svg.querySelectorAll('.chart-line').forEach((line, i) => {
    const len = line.getTotalLength();
    line.style.strokeDasharray = len;
    line.style.strokeDashoffset = len;
    gsap.to(line, { strokeDashoffset: 0, duration: 1.2, delay: 0.2 + i * 0.3, ease: 'power2.inOut' });
  });
  // Areas fade in
  gsap.from(svg.querySelectorAll('.chart-area'), { opacity: 0, duration: 0.8, delay: 0.6, stagger: 0.2 });
  // Dots scale in
  gsap.from(svg.querySelectorAll('.chart-dot'), { scale: 0, transformOrigin: '50% 50%', duration: 0.3, delay: 0.8, stagger: 0.08, ease: 'back.out(2)' });
  // Labels fade in
  gsap.from(svg.querySelectorAll('.chart-label'), { opacity: 0, y: 5, duration: 0.3, delay: 1.2, stagger: 0.06 });
}

// Slide 38: Donut charts — segments grow in
function animateDonutCharts(slideEl) {
  slideEl.querySelectorAll('.donut-chart').forEach(chart => {
    const circumference = 2 * Math.PI * 70; // r=70
    chart.querySelectorAll('.donut-seg').forEach((seg, i) => {
      const finalDash = seg.getAttribute('stroke-dasharray');
      seg.setAttribute('stroke-dasharray', '0 ' + circumference);
      gsap.to(seg, { attr: { 'stroke-dasharray': finalDash }, duration: 0.8, delay: 0.2 + i * 0.12, ease: 'power2.out' });
    });
    gsap.from(chart.querySelectorAll('.donut-text'), { opacity: 0, duration: 0.5, delay: 0.7, stagger: 0.1 });
  });
}

// Slide 42: Break-even chart — lines draw in, areas fade, point pulses
function animateBreakeven() {
  const svg = document.getElementById('chart-breakeven');
  if (!svg) return;
  // Lines draw in
  svg.querySelectorAll('.be-line').forEach((line, i) => {
    const len = line.getTotalLength ? line.getTotalLength() : 400;
    line.style.strokeDasharray = len;
    line.style.strokeDashoffset = len;
    gsap.to(line, { strokeDashoffset: 0, duration: 1, delay: 0.2 + i * 0.2, ease: 'power2.inOut' });
  });
  // Areas fade in
  gsap.from(svg.querySelectorAll('.be-area'), { opacity: 0, duration: 0.6, delay: 1, stagger: 0.15 });
  // Break-even point and label
  const circles = svg.querySelectorAll('circle');
  gsap.from(circles, { scale: 0, transformOrigin: '50% 50%', duration: 0.5, delay: 1.2, ease: 'back.out(2)' });
  // Zone labels
  const zoneTexts = svg.querySelectorAll('text');
  const lastTexts = Array.from(zoneTexts).slice(-4);
  gsap.from(lastTexts, { opacity: 0, duration: 0.4, delay: 1.4, stagger: 0.1 });
}

// Slide 43: TIR bar chart — bars grow, labels fade in
function animateChartTIR() {
  const svg = document.getElementById('chart-tir');
  if (!svg) return;
  svg.querySelectorAll('.tir-bar').forEach((bar, i) => {
    const targetWidth = bar.getAttribute('data-width');
    gsap.to(bar, { attr: { width: targetWidth }, duration: 0.7, delay: 0.2 + i * 0.15, ease: 'power2.out' });
  });
  gsap.to(svg.querySelectorAll('.tir-label'), { opacity: 1, duration: 0.4, delay: 0.6, stagger: 0.12 });
}

// Slide 33: Balance de masa — inputs → process → outputs flow
function animateBalance() {
  const svg = document.getElementById('chart-balance');
  if (!svg) return;
  const groups = svg.querySelectorAll('g');
  // groups: 0=inputs, 1=left arrows, 2=center block, 3=right arrows, 4=outputs
  if (groups.length >= 5) {
    gsap.from(groups[0].children, { x: -40, opacity: 0, duration: 0.5, stagger: 0.1, ease: 'power2.out' });
    gsap.from(groups[1].children, { opacity: 0, scale: 0, transformOrigin: '50% 50%', duration: 0.3, delay: 0.5, stagger: 0.08, ease: 'back.out(1.5)' });
    gsap.from(groups[2].children, { scale: 0.8, opacity: 0, transformOrigin: '450px 160px', duration: 0.6, delay: 0.8, ease: 'back.out(1.4)' });
    gsap.from(groups[3].children, { opacity: 0, scale: 0, transformOrigin: '50% 50%', duration: 0.3, delay: 1.2, stagger: 0.08, ease: 'back.out(1.5)' });
    gsap.from(groups[4].children, { x: 40, opacity: 0, duration: 0.5, delay: 1.4, stagger: 0.1, ease: 'power2.out' });
  }
}

// ═══════════════════ ANIMATED COUNTERS ═══════════════════
</script>
