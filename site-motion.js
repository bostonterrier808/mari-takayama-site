// Gentle motion for every page: parts fade up, photos unveil, the top photo drifts on scroll.
(() => {
  const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
  if (reduce) return;
  if (!('IntersectionObserver' in window)) return;
  // 2) parts fade up as they come into view (Home has its own list)
  const isHome = !!document.getElementById('home');
  const selectors = isHome ? [] : [
    'main .about .wrap > div > *', 'main .row > div:not(.row-photo) > *', 'main .lesson-places > *', 'main .travel-lesson',
    'main .published-interview', 'main .press-language-title', 'main .press-language-nav',
    'main .credits-lead', 'main .credits-page h2', 'main .credit-list li', 'main .credits-empty',
    'main .booking-card-text > *', 'main .booking-grid > *', 'main .booking-steps .wrap > *', 'main .booking-share .wrap > *',
    'main .rates-page h2', 'main .rates-table', 'main form > *'
  ];
  const items = [];
  selectors.forEach(sel => document.querySelectorAll(sel).forEach((el, i) => {
    if (el.classList.contains('rv')) return;
    el.classList.add('rv'); el.style.setProperty('--rv-delay', (i % 6) * 0.08 + 's'); items.push(el);
  }));
  // 3) photos unveil like a curtain
  const photos = isHome ? [] : document.querySelectorAll('main .about img, main .row-photo img, main #grid > button, main .booking-photo');
  // a photo is watched through its parent box, so the clip on the photo never hides it from the watcher
  const watch = new Map();
  photos.forEach((el, i) => { el.classList.add('rv-img'); el.style.transitionDelay = (i % 3) * 0.12 + 's'; const box = el.tagName === 'IMG' ? el.parentElement : el; watch.set(box, [...(watch.get(box) || []), el]); });
  items.forEach(el => watch.set(el, [...(watch.get(el) || []), el]));
  if (watch.size) {
    document.documentElement.classList.add('site-reveal');
    const io = new IntersectionObserver(entries => entries.forEach(e => { if (e.isIntersecting) { (watch.get(e.target) || []).forEach(x => x.classList.add('in')); io.unobserve(e.target); } }), { rootMargin: '0px 0px -6% 0px', threshold: 0.01 });
    watch.forEach((_, box) => io.observe(box));
  }
  // 4) Home: the top photo drifts slower than the page, the outlined "Mari" glides sideways
  if (isHome) {
    const hero = document.querySelector('.hero-photo'), script = document.querySelector('.intro-script');
    let ticking = false;
    const onScroll = () => {
      ticking = false; const y = scrollY;
      if (hero && y < innerHeight * 1.2) hero.style.transform = `translateY(${y * 0.18}px)`;
      if (script) { const r = script.parentElement.getBoundingClientRect(); const p = (innerHeight - r.top) / (innerHeight + r.height); if (p > -0.2 && p < 1.2) script.style.translate = `${(0.5 - p) * 120}px 0`; }
    };
    addEventListener('scroll', () => { if (!ticking) { ticking = true; requestAnimationFrame(onScroll); } }, { passive: true });
    onScroll();
  }
})();
