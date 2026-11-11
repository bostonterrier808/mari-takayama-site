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
    const drift = [...document.querySelectorAll('.home-duo .home-path > img, .home-intro-photo')];
    const onScroll2 = () => drift.forEach((el, i) => { const r = el.getBoundingClientRect(); if (r.bottom < -100 || r.top > innerHeight + 100) return; const p = (r.top + r.height / 2 - innerHeight / 2) / innerHeight; el.style.translate = `0 ${(-p * (i % 2 ? 36 : 24)).toFixed(1)}px`; });
    addEventListener('scroll', () => { if (!ticking) { ticking = true; requestAnimationFrame(() => { onScroll(); onScroll2(); }); } }, { passive: true });
    onScroll(); onScroll2();
    // gallery strip drifts slowly on its own; stops while you touch, drag or hover it
    const strip = document.querySelector('.home-strip');
    if (strip) {
      let hold = false, dir = 1, visible = false, last = 0;
      ['pointerenter', 'pointerdown', 'touchstart', 'focusin'].forEach(t => strip.addEventListener(t, () => { hold = true; }, { passive: true }));
      ['pointerleave', 'touchend', 'focusout'].forEach(t => strip.addEventListener(t, () => { setTimeout(() => { hold = false; }, 1200); }, { passive: true }));
      new IntersectionObserver(([e]) => { visible = e.isIntersecting; }).observe(strip);
      const step = t => { const dt = Math.min(48, t - last || 16); last = t; if (visible && !hold) { const max = strip.scrollWidth - strip.clientWidth; if (max > 0) { strip.scrollLeft += dir * dt * 0.03; if (strip.scrollLeft >= max - 1) dir = -1; else if (strip.scrollLeft <= 0) dir = 1; } } requestAnimationFrame(step); };
      strip.style.scrollSnapType = 'none';
      requestAnimationFrame(step);
    }
  }
})();

// ---------- Stronger movement (only on the real site, not inside the editor preview) ----------
(() => {
  if (matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  const inFrame = window.self !== window.top;
  const root = document.documentElement;
  const fine = matchMedia('(pointer: fine)').matches;

  if (!inFrame) {
    // 2) soft page change: fade out, then go
    root.classList.add('mt-page-in');
    document.addEventListener('click', e => {
      if (e.defaultPrevented || e.button || e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return;
      const a = e.target.closest('a[href]'); if (!a || a.target === '_blank' || a.hasAttribute('download')) return;
      const url = new URL(a.href, location.href);
      if (url.origin !== location.origin || !/\.html$|\/$/.test(url.pathname) || (url.pathname === location.pathname && url.hash)) return;
      e.preventDefault(); root.classList.add('mt-page-out'); setTimeout(() => { location.href = url.href; }, 420);
    });
    addEventListener('pageshow', ev => { if (ev.persisted) root.classList.remove('mt-page-out'); });
  }

  // 3) header slips away when scrolling down, returns when scrolling up
  const header = document.querySelector('header.top');
  let lastY = scrollY;
  addEventListener('scroll', () => {
    const y = scrollY; if (!header || header.classList.contains('menu-open')) { lastY = y; return; }
    header.classList.toggle('mt-hide', y > lastY && y > 240); lastY = y;
  }, { passive: true });

  if (!fine || inFrame) return;
  // 5) a small gold ring follows the mouse and opens up over links and photos
  const ring = document.createElement('div'); ring.className = 'mt-cursor'; document.body.append(ring);
  let x = innerWidth / 2, y = innerHeight / 2, rx = x, ry = y;
  addEventListener('mousemove', e => { x = e.clientX; y = e.clientY; ring.classList.add('on'); }, { passive: true });
  document.addEventListener('mouseleave', () => ring.classList.remove('on'));
  document.addEventListener('mouseover', e => {
    const t = e.target;
    ring.classList.toggle('link', !!t.closest('a, button'));
    ring.classList.toggle('photo', !!t.closest('main img, #grid button, .home-strip a'));
  });
  const loop = () => { rx += (x - rx) * 0.18; ry += (y - ry) * 0.18; ring.style.transform = `translate(${rx}px, ${ry}px)`; requestAnimationFrame(loop); };
  requestAnimationFrame(loop);
  // 7) photos tilt slightly toward the mouse
  document.querySelectorAll('.home-duo .home-path > img').forEach(el => {
    el.style.transition = (el.style.transition ? el.style.transition + ', ' : '') + 'transform .6s cubic-bezier(.2,.7,.2,1)';
    el.addEventListener('mousemove', e => { const r = el.getBoundingClientRect(); const px = (e.clientX - r.left) / r.width - 0.5, py = (e.clientY - r.top) / r.height - 0.5; el.style.transform = `perspective(900px) rotateY(${px * 5}deg) rotateX(${-py * 5}deg) scale(1.02)`; });
    el.addEventListener('mouseleave', () => { el.style.transform = ''; });
  });
})();
