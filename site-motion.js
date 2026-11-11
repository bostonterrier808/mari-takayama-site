// Gentle motion for every page except Home: parts fade up as they scroll into view.
(() => {
  if (!('IntersectionObserver' in window) || matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  const selectors = [
    'main .about .wrap > *', 'main .row > *', 'main .lesson-places > *', 'main .travel-lesson',
    'main #grid > button', 'main .published-interview', 'main .press-language-title', 'main .press-language-nav',
    'main .credits-lead', 'main .credits-page h2', 'main .credit-list li', 'main .credits-empty',
    'main .booking-card .wrap > *', 'main .booking-grid > *', 'main .booking-steps .wrap > *', 'main .booking-share .wrap > *',
    'main .rates-page h2', 'main .rates-table', 'main form > *'
  ];
  const items = [];
  selectors.forEach(sel => document.querySelectorAll(sel).forEach((el, i) => {
    if (el.classList.contains('rv')) return;
    el.classList.add('rv');
    el.style.setProperty('--rv-delay', Math.min(i % 6, 5) * 0.08 + 's');
    items.push(el);
  }));
  if (!items.length) return;
  document.documentElement.classList.add('site-reveal');
  const io = new IntersectionObserver(entries => entries.forEach(e => {
    if (e.isIntersecting) { e.target.classList.add('in'); io.unobserve(e.target); }
  }), { rootMargin: '0px 0px -6% 0px', threshold: 0.06 });
  items.forEach(el => io.observe(el));
})();
