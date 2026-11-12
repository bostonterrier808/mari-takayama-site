// Upcoming performances, shared by the Home page and the Credits page.
// One entry per production. shows: every show as 'YYYY-MM-DD HH:MM' (24-hour). tickets: optional link.
// A production shows under "Upcoming" until its last show day; from the next day it moves to Credits.
const MARI_SHOWS = [
  { work: 'A Southern Nutcracker', role: 'Sugar Plum Fairy', company: 'Southern Arts Alliance', place: 'Biloxi, MS',
    shows: ['2026-11-28 14:00', '2026-11-28 19:00'] },
  { work: 'The Nutcracker', role: 'Sugar Plum Fairy', company: 'Eastern Shore Ballet Theatre', place: 'Salisbury, MD',
    shows: ['2026-12-04 19:30', '2026-12-05 11:00', '2026-12-05 19:00', '2026-12-06 14:00'],
    tickets: 'https://www.etix.com/ticket/v/23565/wicomico-high-school-eastern-shore-ballet-theatre' },
  { work: 'The Nutcracker', role: 'Sugar Plum Fairy', company: 'Jill M. Taylor Dance Company', place: 'Rolesville, NC',
    shows: ['2026-12-12 14:00', '2026-12-12 18:00'],
    tickets: 'https://www.etix.com/ticket/v/29119/rolesville-high-school-auditorium' },
  { work: 'The Nutcracker', role: 'Sugar Plum Fairy', company: 'Dance Wisconsin', place: 'Madison, WI',
    shows: ['2026-12-19 14:00', '2026-12-20 14:00'] }
];

(() => {
  const ja = () => document.documentElement.lang === 'ja';
  const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const day = k => { const [y, m, d] = k.split('-').map(Number); return new Date(y, m - 1, d); };
  const when = (a, b) => {
    const x = day(a), y = day(b);
    if (ja()) return (x.getMonth() + 1) + '/' + x.getDate() + (a === b ? '' : '–' + (x.getMonth() === y.getMonth() ? '' : (y.getMonth() + 1) + '/') + y.getDate());
    return MON[x.getMonth()] + ' ' + x.getDate() + (a === b ? '' : '–' + (x.getMonth() === y.getMonth() ? '' : MON[y.getMonth()] + ' ') + y.getDate());
  };
  const clock = t => { const [h, m] = t.split(':').map(Number); return (h % 12 || 12) + (m ? ':' + String(m).padStart(2, '0') : '') + (h < 12 ? 'am' : 'pm'); };
  const times = list => {
    const byDay = []; list.forEach(([k, t]) => { const g = byDay.find(x => x.k === k); if (g) g.t.push(t); else byDay.push({ k, t: [t] }); });
    const one = byDay.length === 1;
    return byDay.map(g => (one ? '' : when(g.k, g.k) + (ja() ? ' ' : ', ')) + g.t.map(clock).join(ja() ? '・' : ' & ')).join(ja() ? '　' : ' · ');
  };
  const workName = w => !ja() ? w : w === 'The Nutcracker' ? '『くるみ割り人形』' : '『' + w + '』';
  const roleName = r => !ja() ? r : r === 'Sugar Plum Fairy' ? '金平糖の精' : r;
  const span = (parent, cls, text) => { const e = document.createElement('span'); e.className = cls; e.textContent = text; parent.append(e); return e; };
  const prods = () => MARI_SHOWS.map(p => { const s = p.shows.map(x => x.split(' ')); return { ...p, s, first: s[0][0], last: s[s.length - 1][0] }; });
  const today = () => { const t = new Date(); t.setHours(0, 0, 0, 0); return t; };

  function renderUpcoming() {
    document.querySelectorAll('.upcoming-list').forEach(list => {
      list.textContent = '';
      const now = prods().filter(p => day(p.last) >= today());
      now.forEach(p => {
        const li = document.createElement('li'); li.setAttribute('translate', 'no');
        span(li, 'up-date', when(p.first, p.last));
        span(li, 'up-work', workName(p.work) + (ja() ? '　' + roleName(p.role) + '（ゲスト出演）' : ' — ' + p.role + ' (guest artist)'));
        span(li, 'up-meta up-times', times(p.s));
        const meta = span(li, 'up-meta', p.company + ' · ' + p.place);
        if (p.tickets) { const a = document.createElement('a'); a.href = p.tickets; a.target = '_blank'; a.rel = 'noopener'; a.textContent = ja() ? 'チケット' : 'Tickets'; meta.append(a); }
        list.append(li);
      });
      list.hidden = !now.length;
      const empty = list.parentElement.querySelector('.upcoming-empty'); if (empty) empty.hidden = !!now.length;
    });
  }
  // Credits: finished productions, newest first, added to the top of the Ballet list
  function renderPast() {
    const list = document.querySelector('.credit-list.ballet-list'); if (!list) return;
    list.querySelectorAll('li.from-shows').forEach(li => li.remove());
    const done = prods().filter(p => day(p.last) < today()).sort((a, b) => day(b.last) - day(a.last));
    done.slice().reverse().forEach(p => {
      const li = document.createElement('li'); li.className = 'from-shows'; li.setAttribute('translate', 'no');
      const d = day(p.last);
      span(li, 'credit-work', workName(p.work));
      span(li, 'credit-role', roleName(p.role));
      span(li, 'credit-note', ja() ? 'ゲスト出演・' + p.company + '・' + p.place + '（' + d.getFullYear() + '年' + (d.getMonth() + 1) + '月）' : 'Guest artist, ' + p.company + ', ' + p.place + ' (' + MON[d.getMonth()] + ' ' + d.getFullYear() + ')');
      list.prepend(li);
    });
  }
  const render = () => { renderUpcoming(); renderPast(); };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', render); else render();
  document.addEventListener('mari-language-change', render);
})();
