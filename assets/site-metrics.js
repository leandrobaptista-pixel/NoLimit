/* Aggregate website measurement. No visitor IDs, full URLs, form values or IPs. */
(() => {
  if (!['nolimitcontractor.net', 'www.nolimitcontractor.net'].includes(location.hostname)) return;
  if (!['/', '/index.html', '/catalogs.html', '/viewer.html'].includes(location.pathname)) return;
  const optedOut = () => {
    try { return navigator.globalPrivacyControl === true || navigator.doNotTrack === '1' || localStorage.getItem('nolimitAnalyticsOptOut') === '1'; }
    catch { return true; }
  };
  if (optedOut()) return;
  const categories = new Set(['trim','wainscoting','stairs','crown-molding-ceiling-coffered-ceiling','decks','kitchen-vanities','fireplaces-bars','outside-doors-windows','pergola','port-portal','commercial','trash-container','wall-paneling']);
  const categoryName = value => {
    const slug=String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'');
    return ['ceiling','crown-molding','coffered-ceiling'].includes(slug) ? 'crown-molding-ceiling-coffered-ceiling' : slug;
  };
  const page = location.pathname.endsWith('catalogs.html') ? 'catalogs' : location.pathname.endsWith('viewer.html') ? 'photo' : 'home';
  const seen = new Set();
  function metric(event, dimension) {
    if (optedOut() || document.visibilityState === 'hidden') return;
    fetch('/api/site-metrics', {method:'POST',credentials:'omit',referrerPolicy:'no-referrer',headers:{'Content-Type':'application/json'},body:JSON.stringify({event,dimension}),keepalive:true}).catch(() => {});
  }
  function once(event, dimension) {
    const key = `${event}:${dimension}`;
    if (!seen.has(key)) { seen.add(key); metric(event, dimension); }
  }
  function source() {
    const medium = new URLSearchParams(location.search).get('utm_medium');
    if (medium && ['email','newsletter'].includes(medium.toLowerCase())) return 'email';
    try {
      const host = new URL(document.referrer).hostname.toLowerCase();
      if (host === location.hostname || host.endsWith('.nolimitcontractor.net')) return 'internal';
      if (/(^|\.)google\.[a-z.]+$/.test(host)) return 'google';
      if (/(^|\.)bing\.com$/.test(host)) return 'bing';
      if (/(^|\.)(facebook|instagram|linkedin|pinterest|tiktok)\.com$/.test(host)) return 'social';
      return 'other';
    } catch { return 'direct'; }
  }
  function start() {
    if (document.visibilityState === 'hidden') return;
    document.removeEventListener('visibilitychange', start);
    metric('page_view',page);
    try {
      // A boolean per tab, never a visitor/session identifier or browsing history.
      if (!sessionStorage.getItem('nolimitVisitCounted')) {
        metric('visit_start','website'); metric('source',source());
        sessionStorage.setItem('nolimitVisitCounted','1');
      }
    } catch { /* Page views remain measurable; do not invent a unique visitor. */ }
    const category = categoryName(new URLSearchParams(location.search).get('cat'));
    if (categories.has(category)) once('gallery_view',category);
    if ('IntersectionObserver' in window) {
      const observer = new IntersectionObserver(entries => entries.forEach(entry => {
        if (entry.isIntersecting) {
          once('section_view',entry.target.id);
          if (entry.target.id === 'gallery') {
            const selected=categoryName(new URLSearchParams(location.search).get('cat') || document.querySelector('#galleryControls .active')?.dataset.cat);
            if(categories.has(selected))once('gallery_view',selected);
          }
          observer.unobserve(entry.target);
        }
      }), {threshold:0.01});
      ['about','services','gallery','catalogs','process','visit'].forEach(id => { const section=document.getElementById(id); if(section)observer.observe(section); });
    }
  }
  document.addEventListener('click',event => {
    const target=event.target.closest?.('a,button'); if (!target) return;
    const category=categoryName(target.dataset.cat);
    if (categories.has(category)) once('gallery_view',category);
    const href=target.getAttribute('href') || '';
    if (href.startsWith('tel:')) metric('contact_click','phone');
    else if (href.startsWith('mailto:')) metric('contact_click','email');
    else if (/\.pdf(?:[?#]|$)/i.test(href)) metric('catalog_open','document');
  });
  document.getElementById('visitForm')?.addEventListener('input',() => once('form_start','visit'));
  document.addEventListener('nolimit:visit-saved',() => metric('form_success','visit'));
  document.addEventListener('nolimit:visit-failed',() => metric('form_error','visit'));
  if (document.visibilityState === 'hidden') document.addEventListener('visibilitychange',start); else start();
})();
