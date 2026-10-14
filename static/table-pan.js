(() => {
  const registered = new WeakSet();
  function init() {
    const wrappers = new Set();
    document.querySelectorAll('table').forEach(table => {
      let el = table.parentElement.closest('.table-wrap,.central-table,.table-responsive,.data-table-wrapper,.wide-table,.business-scroll,.verification-scroll,.table-pan-shell');
      if (!el) {
        for (let parent = table.parentElement; parent && parent !== document.body; parent = parent.parentElement) {
          if (['auto','scroll'].includes(getComputedStyle(parent).overflowX)) { el = parent; break; }
        }
      }
      if (!el) { el = document.createElement('div'); el.className = 'table-pan-shell'; table.before(el); el.append(table); }
      wrappers.add(el);
    });
    wrappers.forEach(el => {
      if (registered.has(el)) return;
      registered.add(el);
      const refresh = () => el.classList.toggle('can-pan', el.scrollWidth > el.clientWidth + 1);
      refresh();
      if (window.ResizeObserver) { const observer = new ResizeObserver(refresh); observer.observe(el); if (el.querySelector('table')) observer.observe(el.querySelector('table')); }
      let pointer = null, suppressClickUntil = 0;
      el.addEventListener('pointerdown', event => {
        if (!event.isPrimary || event.button !== 0 || !el.classList.contains('can-pan') || event.target.closest('input,select,textarea,button')) return;
        pointer = { id: event.pointerId, x: event.clientX, y: event.clientY, left: el.scrollLeft, dragging: false };
      });
      el.addEventListener('dragstart', event => { if (pointer) event.preventDefault(); });
      document.addEventListener('pointermove', event => {
        if (!pointer || pointer.id !== event.pointerId) return;
        const dx = event.clientX - pointer.x, dy = event.clientY - pointer.y;
        if (!pointer.dragging) {
          if (Math.abs(dy) > 8 && Math.abs(dy) > Math.abs(dx)) { pointer = null; return; }
          if (Math.abs(dx) < 6) return;
          pointer.dragging = true; el.classList.add('is-panning'); el.setPointerCapture(event.pointerId);
        }
        event.preventDefault(); el.scrollLeft = pointer.left - dx;
      }, { passive: false });
      const finish = event => {
        if (!pointer || pointer.id !== event.pointerId) return;
        if (pointer.dragging) suppressClickUntil = Date.now() + 300;
        if (el.hasPointerCapture(event.pointerId)) el.releasePointerCapture(event.pointerId);
        pointer = null; el.classList.remove('is-panning');
      };
      document.addEventListener('pointerup', finish); document.addEventListener('pointercancel', finish);
      el.addEventListener('click', event => { if (Date.now() < suppressClickUntil) { event.preventDefault(); event.stopPropagation(); } }, true);
    });
  }
  function start() {
    init();
    let scheduled = false;
    new MutationObserver(() => { if (!scheduled) { scheduled = true; requestAnimationFrame(() => { scheduled = false; init(); }); } }).observe(document.body, {childList:true,subtree:true});
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
