function openModuleSwitcher(rawUrl) {
  const url = new URL(rawUrl, window.location.href);
  const paths = {attendance:'/launch/attendance',heatmap:'/heat-map/',manuals:'/manuals/',claims:'/launch/claims',insurance:'/launch/insurance'};
  const destination = paths[url.searchParams.get('module')];
  if (!destination) return;
  url.searchParams.set('popup', '1');
  url.searchParams.set('opener_origin', window.location.origin);
  const left = Math.max(0, window.screenX + (window.outerWidth - 560) / 2);
  const top = Math.max(0, window.screenY + (window.outerHeight - 680) / 2);
  const popup = window.open(url.toString(), 'martins_branch_switch', 'popup=yes,width=560,height=680,scrollbars=yes,left='+left+',top='+top);
  if (!popup) { window.alert('Allow popups for this module, then select Switch user / branch again.'); return; }
  if (window.__martinsSwitchCleanup) window.__martinsSwitchCleanup();
  const cleanup = () => { window.removeEventListener('message', receive); window.clearInterval(timer); };
  const receive = (event) => {
    if (event.source !== popup || event.origin !== url.origin || event.data?.type !== 'martins:branch-switched' || event.data.path !== destination) return;
    cleanup(); popup.close(); window.location.assign(new URL(destination, url.origin).toString());
  };
  window.addEventListener('message', receive);
  const timer = window.setInterval(() => { if (popup.closed) cleanup(); }, 1000);
  window.__martinsSwitchCleanup = cleanup;
  popup.focus();
}
document.addEventListener('click',function(event){const link=event.target.closest('[data-module-switch]');if(!link)return;event.preventDefault();openModuleSwitcher(link.href);});
