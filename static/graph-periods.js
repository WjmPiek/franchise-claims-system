document.addEventListener('submit',async event=>{
  const form=event.target.closest('form[data-update-url]');if(!form)return;
  event.preventDefault();const card=form.closest('[data-graph-card]'),status=card.querySelector('.graph-status'),button=form.querySelector('button');
  const first=form.elements.date_from.value,last=form.elements.date_to.value;
  if(!first||!last||first>last){status.textContent='Choose valid dates, with From date before To date.';return;}
  button.disabled=true;status.textContent='Updating this graph…';card.setAttribute('aria-busy','true');
  try{
    const response=await fetch(form.dataset.updateUrl+'?'+new URLSearchParams(new FormData(form)),{credentials:'same-origin',cache:'no-store'});
    if(response.redirected)throw new Error('Your session expired. Please sign in again.');
    const html=await response.text();if(!response.ok)throw new Error(response.status===400?html:'Could not update this graph. Please try again.');
    const template=document.createElement('template');template.innerHTML=html;const updated=template.content.querySelector('[data-graph-card]');
    if(!updated||updated.dataset.graphCard!==card.dataset.graphCard)throw new Error('Could not confirm the updated graph. Please try again.');
    card.replaceWith(updated);
  }catch(error){status.textContent=error.message;button.disabled=false;card.removeAttribute('aria-busy');}
});
