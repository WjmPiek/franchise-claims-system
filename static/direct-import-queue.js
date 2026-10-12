(()=>{
const form=document.getElementById('directQueueForm');if(!form)return;
form.addEventListener('submit',async event=>{
 event.preventDefault();const input=document.getElementById('directQueueFiles'),button=document.getElementById('directQueueStart'),status=document.getElementById('directQueueStatus'),list=document.getElementById('directQueueResults');
 const files=Array.from(input.files);if(!files.length)return;const failures=[];let saved=0,timer;button.disabled=true;input.disabled=true;status.hidden=false;list.replaceChildren();
 try{
 timer=setInterval(()=>fetch('/api/import/session',{credentials:'same-origin',cache:'no-store'}).catch(()=>{}),60000);
 for(let index=0;index<files.length;index++){
 const file=files[index];let confirmed=false;status.textContent=`File ${index+1} of ${files.length}: ${file.name}. Saved: ${saved}. Failed: ${failures.length}.`;
 try{
 if(!/\d{8}_to_\d{8}/.test(file.name)){confirmed=true;throw Error('Filename has no receipt month. Use single-file preparation to assign it.');}
 const data=new FormData();data.append('direct_file',file,file.name);
 const response=await fetch(form.action,{method:'POST',body:data,credentials:'same-origin',headers:{'X-Requested-With':'XMLHttpRequest','X-Import-Progress':'stream'}});let result;
 if((response.headers.get('Content-Type')||'').includes('application/x-ndjson')){
 const reader=response.body.getReader(),decoder=new TextDecoder();let pending='';while(true){const chunk=await reader.read();pending+=decoder.decode(chunk.value||new Uint8Array(),{stream:!chunk.done});const lines=pending.split('\n');pending=lines.pop();for(const line of lines){if(!line.trim())continue;const update=JSON.parse(line);if(update.type==='progress')status.textContent=`File ${index+1}/${files.length}: ${file.name} — ${update.stage}`+(update.current!=null?' — '+Number(update.current).toLocaleString()+(update.total?' / '+Number(update.total).toLocaleString():'')+' rows':'')+`\nSaved: ${saved}. Failed: ${failures.length}.`;if(update.type==='done'){result=update;confirmed=![401,403,409].includes(update.status);}}if(chunk.done)break;}
 }else{result=await response.json();confirmed=response.status!==409&&response.status!==401&&response.status!==403;}
 if(!response.ok||!result?.ok)throw Error(result?.error||'No successful import confirmation.');
 saved++;const row=document.createElement('li');row.textContent=`Saved ${file.name}: ${Number(result.source.rows).toLocaleString()} rows; risk ${result.source.risk}; retail ${result.source.retail}. Review items: ${result.issues}.`;list.append(row);
 }catch(error){failures.push({file:file.name,error:error.message,confirmed});const row=document.createElement('li');row.textContent=`${confirmed?'Fix and re-import':'Unconfirmed — check saved month'}: ${file.name} — ${error.message}`;list.append(row);if(!confirmed){status.textContent=`Queue paused. ${saved} files saved; ${failures.length} need review. Check the saved month and connection before retrying this file and the remaining files.`;return;}}
 }
 status.textContent=`Import queue finished. ${saved} of ${files.length} files saved. ${failures.length} files need correction and re-import. Completed months remain saved. Check Reconciliation and Review before sending insurer files.`;
 const report=document.createElement('a');report.href='/direct-data/reports';report.textContent='Open saved reports';list.append(report);
 }finally{clearInterval(timer);button.disabled=false;input.disabled=false;}
});})();
