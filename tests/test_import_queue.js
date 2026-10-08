// Run with Node: exercise the actual dashboard handler with browser doubles.
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const html = fs.readFileSync('templates/dashboard.html', 'utf8');
const start = html.lastIndexOf('(function(){', html.indexOf("const form = document.getElementById('importForm')"));
const end = html.indexOf('})();', start) + 5;
const script = html.slice(start, end);

async function run(failAt, authAt=-1, errorMode="server"){
  const requests = [];
  const files = Array.from({length:22}, (_, i) => ({name:`month-${i+1}.xlsx`,size:1000}));
  const nodes = new Map();
  const node = id => {
    if(!nodes.has(id)) nodes.set(id, {style:{}, classList:{add(){},remove(){}}, textContent:''});
    return nodes.get(id);
  };
  const form = node('importForm');
  form.querySelector = selector => ({files: selector.includes('excel_file') ? files : []});
  form.getAttribute = () => '/dashboard';
  form.addEventListener = (_, fn) => form.submitHandler = fn;
  class Data {
    constructor(){this.files=[];}
    delete(){}
    append(field, file){this.files.push({field,file});}
  }
  class Xhr {
    constructor(){this.upload={}; requests.push(this);}
    open(){}
    setRequestHeader(){}
    send(data){this.data=data;}
  }
  const location = {href:'https://claims.example/dashboard'};
  let sessionChecks=0;
  const timers=new Map();
  const fetch=async () => {const ok=sessionChecks++ !== authAt; return {ok,json:async()=>({ok,error:ok?'':'Your login session expired.'})};};
  const flush=()=>new Promise(resolve=>setImmediate(resolve));
  vm.runInNewContext(script, {document:{getElementById:node}, window:{location}, XMLHttpRequest:Xhr, FormData:Data, fetch, URL, setInterval:(fn,ms)=>{const id=timers.size+1;timers.set(id,{fn,ms});return id;}, clearInterval(id){timers.delete(id);}, Date, Math, Array, JSON});
  form.submitHandler({preventDefault(){}});
  await flush();
  for(let i=0;i<22;i++){
    if(i===authAt){
      assert.equal(requests.length,i,'expired session prevents upload');
      assert(node('importProgressText').textContent.includes('session expired'));
      assert.equal(timers.size,0);
      return;
    }
    assert.equal(requests.length,i+1, 'only one file in flight');
    const req=requests[i];
    assert.equal(req.data.files.length,1);
    req.upload.onprogress({lengthComputable:true,loaded:100,total:100});
    assert.equal(node('importProgressPct').textContent,'Processing');
    req.upload.onload();
    req.responseText='{"type":"progress","stage":"Reading Excel rows","current":500,"total":1000}\n';
    req.onprogress();
    assert(node('importStepRows').textContent.includes('50% of this stage'));
    assert.equal(node('importProgressPct').textContent,'50% of stage');
    if(i===failAt){
      req.status=errorMode==='session'?401:500; req.responseText=errorMode==='session'?JSON.stringify({ok:false,error:'Your login session expired.'}):'server timeout'; req.onload();
      await flush();
      assert.equal(requests.length,i+1, 'failure stops queue');
      assert.equal(node('importProgressPct').textContent,'Error');
      assert.equal(location.href,'https://claims.example/dashboard');
      assert.equal(timers.size,0);
      if(errorMode==='session') assert(node('importProgressText').textContent.includes('session expired'));
      assert(node('importProgressText').textContent.includes(`Completed: ${i} of 22`));
      return;
    }
    req.status=200; req.responseText+='{"type":"done","ok":'; req.onprogress();
    req.responseText+='true}\n'; req.onload();
    await flush();
  }
  assert.equal(node('importProgressPct').textContent,'100%');
  assert(location.href.includes('dashboard'));
}
(async()=>{await run(-1);await run(3);await run(4,4);await run(4,-1,'session');console.log('PASS: sequential 22-file queue, partial streams, explicit errors, preflight session expiry blocks upload, JSON 401 stops queue, timers cleared.');})().catch(error=>{console.error(error);process.exitCode=1;});
