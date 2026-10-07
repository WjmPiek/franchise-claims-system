// Run with Node: exercise the actual dashboard handler with browser doubles.
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const html = fs.readFileSync('templates/dashboard.html', 'utf8');
const start = html.lastIndexOf('(function(){', html.indexOf("const form = document.getElementById('importForm')"));
const end = html.indexOf('})();', start) + 5;
const script = html.slice(start, end);

function run(failAt){
  const requests = [];
  const files = Array.from({length:22}, (_, i) => ({name:`month-${i+1}.xlsx`}));
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
  const location = {href:''};
  vm.runInNewContext(script, {document:{getElementById:node}, window:{location}, XMLHttpRequest:Xhr, FormData:Data, setInterval:()=>1, clearInterval(){}, Date, Math, Array, JSON});
  form.submitHandler({preventDefault(){}});
  for(let i=0;i<22;i++){
    assert.equal(requests.length,i+1, 'only one file in flight');
    const req=requests[i];
    assert.equal(req.data.files.length,1);
    req.upload.onprogress({lengthComputable:true,loaded:100,total:100});
    assert.equal(node('importProgressPct').textContent,'Processing');
    req.upload.onload();
    if(i===failAt){
      req.status=500; req.responseText='server timeout'; req.onload();
      assert.equal(requests.length,i+1, 'failure stops queue');
      assert.equal(node('importProgressPct').textContent,'Error');
      assert.equal(location.href,'');
      assert(node('importProgressText').textContent.includes(`Completed: ${i} of 22`));
      return;
    }
    req.status=200; req.responseText='{"ok":true}'; req.onload();
  }
  assert.equal(node('importProgressPct').textContent,'100%');
  assert(location.href.includes('dashboard'));
}
run(-1);
run(3);
console.log('PASS: 22 files run sequentially, upload switches to processing, errors stop the queue and report completed files.');
