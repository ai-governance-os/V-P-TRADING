const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),{JSDOM}=require('jsdom');
const code=fs.readFileSync('gas/Code.js','utf8');
function context(){const c={Session:{getScriptTimeZone:()=> 'Asia/Kuala_Lumpur'}};vm.createContext(c);vm.runInContext(code,c);return c}
function backend(){
  const c=context();let reads=0;
  c.sh_=()=>({getDataRange:()=>({getValues:()=>{reads++;return [['ID','VALUE'],['A',reads]]}})});
  c.readScope_(()=>{assert.equal(c.readTable_('TEST').rows[0].VALUE,1);c.readScope_(()=>c.readTable_('TEST'))});
  assert.equal(reads,1,'nested read-only work reads a table once');
  assert.equal(c.readTable_('TEST').rows[0].VALUE,2,'next call reads fresh data');
  assert.throws(()=>c.readScope_(()=>{c.readTable_('TEST');throw Error('failure')}));
  assert.equal(c.READ_SCOPE_,null,'failure clears request scope');
  c.readTable_('TEST');assert.equal(reads,4,'no scope leaks after failure');
  c.readScope_(()=>{c.readTable_('TEST');c.readTable_('TEST',true);assert.equal(c.readTable_('TEST').rows[0].VALUE,6)});
  let formats=0;
  const orders=Array.from({length:1200},(_,i)=>({ORDER_ID:'O'+i,SALESMAN:'CLIENT'+(i%60),DATE:i<100?'2026-08-01':'2026-09-01',MONTH:i<100?8:9,BRANCH:'B'+(i%3),SET_TYPE:'BAG',QTY:1,UNIT_PRICE:10,TOTAL_INCOME:10,INVOICE_TO:'SA'}));
  const invoices=Array.from({length:60},(_,i)=>({INV_NO:'INV'+i,YM:'2609',CUST_KEY:'CLIENT'+i+'|SA',SALESMAN:'CLIENT'+i,BILL_MODE:'SA',BILL_NAME:'CLIENT'+i,ORDERS:1,AMOUNT:10}));
  // Also exercise explicit mappings, per-order legacy keys and voided records.
  invoices[0].ORDER_IDS=JSON.stringify(['O120']);invoices[1].CUST_KEY='CLIENT1|SA#O121';invoices[2].STATUS='VOID';
  c.ensureCols_=name=>({rows:name==='ORDERS'?orders:[]});c.invSheet_=()=>({rows:invoices});
  c.fmtDate_=x=>{formats++;return x};c.billTo_=(name)=>({name,addr:'ADDRESS'});
  const result=c.listInvoiceMonth({ym:'2609'});
  assert.equal(formats,1200,'invoice indexing formats each order once, regardless of invoice count');
  assert.equal(result.candidates.find(x=>x.id==='O120').invNo,'INV0');
  assert.equal(result.candidates.find(x=>x.id==='O180').invNo,'','explicit invoice covers only selected order');
  assert.equal(result.candidates.find(x=>x.id==='O121').invNo,'INV1');
  assert.equal(result.candidates.find(x=>x.id==='O122').invNo,'','void invoices release coverage');
  assert.equal(result.candidates.find(x=>x.id==='O123').invNo,'INV3','legacy monthly invoice still reserves month');
  assert.equal(result.candidates.length,1100,'other month excluded');
  const pending=result.list.filter(x=>!x.issued).reduce((n,x)=>n+x.n,0);
  assert.equal(pending,result.candidates.filter(x=>!x.invNo).length,'monthly pending counts exclude issued orders');
  console.log('Read-scope and 1,200-order / 60-invoice regression checks passed; 1,200 date lookups');
}
async function ui(){
  const dom=new JSDOM(fs.readFileSync('gas/Index.html','utf8'),{runScripts:'dangerously',url:'https://test.local'}),w=dom.window,d=w.document,requests=[];
  w.google={script:{run:{withSuccessHandler:resolve=>({withFailureHandler:reject=>new Proxy({},{get:(_,fn)=>(...args)=>requests.push({fn,args,resolve,reject})})})}}};
  w.HAS_GS=true;w.S.user={name:'TEST',role:'admin'};
  const take=fn=>{const i=requests.findIndex(x=>x.fn===fn);assert.ok(i>=0,fn);return requests.splice(i,1)[0]};
  const response=ym=>({ok:true,ym,list:[],candidates:[],needGrade:[]});
  let p=w.loadInv('2609'),q=w.loadInv('2609');assert.equal(requests.length,1,'duplicate month requests join');
  take('listInvoiceMonth').resolve(response('2609'));await Promise.all([p,q]);
  await w.loadInv('2609');assert.equal(requests.length,0,'fresh invoice revisit avoids a request');
  p=w.loadInv('2608');const august=take('listInvoiceMonth');q=w.loadInv('2607');take('listInvoiceMonth').resolve(response('2607'));await q;
  august.resolve(response('2608'));await p;assert.equal(w.IV.ym,'2607','late invoice month cannot replace newer month');
  const key='listInvoiceMonth'+JSON.stringify([{ym:'2607'}]);w.VIEW_READ.cache[key].at-=16000;
  p=w.loadInv('2607');assert.ok(!d.querySelector('#sheetBody .spin'),'stale invoice list remains visible');
  w.startInvoiceMerge();take('listInvoiceMonth').resolve(response('2607'));await p;
  assert.match(d.getElementById('sheetBody').textContent,/选择订单合并开票/,'background refresh does not close selection');
  w.invalidateViews_('makeInvoicePdf');assert.ok(!w.VIEW_READ.cache[key],'issue/reprint forces invoice list refresh');
  p=w.loadInv('2609');w.closeSheetNow();take('listInvoiceMonth').resolve(response('2609'));await p;
  assert.ok(!d.getElementById('sheet').classList.contains('on'),'closed sheet remains closed');
  w.invalidateViews_('submitOrder');p=w.loadInv('2609');const stale=take('listInvoiceMonth');w.invalidateViews_('voidInvoice');stale.resolve(response('2609'));
  await new Promise(resolve=>setImmediate(resolve));take('listInvoiceMonth').resolve(response('2609'));await p;
  assert.ok(w.VIEW_READ.cache['listInvoiceMonth'+JSON.stringify([{ym:'2609'}])],'save race refetches invoices');
  d.getElementById('app').innerHTML='<div id="v-bill"></div>';w.S.view='bill';w.S.billMonth=9;
  w.renderBill=()=>{d.getElementById('v-bill').textContent='Month '+w.S.bill.month};
  p=w.loadBill();const sept=take('getStatement');w.S.billMonth=8;q=w.loadBill();take('getStatement').resolve({ok:true,month:8});await q;
  sept.resolve({ok:true,month:9});await p;assert.equal(w.S.bill.month,8,'late statement cannot overwrite chosen month');
  await w.loadBill();assert.equal(requests.length,0,'statement revisit is cached');
  dom.window.close();console.log('Invoice and monthly-statement cache/race/selection checks passed');
}
backend();ui().catch(e=>{console.error(e);process.exit(1)});
