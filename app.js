'use strict';
/* SANAD: private client-side accounting-review helper. No network requests, persistence, or bookkeeping writes. */
const COMPANIES=[
 {id:'chips',name:'مصنع الشيبس',sub:'الإنتاج والتصنيع',symbol:'▥',color:'#fff1dd',ink:'#9d601f'},
 {id:'shop',name:'المحل التجاري',sub:'المبيعات والمخزون',symbol:'▦',color:'#e7f0ff',ink:'#3567aa'},
 {id:'rental',name:'شركة تأجير السيارات',sub:'التأجير والأسطول',symbol:'▰',color:'#f9e7e7',ink:'#ad4349'},
 {id:'kitchens',name:'شركة المطابخ',sub:'التصميم والتصنيع',symbol:'▧',color:'#e8f4eb',ink:'#31825c'}
];
const makeData=()=>({ledger:[],bank:[],approved:[],demo:false,ledgerName:'',bankName:''});
const db=Object.fromEntries(COMPANIES.map(c=>[c.id,makeData()]));
let activeCompany='chips', currentView='home',pending=null,modalCallback=null,toastTimer=null,manualSequence=0,entryEditing=null;
const $=id=>document.getElementById(id);
const esc=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const company=()=>COMPANIES.find(c=>c.id===activeCompany);
const data=()=>db[activeCompany];
const hasData=()=>data().ledger.length||data().bank.length;
function toast(message,error=false){const el=$('toast');el.textContent=message;el.className='toast show'+(error?' error':'');clearTimeout(toastTimer);toastTimer=setTimeout(()=>el.className='toast',4200)}
function confirmDialog(title,txt,fn){$('modalTitle').textContent=title;$('modalText').textContent=txt;modalCallback=fn;$('modal').classList.remove('hidden');$('modalCancel').focus()}
function closeModal(){$('modal').classList.add('hidden');modalCallback=null}
const norm=x=>String(x??'').normalize('NFKC').replace(/[\u064b-\u065f\u0670\u0640]/g,'').replace(/[أإآ]/g,'ا').replace(/ة/g,'ه').replace(/[\s_\-\/().]/g,'').toLowerCase();
const arDigits=x=>String(x).replace(/[٠-٩]/g,d=>String(d.charCodeAt(0)-1632)).replace(/[۰-۹]/g,d=>String(d.charCodeAt(0)-1776));
function parseAmount(raw){
 if(raw===null||raw===undefined||String(raw).trim()==='')return NaN;
 let s=arDigits(raw).replace(/\u00a0/g,' ').replace(/[٬\s]/g,'').replace(/[٫]/g,'.').replace(/[−–]/g,'-');
 const par=/^\(.*\)$/.test(s);if(par)s=s.slice(1,-1);
 s=s.replace(/(?:ر\.?ع\.?|ريال(?:\s*عماني)?|omr|rial)/gi,'').replace(/\s/g,'');
 if(!/^[\d.,+\-]+$/.test(s))return NaN;
 if(s.includes(',')&&s.includes('.')){if(s.lastIndexOf(',')>s.lastIndexOf('.'))s=s.replace(/\./g,'').replace(',','.');else s=s.replace(/,/g,'')}
 else if(s.includes(',')){const parts=s.split(',');if(parts.length>2||parts[parts.length-1].length===3)s=parts.join('');else s=s.replace(',','.')} 
 if(!/^[+\-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(s))return NaN;
 const n=Number(s)*(par?-1:1);return Number.isFinite(n)?n:NaN;
}
function normalizeDate(raw){
 if(raw===null||raw===undefined||String(raw).trim()==='')return '';
 let s=arDigits(raw).trim();
 if(/^\d{4,5}(?:\.0+)?$/.test(s)){const serial=Number(s);if(serial>20000&&serial<100000){const dt=new Date(Date.UTC(1899,11,30)+Math.floor(serial)*86400000);return dt.toISOString().slice(0,10)}}
 const match=s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})(?:[T\s].*)?$/);
 const other=s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);
 let y,m,d;if(match){y=+match[1];m=+match[2];d=+match[3]}else if(other){d=+other[1];m=+other[2];y=+other[3]}else return '';
 const dt=new Date(Date.UTC(y,m-1,d));if(dt.getUTCFullYear()!==y||dt.getUTCMonth()!==m-1||dt.getUTCDate()!==d)return '';
 return `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
}
function parseDelimited(input,separator){
 const out=[];let row=[],field='',quoted=false;const text=String(input??'').replace(/^\uFEFF/,'');
 for(let i=0;i<text.length;i++){const ch=text[i];if(quoted){if(ch==='"'&&text[i+1]==='"'){field+='"';i++}else if(ch==='"')quoted=false;else field+=ch;continue}
 if(ch==='"'&&field==='')quoted=true;else if(ch===separator){row.push(field);field=''}else if(ch==='\n'||ch==='\r'){if(ch==='\r'&&text[i+1]==='\n')i++;row.push(field);if(row.some(v=>v.trim()))out.push(row);row=[];field=''}else field+=ch;
 }
 if(quoted)throw Error('ملف CSV فيه علامات اقتباس غير مغلقة.');row.push(field);if(row.some(v=>v.trim()))out.push(row);
 if(out.length>10002)throw Error('الملف أكبر من حد النسخة الأولى: 10,000 عملية.');return out;
}
function detectDelimited(text){const first=String(text).split(/\r?\n/,1)[0]||'';const scored=['\t',',',';'].map(sep=>({sep,count:parseDelimited(first,sep)[0]?.length||0})).sort((a,b)=>b.count-a.count);return parseDelimited(text,scored[0].sep)}
function xml(text){if(/<!DOCTYPE|<!ENTITY/i.test(text))throw Error('ملف Excel يحوي تعريفات XML غير مدعومة.');const doc=new DOMParser().parseFromString(text,'application/xml');if(doc.getElementsByTagName('parsererror').length)throw Error('تعذر تحليل أحد أجزاء ملف Excel.');return doc}
function colIndex(address){const m=String(address||'').match(/^[A-Z]+/i);if(!m)return -1;let n=0;for(const c of m[0].toUpperCase())n=n*26+c.charCodeAt(0)-64;return n-1}
function xlsxText(node,tag){return [...node.getElementsByTagName(tag)].map(n=>n.textContent||'').join('')}
async function parseXlsx(buffer){
 if(typeof JSZip==='undefined')throw Error('مكتبة Excel غير موجودة. افتح المجلد الكامل.');
 const zip=await JSZip.loadAsync(buffer,{checkCRC32:false});const workbook=zip.file('xl/workbook.xml');if(!workbook)throw Error('ملف Excel غير صالح أو محمي بكلمة مرور.');
 const get=async(path,max=20000000)=>{const f=zip.file(path);if(!f)throw Error('جزء مفقود داخل ملف Excel: '+path);if(f._data?.uncompressedSize>max)throw Error('حجم ورقة Excel كبير جدًا.');const s=await f.async('string');if(s.length>max)throw Error('الملف كبير جدًا.');return s};
 const wb=xml(await get('xl/workbook.xml',1000000)), relFile=zip.file('xl/_rels/workbook.xml.rels'), relDoc=relFile?xml(await get('xl/_rels/workbook.xml.rels',1000000)):null;
 const relations=new Map();if(relDoc)for(const r of relDoc.getElementsByTagName('Relationship'))relations.set(r.getAttribute('Id'),r.getAttribute('Target'));
 let shared=[];if(zip.file('xl/sharedStrings.xml')){const strDoc=xml(await get('xl/sharedStrings.xml',12000000));shared=[...strDoc.getElementsByTagName('si')].map(node=>xlsxText(node,'t'))}
 const sheetEls=[...wb.getElementsByTagName('sheet')],sheets=[];
 for(const [idx,sh] of sheetEls.entries()){
  let target=relations.get(sh.getAttribute('r:id'))||`worksheets/sheet${idx+1}.xml`;
  target=target.startsWith('/')?target.slice(1):target.startsWith('xl/')?target:`xl/${target}`;
  const path=target.split('/').reduce((parts,part)=>{if(part==='..')parts.pop();else if(part!=='.')parts.push(part);return parts},[]).join('/');
  if(!zip.file(path))continue;const sd=xml(await get(path));const rows=[];
  for(const row of sd.getElementsByTagName('row')){
   const values=[];let next=0;
   for(const c of row.getElementsByTagName('c')){
    let ix=colIndex(c.getAttribute('r'));if(ix<0)ix=next;if(ix>100)continue;next=ix+1;
    const kind=c.getAttribute('t'),val=c.getElementsByTagName('v')[0]?.textContent??'';
    values[ix]=kind==='s'?(shared[Number(val)]??''):kind==='inlineStr'?xlsxText(c,'t'):kind==='b'?(val==='1'?'TRUE':'FALSE'):val;
   }
   if(values.some(x=>String(x??'').trim()))rows.push(Array.from({length:Math.min(values.length,101)},(_,i)=>String(values[i]??'').slice(0,2500)));
   if(rows.length>10001)throw Error('الورقة أكبر من حد النسخة الأولى: 10,000 عملية.');
  }
  sheets.push({name:sh.getAttribute('name')||`ورقة ${idx+1}`,rows});
 }
 if(!sheets.length)throw Error('لم أجد أوراق عمل قابلة للقراءة داخل الملف.');return sheets;
}
const aliases={
 date:['التاريخ','تاريخ','تاريخالعمليه','تاريخالفاتوره','تاريخالقيد','date','transactiondate','valuedate','postingdate'],
 amount:['المبلغ','المبلغالصافي','قيمه','القيمه','صافيالمبلغ','amount','netamount','value','total'],
 description:['البيان','الوصف','التفاصيل','شرح','الملاحظات','description','details','narration','memo','particulars'],
 reference:['رقمالفاتوره','رقمالمستند','رقمالعمليه','المرجع','رقمالقيد','رقمالايصال','ref','reference','invoicenumber','invoiceno','transactionid','documentnumber'],
 type:['نوعالعمليه','النوع','حركه','اتجاه','نوع','type','transactiontype','direction','drcr'],
 party:['المورد','العميل','الطرف','اسمالطرف','اسمالمورد','اسمالعميل','supplier','customer','party','vendor'],
 debit:['سحب','مدين','المسحوبات','المصروف','خصم','debit','withdrawal','withdrawals','paidout'],
 credit:['ايداع','دائن','الايداعات','القبض','credit','deposit','deposits','paidin']
};
const fields={ledger:[['date','التاريخ *'],['description','البيان *'],['amount','المبلغ *'],['reference','رقم المستند (اختياري)'],['type','نوع الحركة: قبض/صرف (مهم للمطابقة)'],['party','العميل / المورد (اختياري)']],bank:[['date','التاريخ *'],['description','البيان (اختياري)'],['amount','المبلغ الموقّع (إذا موجود)'],['debit','السحب (بديل للمبلغ الموقّع)'],['credit','الإيداع (بديل للمبلغ الموقّع)'],['reference','المرجع (اختياري)'],['type','نوع الحركة (اختياري)']]};
function autoField(headers,key){const candidates=aliases[key].map(norm);const normalized=headers.map(norm);for(const c of candidates){const i=normalized.indexOf(c);if(i!==-1)return String(i)}for(const c of candidates){const i=normalized.findIndex(h=>h.includes(c)&&c.length>=5);if(i!==-1)return String(i)}return ''}
function setupMapping(){
 const rows=pending?.sheets[pending.sheet]?.rows||[];if(rows.length<2)throw Error('الورقة لازم تحتوي عنوان الأعمدة وصف بيانات واحد على الأقل.');
 pending.headers=rows[0].map(x=>String(x??'').trim()).slice(0,100);pending.rows=rows.slice(1).filter(r=>r.some(v=>String(v??'').trim()));
 if(!pending.headers.some(Boolean))throw Error('أول صف لازم يحتوي أسماء الأعمدة.');
 $('mappingPanel').classList.remove('hidden');$('mappingTitle').textContent=pending.kind==='ledger'?'راجع أعمدة دفتر العمليات':'راجع أعمدة كشف البنك';
 $('mappingHelp').textContent='تأكد من التاريخ والمبلغ. إذا كل المبالغ موجبة، حدد عمود نوع الحركة أو السحب والإيداع؛ وإلا ما نقدر نقترح مطابقة آمنة.';
 $('sheetRow').classList.toggle('hidden',pending.sheets.length<2);$('sheetSelect').innerHTML=pending.sheets.map((s,i)=>`<option value="${i}">${esc(s.name)}</option>`).join('');$('sheetSelect').value=String(pending.sheet);
 $('mappingFields').innerHTML=fields[pending.kind].map(([key,label])=>`<div class="map-field"><label for="map-${key}">${esc(label)}</label><select id="map-${key}" data-map="${key}"><option value="">— غير محدد —</option>${pending.headers.map((h,i)=>`<option value="${i}">${esc(h||`العمود ${i+1}`)}</option>`).join('')}</select></div>`).join('');
 for(const [key] of fields[pending.kind])$(`map-${key}`).value=autoField(pending.headers,key);
 $('pendingRows').textContent=`${pending.rows.length.toLocaleString('en-US')} صف بعد العناوين`;
 $('mappingPreview').innerHTML=renderTable(pending.headers.slice(0,9),pending.rows.slice(0,5).map(row=>row.slice(0,9)));
 $('mappingPanel').scrollIntoView({behavior:'smooth',block:'start'});
}
function renderTable(headers,rows){if(!rows.length)return '<div class="empty"><strong>ما توجد بيانات</strong>جرّب استيراد ملف أو اختيار شركة ثانية.</div>';return `<table><thead><tr>${headers.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(row=>`<tr>${row.map(v=>`<td class="wrapcell">${esc(v)}</td>`).join('')}</tr>`).join('')}</tbody></table>`}
function classifyDirection(type){const t=norm(type);if(!t)return 0;if(/مصروف|مشتري|سحب|صرف|دفع|مدين|debit|expense|withdraw|outgoing|purchase|payment|paid/.test(t))return -1;if(/ايراد|مبيع|قبض|ايداع|دخل|دائن|credit|income|sale|deposit|receipt|incoming/.test(t))return 1;return 0}
function getMapped(row,map,key){const idx=map[key];return idx===''||idx===undefined?'':String(row[Number(idx)]??'').trim()}
function makeRecords(p){
 const map=Object.fromEntries(fields[p.kind].map(([key])=>[key,$(`map-${key}`).value]));
 if(map.date===''||(p.kind==='ledger'&&(map.description===''||map.amount===''))||(p.kind==='bank'&&map.amount===''&&map.debit===''&&map.credit===''))throw Error('حدد الأعمدة المطلوبة المعلّمة بالنجمة، وللبنك اختر المبلغ أو السحب/الإيداع.');
 const used=Object.entries(map).filter(([,v])=>v!=='').map(([,v])=>v);if(used.length!==new Set(used).size)throw Error('ما يصير تختار نفس العمود لأكثر من حقل. راجع المطابقة.');
 const records=[];for(const [i,row] of p.rows.entries()){
  const rawAmount=getMapped(row,map,'amount'),rawDebit=getMapped(row,map,'debit'),rawCredit=getMapped(row,map,'credit'),type=getMapped(row,map,'type');
  let amount=parseAmount(rawAmount),unknown=false,dir=classifyDirection(type);
  if(p.kind==='bank'&&map.amount===''){
    const dr=rawDebit===''?0:parseAmount(rawDebit),cr=rawCredit===''?0:parseAmount(rawCredit);
    amount=Number.isFinite(dr)&&Number.isFinite(cr)?Math.abs(cr)-Math.abs(dr):NaN;
    if(dr>0&&cr>0)amount=NaN;
  }else if(Number.isFinite(amount)){
    if(dir!==0)amount=dir*Math.abs(amount);
    else if(amount<0||String(rawAmount).trim().startsWith('+'))unknown=false;
    else unknown=true;
  }
  records.push({id:`${p.kind}-${i+1}`,row:i+2,date:normalizeDate(getMapped(row,map,'date')),rawDate:getMapped(row,map,'date'),description:getMapped(row,map,'description'),reference:getMapped(row,map,'reference'),party:getMapped(row,map,'party'),type,amount,rawAmount:p.kind==='bank'&&map.amount===''?`${rawCredit}|${rawDebit}`:rawAmount,directionUnknown:unknown});
 }
 return records;
}
function analyseLedger(records){
 const seen=new Map();return records.map(r=>{
  const issues=[];
  if(!r.rawDate||!r.date)issues.push('تاريخ مفقود أو غير صالح');
  if(!r.description)issues.push('البيان ناقص');
  if(!Number.isFinite(r.amount))issues.push('مبلغ مفقود أو غير صالح');
  else if(r.amount===0)issues.push('مبلغ يساوي صفر');
  if(r.directionUnknown&&Number.isFinite(r.amount)&&r.amount!==0)issues.push('اتجاه الحركة غير محدد للمطابقة');
  if(Number.isFinite(r.amount)&&r.amount!==0){
   const amount=Math.abs(r.amount).toFixed(3),ref=norm(r.reference),party=norm(r.party),description=norm(r.description);
   const key=ref?`ref|${ref}|${party}|${amount}`:(r.date&&description?`row|${r.date}|${description}|${r.amount.toFixed(3)}`:'');
   if(key){if(seen.has(key))issues.push(`تكرار محتمل مع الصف ${seen.get(key)}`);else seen.set(key,r.row)}
  }
  return {...r,issues};
 });
}
const keyPair=(l,b)=>`${l.id}|${b.id}`;
const dateDiff=(a,b)=>Math.abs((new Date(a+'T00:00:00Z')-new Date(b+'T00:00:00Z'))/86400000);
function matchResult(){
 const d=data(),ledger=analyseLedger(d.ledger),bank=d.bank,approved=[];
 for(const k of d.approved){const [lId,bId]=k.split('|'),l=ledger.find(x=>x.id===lId),b=bank.find(x=>x.id===bId);if(l&&b)approved.push({l,b,key:k,approved:true})}
 const approvedL=new Set(approved.map(x=>x.l.id)),approvedB=new Set(approved.map(x=>x.b.id));
 const eligibleL=ledger.filter(r=>r.date&&Number.isFinite(r.amount)&&r.amount!==0&&!r.directionUnknown&&!approvedL.has(r.id));
 const eligibleB=bank.filter(r=>r.date&&Number.isFinite(r.amount)&&r.amount!==0&&!r.directionUnknown&&!approvedB.has(r.id));
 const ledgerCandidates=new Map(),bankCandidates=new Map();
 for(const l of eligibleL)for(const b of eligibleB){if(Math.abs(l.amount-b.amount)>.00001||dateDiff(l.date,b.date)>3)continue;const k=keyPair(l,b);if(!ledgerCandidates.has(l.id))ledgerCandidates.set(l.id,[]);if(!bankCandidates.has(b.id))bankCandidates.set(b.id,[]);ledgerCandidates.get(l.id).push(b);bankCandidates.get(b.id).push(l)}
 const suggested=[];for(const l of eligibleL){const bs=ledgerCandidates.get(l.id)||[];if(bs.length===1&&(bankCandidates.get(bs[0].id)||[]).length===1)suggested.push({l,b:bs[0],key:keyPair(l,bs[0]),approved:false})}
 const unmatchedLedger=ledger.filter(r=>!approvedL.has(r.id));const unmatchedBank=bank.filter(r=>!approvedB.has(r.id));
 return {ledger,bank,approved,suggested,unmatchedLedger,unmatchedBank,ambiguous:eligibleL.filter(l=>(ledgerCandidates.get(l.id)||[]).length>1).length};
}
function money(n){return Number.isFinite(n)?new Intl.NumberFormat('en-US',{minimumFractionDigits:3,maximumFractionDigits:3}).format(n):'غير صالح'}
function number(n){return n.toLocaleString('en-US')}
function renderCompanies(){
 $('companyCards').innerHTML=COMPANIES.map(c=>`<button class="company-card ${c.id===activeCompany?'selected':''}" data-company="${c.id}" aria-pressed="${c.id===activeCompany}"><span class="company-icon" style="background:${c.color};color:${c.ink}">${c.symbol}</span><span class="company-name">${esc(c.name)}</span><span class="company-footer"><span>${esc(c.sub)}</span><span>${c.id===activeCompany?'● الحالية':'اختيار ←'}</span></span></button>`).join('');
}
function renderHome(){const m=matchResult(),d=data(),flagged=m.ledger.filter(x=>x.issues.length).length;$('statOperations').textContent=number(d.ledger.length);$('statWarnings').textContent=number(flagged);$('statApproved').textContent=number(m.approved.length);$('statBank').textContent=number(m.unmatchedBank.length);$('demoIndicator').textContent=d.demo?'بيانات تجريبية':hasData()?'بيانات مسجّلة':'بدون بيانات';$('demoIndicator').className='badge '+(d.demo?'badge-green':'badge-muted');$('sideFlag').textContent=flagged;$('sideFlag').classList.toggle('hidden',!flagged)}
function renderReview(){const records=analyseLedger(data().ledger),flags=records.filter(r=>r.issues.length),duplicates=flags.filter(r=>r.issues.some(x=>x.includes('تكرار'))),missing=flags.filter(r=>r.issues.some(x=>x.includes('ناقص')||x.includes('غير صالح')));$('reviewTotal').textContent=number(records.length);$('reviewFlags').textContent=number(flags.length);$('reviewDuplicates').textContent=number(duplicates.length);$('reviewMissing').textContent=number(missing.length);
 const filter=$('reviewFilter').value,q=norm($('reviewSearch').value);let filtered=records.filter(r=>{if(filter==='flagged'&&!r.issues.length)return false;if(filter==='duplicate'&&!r.issues.some(x=>x.includes('تكرار')))return false;if(filter==='missing'&&!r.issues.some(x=>x.includes('ناقص')||x.includes('غير صالح')))return false;return !q||norm([r.description,r.reference,r.party].join(' ')).includes(q)});
 $('reviewTable').innerHTML=renderTable(['الصف','التاريخ','البيان','المرجع','المبلغ (ر.ع)','الملاحظات'],filtered.map(r=>[r.row,r.date||r.rawDate||'—',r.description||'—',r.reference||'—',money(r.amount),r.issues.length?r.issues.join('، '):'لا توجد ملاحظات آلية']));
}
function renderMatch(){const m=matchResult();$('matchSuggested').textContent=number(m.suggested.length);$('matchApproved').textContent=number(m.approved.length);$('matchLedger').textContent=number(m.unmatchedLedger.length);$('matchBank').textContent=number(m.unmatchedBank.length);
 const pairs=[...m.approved,...m.suggested];$('matchTable').innerHTML=pairs.length?`<table><thead><tr>${['دفتر: البيان','دفتر: التاريخ','البنك: البيان','البنك: التاريخ','المبلغ (ر.ع)','الحالة','الإجراء'].map(h=>`<th>${h}</th>`).join('')}</tr></thead><tbody>${pairs.map(p=>`<tr><td class="wrapcell">${esc(p.l.description||'—')}</td><td>${esc(p.l.date)}</td><td class="wrapcell">${esc(p.b.description||'—')}</td><td>${esc(p.b.date)}</td><td>${esc(money(p.l.amount))}</td><td><span class="badge ${p.approved?'badge-green':'badge-orange'}">${p.approved?'اعتمدتها':'مقترحة فقط'}</span></td><td><button class="inline-action ${p.approved?'undo':''}" data-pair="${esc(p.key)}" data-pair-action="${p.approved?'undo':'approve'}">${p.approved?'إلغاء الاعتماد':'اعتماد المقترح'}</button></td></tr>`).join('')}</tbody></table>`:'<div class="empty"><strong>ما توجد مطابقات مقترحة</strong>استورد الملفين، أو راجع اتجاه الحركة والتاريخ والمبلغ. العمليات المتشابهة المتعددة تُترك للمراجعة.</div>';
 const recordRows=rows=>rows.map(r=>[r.date||'—',r.description||'—',money(r.amount),r.directionUnknown?'اتجاه غير محدد':!Number.isFinite(r.amount)?'مبلغ غير صالح':'غير معتمدة']);
 $('unmatchedLedger').innerHTML=renderTable(['التاريخ','البيان','المبلغ (ر.ع)','السبب / الحالة'],recordRows(m.unmatchedLedger));$('unmatchedBank').innerHTML=renderTable(['التاريخ','البيان','المبلغ (ر.ع)','السبب / الحالة'],recordRows(m.unmatchedBank));
}
function getReport(){const m=matchResult(),d=data(),flags=m.ledger.filter(r=>r.issues.length),dups=flags.filter(r=>r.issues.some(x=>x.includes('تكرار'))),invalid=flags.filter(r=>r.issues.some(x=>x.includes('ناقص')||x.includes('غير صالح')));return {m,d,flags,dups,invalid,company:company().name}}
function renderReport(){const r=getReport();$('reportSubheading').textContent=`${r.company} | ${r.d.demo?'بيانات تجريبية':'بيانات الجلسة الحالية'}`;$('reportDate').textContent=`تاريخ إعداد التقرير: ${new Date().toLocaleString('en-GB')}`;
 const vals=[['عمليات دفتر الحسابات',r.d.ledger.length],['العمليات بملاحظات',r.flags.length],['التكرار المحتمل',r.dups.length],['المطابقات المقترحة',r.m.suggested.length],['المطابقات المعتمدة',r.m.approved.length],['حركات البنك غير المعتمدة',r.m.unmatchedBank.length]];
 $('reportStats').innerHTML=vals.map(([k,v])=>`<div class="report-stat"><span>${esc(k)}</span><strong>${number(v)}</strong></div>`).join('');
 const entries=[`تم فحص ${number(r.d.ledger.length)} عملية في دفتر العمليات و${number(r.d.bank.length)} حركة في كشف البنك.`,`ظهرت ملاحظات محتملة في ${number(r.flags.length)} عملية، تشمل ${number(r.dups.length)} تكرارًا محتملًا و${number(r.invalid.length)} عملية ببيانات ناقصة أو غير صالحة.`,`يوجد ${number(r.m.suggested.length)} اقتراح مطابقة ينتظر اعتمادك، و${number(r.m.approved.length)} مطابقة اعتمدتها داخل الأداة.`,`المتبقي بدون اعتماد: ${number(r.m.unmatchedLedger.length)} عملية دفتر و${number(r.m.unmatchedBank.length)} حركة بنك.`,...(r.m.ambiguous?[`توجد ${number(r.m.ambiguous)} عملية دفتر لها أكثر من مرشح بنكي، لذلك لم تُقترح مطابقة لها.`]:[])];
 $('reportNarrative').innerHTML=`<ul>${entries.map(t=>`<li>${esc(t)}</li>`).join('')}</ul>`;
}
function renderAll(){renderCompanies();renderHome();renderReview();renderMatch();renderReport();renderEntry();$('ledgerCount').textContent=data().ledger.length?`تم استيراد ${number(data().ledger.length)} عملية • ${data().ledgerName}`:'ما تم استيراد أي ملف بعد.';$('bankCount').textContent=data().bank.length?`تم استيراد ${number(data().bank.length)} حركة • ${data().bankName}`:'ما تم استيراد أي كشف بعد.';$('currentCompanyLabel').textContent=company().name}
function goto(view){currentView=view;document.querySelectorAll('.view').forEach(el=>el.classList.toggle('active',el.id===`view-${view}`));document.querySelectorAll('.nav-btn').forEach(b=>b.classList.toggle('active',b.dataset.view===view));renderAll();window.scrollTo({top:0,behavior:'instant'})}
function changeCompany(id){if(!db[id]||id===activeCompany)return;activeCompany=id;resetEntryForm(false);pending=null;$('mappingPanel').classList.add('hidden');$('ledgerFile').value='';$('bankFile').value='';renderAll();toast(`تم اختيار ${company().name} — بيانات الشركات الأخرى منفصلة.`)}
async function readFile(file,kind){if(!file)return;if(file.size>10*1024*1024)throw Error('الملف أكبر من 10 ميجابايت، قسّمه لملفات أصغر.');const ext=file.name.split('.').pop().toLowerCase();let sheets;
 if(ext==='xlsx')sheets=await parseXlsx(await file.arrayBuffer());else if(['csv','tsv'].includes(ext)){const bytes=new Uint8Array(await file.arrayBuffer());const encoding=bytes[0]===0xff&&bytes[1]===0xfe?'utf-16le':'utf-8';const text=new TextDecoder(encoding).decode(bytes);sheets=[{name:'البيانات',rows:ext==='tsv'?parseDelimited(text,'\t'):detectDelimited(text)}]}else throw Error('صيغة الملف غير مدعومة. استخدم xlsx أو CSV أو TSV، وليس xls القديم.');
 pending={kind,name:file.name,sheets,sheet:0,company:activeCompany};setupMapping();toast('تم قراءة الملف محليًا. راجع الأعمدة قبل الاستيراد.');
}
function importPending(){if(!pending)throw Error('ما فيه ملف بانتظار الاستيراد.');if(pending.company!==activeCompany)throw Error('تغيّرت الشركة المختارة. أعد الاستيراد.');if(pending.rows.length>10000)throw Error('الحد الأقصى 10,000 عملية لكل ملف.');const records=makeRecords(pending),kind=pending.kind,d=data();
 const commit=()=>{d[kind]=records;d[kind+'Name']=pending.name;d.approved=[];d.demo=false;pending=null;$('mappingPanel').classList.add('hidden');$('ledgerFile').value='';$('bankFile').value='';renderAll();goto(kind==='ledger'?'review':'match');toast(`تم استيراد ${number(records.length)} صف إلى ${company().name} بنجاح.`)};
 if(d[kind].length)confirmDialog('استبدال بيانات الشركة؟',`هذا بيستبدل ${kind==='ledger'?'دفتر العمليات':'كشف البنك'} في ${company().name}، وبيُلغي المطابقات المعتمدة حاليًا. بيانات باقي الشركات ما تتأثر.`,commit);else commit();
}
function loadDemo(){const d=data();const commit=()=>{
 const ledgerRows=[['2026-09-01','شراء بضاعة','125','INV-101','مصروف','مورد ألف'],['2026-09-01','شراء بضاعة','125','INV-101','مصروف','مورد ألف'],['2026-09-02','إيراد مبيعات','450','SALE-11','إيراد','عميل باء'],['2026-09-03','فاتورة كهرباء','80','BILL-8','مصروف','الكهرباء'],['2026-09-04','مصروف نقل','65','BILL-9','مصروف','شركة نقل'],['','مشتريات بدون تاريخ','90','INV-103','مصروف','مورد جيم'],['2026-09-05','مصروف غير واضح','ليس رقمًا','INV-104','مصروف','مورد دال']];
 d.ledger=ledgerRows.map((r,i)=>({id:`ledger-${i+1}`,row:i+2,date:normalizeDate(r[0]),rawDate:r[0],description:r[1],amount:classifyDirection(r[4])*parseAmount(r[2]),rawAmount:r[2],reference:r[3],type:r[4],party:r[5],directionUnknown:false}));
 d.bank=[['2026-09-02','إيداع مبيعات',450,'SALE-11'],['2026-09-03','فاتورة كهرباء',-80,'BILL-8'],['2026-09-01','شراء بضاعة',-125,'INV-101'],['2026-09-05','عمولة بنك',-15,'BANK-1']].map((r,i)=>({id:`bank-${i+1}`,row:i+2,date:r[0],rawDate:r[0],description:r[1],amount:r[2],rawAmount:String(r[2]),reference:r[3],type:'',party:'',directionUnknown:false}));
 d.approved=[];d.demo=true;d.ledgerName='بيانات وهمية للتجربة';d.bankName='كشف وهمي للتجربة';renderAll();goto('review');toast('تمت إضافة مثال تجريبي للشركة الحالية فقط. ما فيه بيانات حقيقية.');};
 if(hasData())confirmDialog('استبدال بيانات الشركة التجريبية؟','بيانات الشركة المختارة الحالية بتُستبدل بمثال وهمي، والمطابقات المعتمدة بتنمسح. باقي الشركات ما تتأثر.',commit);else commit();
}
// Manual entry is an append-only-in-session helper, never a financial posting operation.
function todayLocal(){const d=new Date();return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`}
function setEntryDirection(kind,chosen){
 const direction=$('entryDirection');
 direction.innerHTML=(kind==='ledger'?[['in','قبض / داخل'],['out','صرف / خارج']]:[['in','إيداع / داخل'],['out','سحب / خارج']]).map(([val,label])=>`<option value="${val}">${label}</option>`).join('');
 direction.value=chosen==='out'?'out':'in';
}
function resetEntryForm(keepKind=false){
 const kind=keepKind?$('entryKind').value:'ledger';
 entryEditing=null;$('entryForm').reset();$('entryCompany').disabled=false;$('entryKind').disabled=false;
 $('entryCompany').value=activeCompany;$('entryKind').value=kind;
 $('entryDate').value=todayLocal();setEntryDirection(kind,'in');
 $('entryFormTitle').textContent='عملية جديدة';$('entrySubmit').textContent='＋ حفظ وإضافة عملية ثانية';
 $('entryError').textContent='';$('entryError').classList.add('hidden');
}
function entryError(message,field){$('entryError').textContent=message;$('entryError').classList.remove('hidden');field?.focus();}
function renderEntry(){
 const d=data();$('entryCompany').value=activeCompany;$('entrySelectedCompany').textContent=company().name;
 $('entryLedgerCount').textContent=number(d.ledger.length);$('entryBankCount').textContent=number(d.bank.length);
 const manual=[...d.ledger.filter(r=>r.manual).map(r=>({kind:'ledger',r})),...d.bank.filter(r=>r.manual).map(r=>({kind:'bank',r}))].sort((a,b)=>b.r.manualOrder-a.r.manualOrder);
 $('entryRecentCount').textContent=`${number(manual.length)} عملية`;
 $('entryRecent').innerHTML=manual.length?manual.slice(0,8).map(({kind,r})=>`<div class="entry-record">
   <span class="entry-record-icon ${r.amount<0?'negative':''}" aria-hidden="true">${r.amount<0?'−':'+'}</span>
   <div class="entry-record-info"><strong>${esc(r.description)}</strong><small>${kind==='ledger'?'دفتر العمليات':'كشف البنك'} · ${esc(r.date)} ${r.reference?' · '+esc(r.reference):''}</small></div>
   <div class="entry-record-end"><strong dir="ltr" class="${r.amount<0?'negative':''}">${r.amount>0?'+':''}${money(r.amount)} ر.ع</strong><div class="entry-record-actions"><button type="button" data-manual-edit="${esc(r.id)}" data-manual-kind="${kind}" aria-label="تعديل ${esc(r.description)}">تعديل</button><button type="button" data-manual-delete="${esc(r.id)}" data-manual-kind="${kind}" aria-label="حذف ${esc(r.description)}">حذف</button></div></div>
 </div>`).join(''):'<div class="entry-empty"><span aria-hidden="true">✎</span><strong>ما دخلت أي عملية يدويًا بعد</strong><p>أول عملية بتظهر هنا وتقدر تعدلها أو تحذفها.</p></div>';
}
function entryValues(){
 const kind=$('entryKind').value, date=$('entryDate').value, description=$('entryDescription').value.trim(), raw=$('entryAmount').value.trim(),direction=$('entryDirection').value;
 if(!['ledger','bank'].includes(kind))throw Error('حدد مكان التسجيل الصحيح.');
 if(!normalizeDate(date)) {entryError('حدد تاريخ صحيح للعملية.', $('entryDate'));return null}
 if(!description){entryError('اكتب بيان العملية عشان تقدر تراجعها بعدين.',$('entryDescription'));return null}
 const amount=parseAmount(raw);
 if(!Number.isFinite(amount)||amount<=0||amount>999999999999.999||Math.abs(amount*1000-Math.round(amount*1000))>0.00001){entryError('أدخل مبلغًا موجبًا صحيحًا، بحد أقصى ثلاث خانات بعد الفاصلة.',$('entryAmount'));return null}
 if(!['in','out'].includes(direction)){entryError('اختر نوع الحركة.',$('entryDirection'));return null}
 const sign=direction==='out'?-1:1;
 return {date,rawDate:date,description,amount:sign*Math.round(amount*1000)/1000,rawAmount:raw,reference:$('entryReference').value.trim(),party:$('entryParty').value.trim(),type:kind==='ledger'?(sign>0?'قبض':'صرف'):(sign>0?'إيداع':'سحب'),directionUnknown:false};
}
function duplicateManual(kind,record,editingId){
 return data()[kind].find(r=>r.id!==editingId&&Number.isFinite(r.amount)&&Math.abs(r.amount-record.amount)<.00001&&(
  record.reference&&norm(record.reference)===norm(r.reference)&&norm(record.party)===norm(r.party) ||
  r.date===record.date&&norm(r.description)===norm(record.description)
 ));
}
function invalidateApprovals(kind,id){const d=data(),before=d.approved.length;d.approved=d.approved.filter(key=>{const ids=key.split('|');return (kind==='ledger'?ids[0]:ids[1])!==id});return before-d.approved.length}
function addOrUpdateEntry(e){e.preventDefault();$('entryError').classList.add('hidden');
 const kind=$('entryKind').value;if($('entryCompany').value!==activeCompany)return entryError('الشركة تغيّرت. اختر الشركة من جديد.',$('entryCompany'));
 const record=entryValues();if(!record)return;
 const d=data(),editing=entryEditing&&entryEditing.company===activeCompany&&entryEditing.kind===kind?entryEditing:null;
 if(!editing&&d[kind].length>=10000)return entryError('وصلت للحد الأقصى 10,000 عملية. صدّر البيانات قبل بدء جلسة جديدة.');
 const duplicate=!d.demo&&duplicateManual(kind,record,editing?.id);
 const commit=()=>{
   const current=data();if(current!==d)return toast('الشركة تغيّرت. أعد المحاولة.',true);
   if(!editing&&current.demo){current.ledger=[];current.bank=[];current.approved=[];current.ledgerName='';current.bankName='';current.demo=false}
   let invalidated=0;
   if(editing){const index=current[kind].findIndex(r=>r.id===editing.id&&r.manual);if(index<0)return toast('العملية غير موجودة أو غير قابلة للتعديل.',true);
     const old=current[kind][index];invalidated=invalidateApprovals(kind,old.id);
     current[kind][index]={...old,...record};
   }else{
     const sequence=++manualSequence;
     current[kind].push({...record,id:`${kind}-manual-${sequence}`,row:Math.max(1,...current[kind].map(r=>Number(r.row)||1))+1,manual:true,manualOrder:sequence});
     if(!current[kind+'Name'])current[kind+'Name']='إدخال يدوي';
   }
   const didEdit=Boolean(editing);resetEntryForm(true);renderAll();$('entryDescription').focus();
   toast((didEdit?'تم تعديل العملية.':'تم حفظ العملية في '+company().name+'.')+(invalidated?' ألغينا المطابقة المرتبطة عشان تراجعها.':'')+' صدّر CSV قبل ما تطلع.');
 };
 if(!editing&&d.demo)confirmDialog('إلغاء البيانات التجريبية؟','عشان ما نخلط بياناتك الحقيقية بالمثال الوهمي، بنمسح بيانات التجربة من هذه الشركة فقط ونبدأ بأول عملية منك.',commit);
 else if(duplicate)confirmDialog('تكرار محتمل','فيه عملية موجودة بنفس المرجع والمبلغ أو بنفس التاريخ والبيان والمبلغ. متأكد تبغى تحفظها مرة ثانية؟',commit);
 else commit();
}
function editEntry(kind,id){const record=data()[kind]?.find(r=>r.id===id&&r.manual);if(!record)return toast('العملية ما موجودة، أو جاية من ملف ولا تتعدل هنا.',true);
 entryEditing={kind,id,company:activeCompany};$('entryCompany').value=activeCompany;$('entryCompany').disabled=true;$('entryKind').value=kind;$('entryKind').disabled=true;
 $('entryDate').value=record.date;$('entryDescription').value=record.description;$('entryAmount').value=money(Math.abs(record.amount)).replace(/,/g,'');
 $('entryReference').value=record.reference||'';$('entryParty').value=record.party||'';setEntryDirection(kind,record.amount<0?'out':'in');
 $('entryFormTitle').textContent='تعديل عملية مسجّلة';$('entrySubmit').textContent='✓ حفظ التعديل';
 $('entryError').classList.add('hidden');$('entryForm').scrollIntoView({behavior:'smooth',block:'start'});$('entryDescription').focus();
}
function deleteEntry(kind,id){const r=data()[kind]?.find(x=>x.id===id&&x.manual);if(!r)return toast('ما حصلنا العملية.',true);
 const selectedCompany=activeCompany;
 confirmDialog('حذف العملية اليدوية؟',`بتحذف «${r.description.slice(0,60)}» من ${company().name} داخل جلسة الأداة فقط. التعديل ما يأثر على ملف Excel الأصلي.`,()=>{
  if(activeCompany!==selectedCompany)return toast('الشركة تغيّرت. أعد المحاولة.',true);
  const d=data();d[kind]=d[kind].filter(x=>x.id!==id||!x.manual);const invalidated=invalidateApprovals(kind,id);
  if(entryEditing?.id===id&&entryEditing.kind===kind)resetEntryForm(true);
  renderAll();toast('تم حذف العملية.'+(invalidated?' وتم إلغاء المطابقة المرتبطة.':''));
 });
}
function exportFullData(kind){const d=data(),rows=d[kind];if(!rows.length)return toast('ما فيه عمليات لتصديرها من '+company().name+'.',true);
 downloadCsv([['التاريخ','البيان','المبلغ','رقم المستند','نوع العملية','الطرف','مصدر الإدخال'],...rows.map(r=>[r.date||r.rawDate,r.description,Number.isFinite(r.amount)?r.amount:r.rawAmount,r.reference,r.type,r.party,r.manual?'يدوي':'ملف Excel / CSV'])],kind==='ledger'?'ledger-all':'bank-all');
}
function safeCsvCell(value){let s=String(value??'');if(/^[\s\uFEFF]*[=+\-@\t\r]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"'}
function downloadCsv(rows,name){const content='\uFEFF'+rows.map(row=>row.map(v=>typeof v==='number'&&Number.isFinite(v)?String(v):safeCsvCell(v)).join(',')).join('\r\n');const blob=new Blob([content],{type:'text/csv;charset=utf-8'});const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=`sanad-${activeCompany}-${name}-${new Date().toISOString().slice(0,10)}.csv`;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),3000);toast('تم تجهيز ملف CSV. افتحه في Excel للمراجعة.')}
function exportReview(){const records=analyseLedger(data().ledger);if(!records.length)return toast('استورد دفتر العمليات أولًا.',true);downloadCsv([['رقم الصف','التاريخ','البيان','المرجع','الطرف','المبلغ','الملاحظات'],...records.filter(x=>x.issues.length).map(r=>[r.row,r.date||r.rawDate,r.description,r.reference,r.party,Number.isFinite(r.amount)?r.amount:r.rawAmount,r.issues.join(' | ')])],'review')}
function exportMatches(){const m=matchResult();if(!m.ledger.length&&!m.bank.length)return toast('ما فيه بيانات للتصدير.',true);const rows=[['المصدر','الصف','التاريخ','البيان','المرجع','المبلغ','الحالة','الصف المقابل'],...m.approved.flatMap(p=>[['الدفتر',p.l.row,p.l.date,p.l.description,p.l.reference,p.l.amount,'معتمدة',p.b.row],['البنك',p.b.row,p.b.date,p.b.description,p.b.reference,p.b.amount,'معتمدة',p.l.row]]),...m.suggested.map(p=>['الدفتر',p.l.row,p.l.date,p.l.description,p.l.reference,p.l.amount,'مطابقة مقترحة غير معتمدة',p.b.row]),...m.unmatchedLedger.filter(x=>!m.suggested.some(p=>p.l.id===x.id)).map(r=>['الدفتر',r.row,r.date,r.description,r.reference,Number.isFinite(r.amount)?r.amount:r.rawAmount,'غير معتمدة','']),...m.unmatchedBank.filter(x=>!m.suggested.some(p=>p.b.id===x.id)).map(r=>['البنك',r.row,r.date,r.description,r.reference,Number.isFinite(r.amount)?r.amount:r.rawAmount,'غير معتمدة',''])];downloadCsv(rows,'bank-match')}
function exportReport(){const r=getReport();if(!r.d.ledger.length&&!r.d.bank.length)return toast('استورد البيانات أو جرّب المثال أولًا.',true);const rows=[['القسم','المرجع','التاريخ','البيان','المبلغ','الحالة / الملاحظة'],['الشركة','', '',r.company,'',r.d.demo?'بيانات تجريبية':'بيانات الجلسة'],['الملخص','','','عمليات دفتر',r.d.ledger.length,''],['الملخص','','','حركات بنك',r.d.bank.length,''],['الملخص','','','عمليات بملاحظات',r.flags.length,''],['الملخص','','','مطابقات مقترحة',r.m.suggested.length,'غير معتمدة'],['الملخص','','','مطابقات معتمدة',r.m.approved.length,''],...r.flags.map(x=>['الملاحظات',x.reference,x.date||x.rawDate,x.description,Number.isFinite(x.amount)?x.amount:x.rawAmount,x.issues.join(' | ')]),...r.m.unmatchedBank.map(x=>['حركات بنك غير معتمدة',x.reference,x.date,x.description,Number.isFinite(x.amount)?x.amount:x.rawAmount,'بانتظار المراجعة'])];downloadCsv(rows,'summary')}
function bindEvents(){resetEntryForm(false);document.querySelectorAll('[data-view],[data-goto]').forEach(b=>b.addEventListener('click',()=>goto(b.dataset.view||b.dataset.goto)));$('companyCards').addEventListener('click',e=>{const card=e.target.closest('[data-company]');if(card)changeCompany(card.dataset.company)});
 $('entryForm').addEventListener('submit',addOrUpdateEntry);
 $('entryKind').addEventListener('change',()=>setEntryDirection($('entryKind').value,'in'));
 $('entryCompany').addEventListener('change',()=>changeCompany($('entryCompany').value));
 $('entryReset').addEventListener('click',()=>{resetEntryForm(true);renderEntry()});
 $('entryExportLedger').addEventListener('click',()=>exportFullData('ledger'));
 $('entryExportBank').addEventListener('click',()=>exportFullData('bank'));
 $('entryRecent').addEventListener('click',event=>{const edit=event.target.closest('[data-manual-edit]'),del=event.target.closest('[data-manual-delete]');if(edit)editEntry(edit.dataset.manualKind,edit.dataset.manualEdit);if(del)deleteEntry(del.dataset.manualKind,del.dataset.manualDelete)});
 $('demoBtn').addEventListener('click',loadDemo);$('wipeBtn').addEventListener('click',()=>confirmDialog('مسح بيانات الجلسة؟','هذا يمسح بيانات الشركات الأربع من ذاكرة الصفحة، بما فيها المطابقات المعتمدة. احفظ تقاريرك أولًا.',()=>{for(const c of COMPANIES)db[c.id]=makeData();pending=null;$('mappingPanel').classList.add('hidden');resetEntryForm(false);renderAll();goto('home');toast('تم مسح بيانات الجلسة.') }));
 $('modalCancel').addEventListener('click',closeModal);$('modalConfirm').addEventListener('click',()=>{const fn=modalCallback;closeModal();fn?.()});$('modal').addEventListener('click',e=>{if(e.target===$('modal'))closeModal()});document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('modal').classList.contains('hidden'))closeModal()});
 for(const kind of ['ledger','bank']){const input=$(`${kind}File`),drop=$(`${kind}Drop`);input.addEventListener('change',async()=>{try{await readFile(input.files[0],kind)}catch(e){toast(e.message||'تعذر قراءة الملف.',true)}});drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('dragover')});drop.addEventListener('dragleave',()=>drop.classList.remove('dragover'));drop.addEventListener('drop',async e=>{e.preventDefault();drop.classList.remove('dragover');try{await readFile(e.dataTransfer.files[0],kind)}catch(err){toast(err.message||'تعذر قراءة الملف.',true)}})}
 $('pasteToggle').addEventListener('click',()=>$('pasteBox').classList.toggle('hidden'));
 $('pasteRead').addEventListener('click',()=>{try{const rows=detectDelimited($('pasteInput').value);pending={kind:'ledger',name:'جدول ملصوق من Excel',sheets:[{name:'جدول ملصوق',rows}],sheet:0,company:activeCompany};setupMapping()}catch(e){toast(e.message,true)}});
 $('sheetSelect').addEventListener('change',()=>{if(!pending)return;pending.sheet=Number($('sheetSelect').value);try{setupMapping()}catch(e){toast(e.message,true)}});$('cancelMapping').addEventListener('click',()=>{pending=null;$('mappingPanel').classList.add('hidden')});$('confirmImport').addEventListener('click',()=>{try{importPending()}catch(e){toast(e.message,true)}});
 $('reviewFilter').addEventListener('change',renderReview);$('reviewSearch').addEventListener('input',renderReview);
 $('matchTable').addEventListener('click',e=>{const btn=e.target.closest('[data-pair]');if(!btn)return;const d=data(),k=btn.dataset.pair,result=matchResult();if(btn.dataset.pairAction==='approve'){if(!result.suggested.some(p=>p.key===k))return toast('المقترح تغيّر؛ أعد المراجعة.',true);d.approved.push(k)}else d.approved=d.approved.filter(x=>x!==k);renderAll();toast(btn.dataset.pairAction==='approve'?'تم اعتماد هذا الربط داخل الأداة فقط.':'تم إلغاء اعتماد الربط.')});
 $('exportReviewTop').addEventListener('click',exportReview);$('exportMatch').addEventListener('click',exportMatches);$('exportReport').addEventListener('click',exportReport);$('printReport').addEventListener('click',()=>{if(!hasData())return toast('ما فيه بيانات للطباعة.',true);renderReport();window.print()});
}
bindEvents();renderAll();goto('home');
