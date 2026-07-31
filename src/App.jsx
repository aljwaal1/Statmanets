import { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { CLASSIFICATIONS, buildFinancialStatements, detectColumns, normalize, processTrialBalance, recalculate } from './services/trialBalanceEngine';

const money=v=>new Intl.NumberFormat('ar-JO',{maximumFractionDigits:2}).format(v||0);
const SAMPLE_ROWS=[
 {'رقم الحساب':'101001','اسم الحساب':'الصندوق',مدين:15000,دائن:0},{'رقم الحساب':'102001','اسم الحساب':'البنك',مدين:45000,دائن:0},
 {'رقم الحساب':'103001','اسم الحساب':'العملاء',مدين:28000,دائن:0},{'رقم الحساب':'104001','اسم الحساب':'مخزون البضاعة',مدين:32000,دائن:0},
 {'رقم الحساب':'151001','اسم الحساب':'الأثاث والمعدات',مدين:20000,دائن:0},{'رقم الحساب':'159001','اسم الحساب':'مجمع إهلاك الأثاث',مدين:0,دائن:4000},
 {'رقم الحساب':'201001','اسم الحساب':'الموردون',مدين:0,دائن:23000},{'رقم الحساب':'202001','اسم الحساب':'مصروفات مستحقة',مدين:0,دائن:5000},
 {'رقم الحساب':'301001','اسم الحساب':'رأس المال',مدين:0,دائن:60000},{'رقم الحساب':'302001','اسم الحساب':'أرباح محتجزة',مدين:0,دائن:4000},
 {'رقم الحساب':'401001','اسم الحساب':'المبيعات',مدين:0,دائن:125000},{'رقم الحساب':'501001','اسم الحساب':'تكلفة البضاعة المباعة',مدين:70000,دائن:0},
 {'رقم الحساب':'601001','اسم الحساب':'مصروف الرواتب',مدين:12000,دائن:0},{'رقم الحساب':'602001','اسم الحساب':'مصروف الإيجار',مدين:6000,دائن:0},
 {'رقم الحساب':'603001','اسم الحساب':'مصروف الكهرباء',مدين:2000,دائن:0},{'رقم الحساب':'604001','اسم الحساب':'مصروف الإهلاك',مدين:4000,دائن:0}
];

const loadJSON=(key,fallback)=>{try{return JSON.parse(localStorage.getItem(key)||JSON.stringify(fallback));}catch{return fallback;}};
const saveCompanies=value=>localStorage.setItem('statmanets-companies',JSON.stringify(value));

export default function App(){
 const [view,setView]=useState('home');
 const [companies,setCompanies]=useState(()=>loadJSON('statmanets-companies',[]));
 const [companyId,setCompanyId]=useState('');
 const [company,setCompany]=useState('شركة جديدة');
 const [result,setResult]=useState(null);
 const [error,setError]=useState('');
 const [rawRows,setRawRows]=useState(null);
 const [headers,setHeaders]=useState([]);
 const [mapping,setMapping]=useState({code:'',name:'',debit:'',credit:'',balance:''});
 const [showMapping,setShowMapping]=useState(false);
 const [activeTab,setActiveTab]=useState('review');
 const [query,setQuery]=useState('');

 const learned=useMemo(()=>loadJSON(`statmanets-learned-${companyId||'default'}`,{}),[companyId,result]);
 const statements=useMemo(()=>result?buildFinancialStatements(result.accounts):null,[result]);
 const filtered=useMemo(()=>result?result.accounts.filter(a=>normalize(`${a.name} ${a.code} ${a.label}`).includes(normalize(query))):[],[result,query]);

 function persistCompany(nextResult=result,nextName=company){
  if(!nextResult)return;
  const id=companyId||`c-${Date.now()}`;
  const record={id,name:nextName,updatedAt:new Date().toISOString(),result:nextResult};
  const next=[record,...companies.filter(c=>c.id!==id)];
  setCompanyId(id);setCompanies(next);saveCompanies(next);
 }

 function loadRows(rows,name=company,columns=null){
  try{const processed=processTrialBalance(rows,learned,columns);setResult(processed);setCompany(name);setError('');setShowMapping(false);setView('workspace');setActiveTab('review');setTimeout(()=>persistCompany(processed,name),0);}catch(e){setError(e.message||'تعذر تحليل الملف.');}
 }

 async function handleFile(event){
  const file=event.target.files?.[0];if(!file)return;
  try{
   const wb=XLSX.read(await file.arrayBuffer(),{type:'array'});const sheet=wb.Sheets[wb.SheetNames[0]];
   const rows=XLSX.utils.sheet_to_json(sheet,{defval:''});if(!rows.length)throw new Error('الملف لا يحتوي على بيانات.');
   const detected=detectColumns(rows);setRawRows(rows);setHeaders(Object.keys(rows[0]||{}));setMapping({...detected});
   if(detected.name&&(detected.balance||detected.debit||detected.credit)) loadRows(rows); else setShowMapping(true);
  }catch(e){setError(e.message||'حدث خطأ أثناء قراءة الملف.');}
 }

 function applyMapping(){
  if(!rawRows)return;if(!mapping.name){setError('اختر عمود اسم الحساب.');return;}
  if(!mapping.balance&&!mapping.debit&&!mapping.credit){setError('اختر الرصيد أو المدين والدائن.');return;}
  loadRows(rawRows,company,mapping);
 }

 function changeClassification(id,section){
  const option=CLASSIFICATIONS.find(x=>x.section===section);if(!option||!result)return;
  const accounts=result.accounts.map(a=>a.id===id?{...a,...option,confidence:1,source:'manual'}:a);
  const account=accounts.find(a=>a.id===id);const memory={...learned,[normalize(`${account.code}|${account.name}`)]:section};
  localStorage.setItem(`statmanets-learned-${companyId||'default'}`,JSON.stringify(memory));
  const next=recalculate(accounts,result.columns);setResult(next);persistCompany(next);
 }

 function openCompany(id){const c=companies.find(x=>x.id===id);if(!c)return;setCompanyId(c.id);setCompany(c.name);setResult(c.result);setView('workspace');setActiveTab('review');}
 function newCompany(){setCompanyId('');setCompany('شركة جديدة');setResult(null);setRawRows(null);setError('');setView('workspace');}
 function deleteCompany(id){const next=companies.filter(c=>c.id!==id);setCompanies(next);saveCompanies(next);if(companyId===id)newCompany();}
 function renameCompany(name){setCompany(name);if(result)persistCompany(result,name);}

 function exportStatements(){
  if(!statements)return;const wb=XLSX.utils.book_new(),i=statements.incomeStatement,p=statements.financialPosition;
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['قائمة الدخل',company],['البند','المبلغ'],['الإيرادات',i.revenue],['تكلفة الإيرادات',i.cost],['مجمل الربح',i.grossProfit],['المصروفات التشغيلية',i.expenses],['بنود أخرى',i.other],['صافي الربح',i.netProfit]]),'قائمة الدخل');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['قائمة المركز المالي',company],['الأصول','المبلغ'],...p.assetLines.map(x=>[x.label,x.amount]),['إجمالي الأصول',p.assets],[],['الالتزامات','المبلغ'],...p.liabilityLines.map(x=>[x.label,x.amount]),['إجمالي الالتزامات',p.liabilities],[],['حقوق الملكية','المبلغ'],...p.equityLines.map(x=>[x.label,x.amount]),['صافي ربح الفترة',i.netProfit],['إجمالي حقوق الملكية',p.equityWithProfit],['فرق المعادلة',p.difference]]),'المركز المالي');
  XLSX.writeFile(wb,`القوائم-المالية-${company}.xlsx`);
 }

 const MappingSelect=({field,label,optional=false})=><label className="mapping-field"><span>{label}{optional?' (اختياري)':''}</span><select value={mapping[field]||''} onChange={e=>setMapping({...mapping,[field]:e.target.value})}><option value="">— غير محدد —</option>{headers.map(h=><option key={h} value={h}>{h}</option>)}</select></label>;

 return <main className="app-shell">
  <nav className="topbar"><button className="brand" onClick={()=>setView('home')}><span>S</span><b>Statmanets</b></button><div className="nav-actions"><button onClick={()=>setView('home')}>الرئيسية</button><button onClick={newCompany}>شركة جديدة</button><button onClick={()=>setView('companies')}>الشركات</button><button onClick={()=>setView('guide')}>الدليل</button></div></nav>

  {view==='home'&&<><header className="hero"><div><span className="eyebrow">SMART FINANCIAL STATEMENTS</span><h1>أنشئ القوائم المالية من أي ميزان مراجعة</h1><p>حتى لو اختلفت أسماء الأعمدة والحسابات، يمكنك مطابقتها وتصحيح التصنيف وحفظ كل شركة بصورة مستقلة.</p><div className="hero-buttons"><button className="primary" onClick={newCompany}>ابدأ الآن</button><button className="secondary" onClick={()=>loadRows(SAMPLE_ROWS,'شركة الأفق التجارية')}>تجربة مباشرة</button><button className="ghost" onClick={()=>setView('companies')}>شركاتي</button></div></div><div className="hero-panel"><b>Excel مختلف؟</b><span>مطابقة أعمدة مرنة</span><span>تصنيف قابل للتعديل</span><strong>حفظ تلقائي لكل شركة</strong></div></header><section className="steps-section"><div className="section-head"><span>المسار الكامل</span><h2>رفع، مطابقة، مراجعة، إصدار</h2></div><div className="steps-grid">{[['01','ارفع الملف'],['02','طابق الأعمدة'],['03','راجع الحسابات'],['04','احفظ وصدّر']].map(x=><article key={x[0]}><i>{x[0]}</i><h3>{x[1]}</h3></article>)}</div></section></>}

  {view==='companies'&&<section className="guide-page"><div className="section-head"><span>COMPANIES</span><h1>ملفات الشركات</h1><p>كل شركة تحتفظ بميزانها وتصنيفاتها ونتائجها على هذا الجهاز.</p></div><button className="primary" onClick={newCompany}>+ إضافة شركة</button><div className="company-grid">{companies.length?companies.map(c=><article className="company-card" key={c.id}><div><small>آخر تحديث</small><b>{c.name}</b><span>{new Date(c.updatedAt).toLocaleString('ar-JO')}</span></div><div><button className="secondary" onClick={()=>openCompany(c.id)}>فتح</button><button className="danger" onClick={()=>deleteCompany(c.id)}>حذف</button></div></article>):<div className="empty-company">لا توجد شركات محفوظة بعد.</div>}</div></section>}

  {view==='guide'&&<section className="guide-page"><div className="section-head"><span>MANUAL</span><h1>دليل الاستخدام</h1></div><div className="guide-grid"><article><h3>1. اختر الشركة</h3><p>أنشئ شركة جديدة أو افتح شركة محفوظة.</p></article><article><h3>2. ارفع Excel</h3><p>سيحاول التطبيق اكتشاف الأعمدة، وإن لم ينجح ستظهر شاشة المطابقة.</p></article><article><h3>3. راجع الحسابات</h3><p>عدّل التصنيف من القائمة وسيحفظ القرار للشركة.</p></article><article><h3>4. صدّر القوائم</h3><p>افتح تبويب القوائم ثم نزّل ملف Excel.</p></article></div></section>}

  {view==='workspace'&&<section className="workspace"><aside className="upload-card"><label>اسم الشركة</label><input value={company} onChange={e=>renameCompany(e.target.value)}/><label className="drop-zone"><strong>رفع ميزان المراجعة</strong><span>XLSX أو XLS</span><input type="file" accept=".xlsx,.xls" onChange={handleFile}/></label><button className="primary wide" onClick={()=>loadRows(SAMPLE_ROWS,'شركة الأفق التجارية')}>تشغيل الملف التجريبي</button><button className="secondary wide" onClick={()=>setView('companies')}>عرض الشركات المحفوظة</button>{error&&<div className="error">{error}</div>}</aside><section className="content-card">
   {!result?<div className="empty-state"><div className="empty-icon">↥</div><h2>ارفع ميزان المراجعة</h2><p>إن لم يكتشف التطبيق الأعمدة، ستتمكن من تحديدها يدويًا.</p></div>:<><div className="report-head"><div><small>ملف الشركة</small><h2>{company}</h2></div><span className={result.balanced?'status ok':'status warn'}>{result.balanced?'الميزان متوازن':'يوجد فرق في الميزان'}</span></div><div className="kpis"><article><small>إجمالي المدين</small><b>{money(result.totals.debit)}</b></article><article><small>إجمالي الدائن</small><b>{money(result.totals.credit)}</b></article><article><small>الحسابات</small><b>{result.accounts.length}</b></article><article><small>غير مصنف</small><b>{statements.unmapped.length}</b></article></div><div className="tabs"><button className={activeTab==='review'?'active':''} onClick={()=>setActiveTab('review')}>مراجعة الحسابات</button><button className={activeTab==='statements'?'active':''} onClick={()=>setActiveTab('statements')}>القوائم المالية</button></div>
    {activeTab==='review'?<><div className="toolbar"><input placeholder="ابحث باسم أو رقم الحساب" value={query} onChange={e=>setQuery(e.target.value)}/><button className="secondary" onClick={()=>setShowMapping(true)}>تعديل مطابقة الأعمدة</button></div><div className="table-wrap"><table><thead><tr><th>الحساب</th><th>الرصيد</th><th>التصنيف</th><th>الثقة</th></tr></thead><tbody>{filtered.map(a=><tr key={a.id} className={a.section==='unmapped'?'needs-review':''}><td><b>{a.name}</b><small>{a.code}</small></td><td>{money(a.balance)}</td><td><select value={a.section} onChange={e=>changeClassification(a.id,e.target.value)}>{CLASSIFICATIONS.map(c=><option key={c.section} value={c.section}>{c.label}</option>)}</select></td><td><span className={a.confidence>=.7?'confidence high':'confidence low'}>{Math.round(a.confidence*100)}%</span></td></tr>)}</tbody></table></div></>:<div className="statements"><div className="statement-actions"><h3>القوائم المالية</h3><button className="primary" onClick={exportStatements}>تصدير إلى Excel</button></div>{Math.abs(statements.financialPosition.difference)>.01&&<div className="warning-box">فرق المعادلة: {money(statements.financialPosition.difference)} — راجع الحسابات غير المصنفة أو التصنيفات غير الصحيحة.</div>}<div className="statement-grid"><article className="statement-card"><h3>قائمة الدخل</h3>{[['الإيرادات',statements.incomeStatement.revenue],['تكلفة الإيرادات',statements.incomeStatement.cost],['مجمل الربح',statements.incomeStatement.grossProfit],['المصروفات التشغيلية',statements.incomeStatement.expenses],['صافي الربح',statements.incomeStatement.netProfit]].map((x,i)=><div className={`statement-line ${i===4?'grand':''}`} key={x[0]}><span>{x[0]}</span><b>{money(x[1])}</b></div>)}</article><article className="statement-card"><h3>المركز المالي</h3><h4>الأصول</h4>{statements.financialPosition.assetLines.map(x=><div className="statement-line" key={x.section}><span>{x.label}</span><b>{money(x.amount)}</b></div>)}<div className="statement-line total"><span>إجمالي الأصول</span><b>{money(statements.financialPosition.assets)}</b></div><h4>الالتزامات وحقوق الملكية</h4>{[...statements.financialPosition.liabilityLines,...statements.financialPosition.equityLines].map(x=><div className="statement-line" key={x.section}><span>{x.label}</span><b>{money(x.amount)}</b></div>)}<div className="statement-line grand"><span>الإجمالي</span><b>{money(statements.financialPosition.liabilities+statements.financialPosition.equityWithProfit)}</b></div></article></div></div>}
   </>}
  </section></section>}

  {showMapping&&<div className="tour-overlay"><div className="mapping-card"><span className="tour-badge">مطابقة الأعمدة</span><h2>حدد وظيفة كل عمود</h2><p>اسم الحساب إلزامي. استخدم الرصيد الصافي، أو المدين والدائن.</p><div className="mapping-grid"><MappingSelect field="code" label="رقم الحساب" optional/><MappingSelect field="name" label="اسم الحساب"/><MappingSelect field="debit" label="مدين" optional/><MappingSelect field="credit" label="دائن" optional/><MappingSelect field="balance" label="الرصيد الصافي" optional/></div><div className="modal-actions"><button className="secondary" onClick={()=>setShowMapping(false)}>إلغاء</button><button className="primary" onClick={applyMapping}>اعتماد المطابقة</button></div></div></div>}
 </main>;
}
