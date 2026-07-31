import { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { CLASSIFICATIONS, buildFinancialStatements, normalize, processTrialBalance, recalculate } from './services/trialBalanceEngine';

const money = value => new Intl.NumberFormat('ar-JO', { maximumFractionDigits: 2 }).format(value || 0);
const SAMPLE_ROWS = [
  { 'رقم الحساب': '101001', 'اسم الحساب': 'الصندوق', مدين: 15000, دائن: 0 },
  { 'رقم الحساب': '102001', 'اسم الحساب': 'البنك', مدين: 45000, دائن: 0 },
  { 'رقم الحساب': '103001', 'اسم الحساب': 'العملاء', مدين: 28000, دائن: 0 },
  { 'رقم الحساب': '104001', 'اسم الحساب': 'مخزون البضاعة', مدين: 32000, دائن: 0 },
  { 'رقم الحساب': '151001', 'اسم الحساب': 'الأثاث والمعدات', مدين: 20000, دائن: 0 },
  { 'رقم الحساب': '159001', 'اسم الحساب': 'مجمع إهلاك الأثاث', مدين: 0, دائن: 4000 },
  { 'رقم الحساب': '201001', 'اسم الحساب': 'الموردون', مدين: 0, دائن: 23000 },
  { 'رقم الحساب': '202001', 'اسم الحساب': 'مصروفات مستحقة', مدين: 0, دائن: 5000 },
  { 'رقم الحساب': '301001', 'اسم الحساب': 'رأس المال', مدين: 0, دائن: 60000 },
  { 'رقم الحساب': '401001', 'اسم الحساب': 'المبيعات', مدين: 0, دائن: 125000 },
  { 'رقم الحساب': '501001', 'اسم الحساب': 'تكلفة البضاعة المباعة', مدين: 70000, دائن: 0 },
  { 'رقم الحساب': '601001', 'اسم الحساب': 'مصروف الرواتب', مدين: 12000, دائن: 0 },
  { 'رقم الحساب': '602001', 'اسم الحساب': 'مصروف الإيجار', مدين: 6000, دائن: 0 },
  { 'رقم الحساب': '603001', 'اسم الحساب': 'مصروف الكهرباء', مدين: 2000, دائن: 0 },
  { 'رقم الحساب': '604001', 'اسم الحساب': 'مصروف الإهلاك', مدين: 4000, دائن: 0 },
  { 'رقم الحساب': '302001', 'اسم الحساب': 'أرباح محتجزة', مدين: 0, دائن: 4000 },
];
const EMPTY_ROWS = Array.from({ length: 5 }, () => ({ 'رقم الحساب': '', 'اسم الحساب': '', مدين: '', دائن: '' }));

function makeWorkbook(rows, fileName) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [{ wch: 16 }, { wch: 34 }, { wch: 16 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(wb, ws, 'ميزان المراجعة');
  const help = XLSX.utils.aoa_to_sheet([
    ['دليل تعبئة النموذج'], ['1', 'اكتب كل حساب في صف مستقل.'], ['2', 'لا تدمج الخلايا داخل الجدول.'],
    ['3', 'أدخل الرصيد في المدين أو الدائن.'], ['4', 'تأكد من تساوي الإجماليين.'], ['5', 'يمكن استخدام العربية أو الإنجليزية.'],
  ]);
  help['!cols'] = [{ wch: 8 }, { wch: 75 }];
  XLSX.utils.book_append_sheet(wb, help, 'اقرأني أولًا');
  XLSX.writeFile(wb, fileName);
}

export default function App() {
  const [view, setView] = useState('home');
  const [company, setCompany] = useState('شركة تجريبية');
  const [result, setResult] = useState(null);
  const [error, setError] = useState('');
  const [query, setQuery] = useState('');
  const [activeTab, setActiveTab] = useState('review');
  const [tour, setTour] = useState(() => localStorage.getItem('statmanets-tour-done') ? false : true);

  const learned = useMemo(() => {
    try { return JSON.parse(localStorage.getItem('statmanets-learned') || '{}'); } catch { return {}; }
  }, [result]);

  const statements = useMemo(() => result ? buildFinancialStatements(result.accounts) : null, [result]);
  const filtered = useMemo(() => !result ? [] : result.accounts.filter(a => `${a.name} ${a.code} ${a.label}`.toLowerCase().includes(query.toLowerCase())), [result, query]);

  function loadRows(rows, name = company) {
    try {
      const processed = processTrialBalance(rows, learned);
      setResult(processed); setCompany(name); setError(''); setView('workspace'); setActiveTab('review');
    } catch (e) { setError(e.message || 'تعذر تحليل الملف.'); }
  }

  async function handleFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
      loadRows(rows);
    } catch (e) { setError(e.message || 'حدث خطأ أثناء قراءة الملف.'); }
  }

  function changeClassification(accountId, section) {
    const option = CLASSIFICATIONS.find(x => x.section === section);
    if (!option || !result) return;
    const accounts = result.accounts.map(a => a.id === accountId ? { ...a, ...option, confidence: 1, source: 'manual' } : a);
    const account = accounts.find(a => a.id === accountId);
    const memory = { ...learned, [normalize(`${account.code}|${account.name}`)]: section };
    localStorage.setItem('statmanets-learned', JSON.stringify(memory));
    setResult(recalculate(accounts, result.columns));
  }

  function exportStatements() {
    if (!statements) return;
    const wb = XLSX.utils.book_new();
    const income = statements.incomeStatement;
    const incomeRows = [
      ['قائمة الدخل', company], ['البند', 'المبلغ'], ['الإيرادات', income.revenue], ['تكلفة الإيرادات', income.cost],
      ['مجمل الربح', income.grossProfit], ['المصروفات التشغيلية', income.expenses], ['بنود أخرى', income.other], ['صافي الربح', income.netProfit],
    ];
    const fp = statements.financialPosition;
    const positionRows = [['قائمة المركز المالي', company], ['الأصول', 'المبلغ'], ...fp.assetLines.map(x => [x.label, x.amount]), ['إجمالي الأصول', fp.assets], [], ['الالتزامات', 'المبلغ'], ...fp.liabilityLines.map(x => [x.label, x.amount]), ['إجمالي الالتزامات', fp.liabilities], [], ['حقوق الملكية', 'المبلغ'], ...fp.equityLines.map(x => [x.label, x.amount]), ['صافي ربح الفترة', income.netProfit], ['إجمالي حقوق الملكية', fp.equityWithProfit], ['فرق المعادلة', fp.difference]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(incomeRows), 'قائمة الدخل');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(positionRows), 'المركز المالي');
    XLSX.writeFile(wb, `القوائم-المالية-${company}.xlsx`);
  }

  return <main className="app-shell">
    <nav className="topbar">
      <button className="brand" onClick={() => setView('home')}><span>S</span><b>Statmanets</b></button>
      <div className="nav-actions">
        <button onClick={() => setView('home')}>الرئيسية</button>
        <button onClick={() => setView('workspace')}>مساحة العمل</button>
        <button onClick={() => setView('guide')}>دليل الاستخدام</button>
      </div>
    </nav>

    {view === 'home' && <>
      <header className="hero landing-hero">
        <div><span className="eyebrow">SMART FINANCIAL STATEMENTS</span><h1>من ميزان المراجعة إلى قوائم مالية قابلة للمراجعة</h1><p>ارفع Excel، صحح التصنيفات عند الحاجة، ثم أنشئ قائمة الدخل والمركز المالي واحفظهما في ملف جديد.</p>
          <div className="hero-buttons"><button className="primary" onClick={() => setView('workspace')}>ابدأ بملفي</button><button className="secondary" onClick={() => loadRows(SAMPLE_ROWS, 'شركة الأفق التجارية')}>تجربة مباشرة</button><button className="ghost" onClick={() => setView('guide')}>دليل الاستخدام</button></div>
        </div>
        <div className="hero-panel"><b>Excel</b><span>↓ اكتشاف الأعمدة</span><span>↓ تصنيف ومراجعة</span><strong>قائمة الدخل + المركز المالي</strong></div>
      </header>
      <section className="steps-section"><div className="section-head"><span>رحلة واضحة</span><h2>أربع مراحل حتى القوائم</h2></div><div className="steps-grid">
        {['رفع الملف','فحص التوازن','مراجعة التصنيف','إنشاء القوائم'].map((x,i)=><article key={x}><i>0{i+1}</i><h3>{x}</h3><p>{['يدعم أسماء أعمدة عربية وإنجليزية.','يعرض المدين والدائن والفرق مباشرة.','يمكن تعديل كل حساب وحفظ القرار.','عرض وتصدير النتائج إلى Excel.'][i]}</p></article>)}
      </div></section>
      <section className="sample-section"><div><h2>ابدأ بملف جاهز</h2><p>اختبر النظام أو نزّل نموذجًا فارغًا لتعبئته.</p></div><div className="sample-actions"><button className="primary" onClick={() => loadRows(SAMPLE_ROWS, 'شركة الأفق التجارية')}>تشغيل التجربة</button><button className="secondary" onClick={() => makeWorkbook(SAMPLE_ROWS,'ميزان-مراجعة-تجريبي.xlsx')}>تنزيل التجريبي</button><button className="secondary" onClick={() => makeWorkbook(EMPTY_ROWS,'نموذج-ميزان-مراجعة.xlsx')}>تنزيل نموذج فارغ</button></div></section>
    </>}

    {view === 'guide' && <section className="guide-page"><div className="section-head"><span>MANUAL</span><h1>دليل الاستخدام</h1><p>1) جهّز ملف Excel. 2) ارفع الملف. 3) راجع الأعمدة والتوازن. 4) صحح الحسابات منخفضة الثقة. 5) افتح تبويب القوائم وصدّر النتيجة.</p></div><div className="guide-grid">
      <article><h3>الملف المطلوب</h3><p>اسم الحساب إلزامي، ومعه مدين ودائن أو رصيد صافٍ. ضع كل حساب في صف مستقل وتجنب الخلايا المدمجة.</p></article>
      <article><h3>التصنيف اليدوي</h3><p>اختر التصنيف الصحيح من القائمة بجانب الحساب. يحفظ التطبيق القرار محليًا ويستخدمه عند رفع الحساب نفسه لاحقًا.</p></article>
      <article><h3>القوائم المالية</h3><p>قائمة الدخل والمركز المالي تتغيران مباشرة بعد كل تعديل، مع إظهار فرق المعادلة المحاسبية والحسابات غير المصنفة.</p></article>
      <article><h3>الخصوصية</h3><p>تتم القراءة والحساب داخل الجهاز في هذه النسخة ولا يحتاج تحليل الملف إلى خادم خارجي.</p></article>
    </div></section>}

    {view === 'workspace' && <section className="workspace">
      <aside className="upload-card"><label>اسم الشركة</label><input value={company} onChange={e=>setCompany(e.target.value)}/><label className="drop-zone"><strong>رفع ميزان المراجعة</strong><span>XLSX أو XLS</span><input type="file" accept=".xlsx,.xls" onChange={handleFile}/></label><button className="primary wide" onClick={() => loadRows(SAMPLE_ROWS,'شركة الأفق التجارية')}>تشغيل الملف التجريبي</button><button className="secondary wide" onClick={() => makeWorkbook(EMPTY_ROWS,'نموذج-ميزان-مراجعة.xlsx')}>تنزيل النموذج</button>{error&&<div className="error">{error}</div>}</aside>
      <section className="content-card">{!result ? <div className="empty-state"><div className="empty-icon">↥</div><h2>ارفع ملفًا أو شغّل التجربة</h2><p>سيظهر فحص الميزان والتصنيف والقوائم هنا.</p></div> : <>
        <div className="report-head"><div><small>ملف الشركة</small><h2>{company}</h2></div><span className={result.balanced?'status ok':'status warn'}>{result.balanced?'الميزان متوازن':'يوجد فرق في الميزان'}</span></div>
        <div className="kpis"><article><small>إجمالي المدين</small><b>{money(result.totals.debit)}</b></article><article><small>إجمالي الدائن</small><b>{money(result.totals.credit)}</b></article><article><small>الحسابات</small><b>{result.accounts.length}</b></article><article><small>غير مصنف</small><b>{statements.unmapped.length}</b></article></div>
        <div className="tabs"><button className={activeTab==='review'?'active':''} onClick={()=>setActiveTab('review')}>مراجعة الحسابات</button><button className={activeTab==='statements'?'active':''} onClick={()=>setActiveTab('statements')}>القوائم المالية</button></div>
        {activeTab==='review' ? <><div className="toolbar"><input placeholder="بحث عن حساب..." value={query} onChange={e=>setQuery(e.target.value)}/><span>{filtered.length} حساب</span></div><div className="table-wrap"><table><thead><tr><th>الحساب</th><th>التصنيف</th><th>الرصيد</th><th>الثقة</th></tr></thead><tbody>{filtered.map(a=><tr key={a.id} className={a.section==='unmapped'?'needs-review':''}><td><b>{a.name}</b><small>{a.code}</small></td><td><select value={a.section} onChange={e=>changeClassification(a.id,e.target.value)}>{CLASSIFICATIONS.map(x=><option key={x.section} value={x.section}>{x.label}</option>)}</select></td><td>{money(a.balance)}</td><td><span className={a.confidence>=.7?'confidence high':'confidence low'}>{Math.round(a.confidence*100)}%</span></td></tr>)}</tbody></table></div></> : <Statements statements={statements} company={company} exportStatements={exportStatements}/>} 
      </>}</section>
    </section>}

    {tour && <div className="tour-overlay"><div className="tour-card"><span className="tour-badge">مرحبًا</span><h2>جرّب التطبيق خلال دقيقة</h2><p>شغّل الملف التجريبي، راجع التصنيفات، ثم افتح القوائم المالية وشاهد النتيجة.</p><div><button className="secondary" onClick={()=>{localStorage.setItem('statmanets-tour-done','1');setTour(false)}}>إغلاق</button><button className="primary" onClick={()=>{localStorage.setItem('statmanets-tour-done','1');setTour(false);loadRows(SAMPLE_ROWS,'شركة الأفق التجارية')}}>ابدأ التجربة</button></div></div></div>}
  </main>;
}

function Statements({ statements, company, exportStatements }) {
  const income = statements.incomeStatement; const fp = statements.financialPosition;
  return <div className="statements"><div className="statement-actions"><div><small>القوائم المولدة</small><h3>{company}</h3></div><button className="primary" onClick={exportStatements}>تصدير إلى Excel</button></div>
    {statements.unmapped.length>0&&<div className="warning-box">يوجد {statements.unmapped.length} حساب غير مصنف. راجعه للحصول على قوائم أدق.</div>}
    <div className="statement-grid"><article className="statement-card"><h3>قائمة الدخل</h3><Line label="الإيرادات" value={income.revenue}/><Line label="تكلفة الإيرادات" value={income.cost}/><Line label="مجمل الربح" value={income.grossProfit} total/><Line label="المصروفات التشغيلية" value={income.expenses}/><Line label="بنود أخرى" value={income.other}/><Line label="صافي الربح" value={income.netProfit} grand/></article>
    <article className="statement-card"><h3>قائمة المركز المالي</h3><h4>الأصول</h4>{fp.assetLines.map(x=><Line key={x.section} label={x.label} value={x.amount}/>)}<Line label="إجمالي الأصول" value={fp.assets} total/><h4>الالتزامات</h4>{fp.liabilityLines.map(x=><Line key={x.section} label={x.label} value={x.amount}/>)}<Line label="إجمالي الالتزامات" value={fp.liabilities} total/><h4>حقوق الملكية</h4>{fp.equityLines.map(x=><Line key={x.section} label={x.label} value={x.amount}/>)}<Line label="صافي ربح الفترة" value={income.netProfit}/><Line label="إجمالي حقوق الملكية" value={fp.equityWithProfit} total/><Line label="فرق المعادلة" value={fp.difference} grand/></article></div>
  </div>;
}
function Line({label,value,total,grand}) { return <div className={`statement-line ${total?'total':''} ${grand?'grand':''}`}><span>{label}</span><b>{money(value)}</b></div>; }
