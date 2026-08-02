import { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { CLASSIFICATIONS, buildFinancialStatements, detectColumns, normalize, processTrialBalance, recalculate } from './services/trialBalanceEngine';

const money = value => new Intl.NumberFormat('ar-JO', { maximumFractionDigits: 2 }).format(value || 0);
const loadJSON = (key, fallback) => { try { return JSON.parse(localStorage.getItem(key) || JSON.stringify(fallback)); } catch { return fallback; } };
const saveCompanies = value => localStorage.setItem('statmanets-companies', JSON.stringify(value));

const SAMPLE_ROWS = [
  { 'رقم الحساب': '101001', 'اسم الحساب': 'الصندوق', مدين: 15000, دائن: 0 },
  { 'رقم الحساب': '102001', 'اسم الحساب': 'البنك', مدين: 45000, دائن: 0 },
  { 'رقم الحساب': '103001', 'اسم الحساب': 'العملاء', مدين: 28000, دائن: 0 },
  { 'رقم الحساب': '104001', 'اسم الحساب': 'مخزون البضاعة', مدين: 32000, دائن: 0 },
  { 'رقم الحساب': '151001', 'اسم الحساب': 'الأثاث والمعدات', مدين: 20000, دائن: 0 },
  { 'رقم الحساب': '201001', 'اسم الحساب': 'الموردون', مدين: 0, دائن: 23000 },
  { 'رقم الحساب': '301001', 'اسم الحساب': 'رأس المال', مدين: 0, دائن: 60000 },
  { 'رقم الحساب': '401001', 'اسم الحساب': 'المبيعات', مدين: 0, دائن: 125000 },
  { 'رقم الحساب': '501001', 'اسم الحساب': 'تكلفة البضاعة المباعة', مدين: 70000, دائن: 0 },
  { 'رقم الحساب': '601001', 'اسم الحساب': 'مصروف الرواتب', مدين: 12000, دائن: 0 },
  { 'رقم الحساب': '602001', 'اسم الحساب': 'مصروف الإيجار', مدين: 6000, دائن: 0 }
];

export default function App() {
  const [view, setView] = useState('home');
  const [companies, setCompanies] = useState(() => loadJSON('statmanets-companies', []));
  const [companyId, setCompanyId] = useState('');
  const [company, setCompany] = useState('شركة جديدة');
  const [result, setResult] = useState(null);
  const [error, setError] = useState('');
  const [rawRows, setRawRows] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [mapping, setMapping] = useState({ code: '', name: '', debit: '', credit: '', balance: '' });
  const [showMapping, setShowMapping] = useState(false);
  const [activeTab, setActiveTab] = useState('review');
  const [query, setQuery] = useState('');

  const learned = useMemo(() => loadJSON(`statmanets-learned-${companyId || 'default'}`, {}), [companyId, result]);
  const statements = useMemo(() => result ? buildFinancialStatements(result.accounts) : null, [result]);
  const filtered = useMemo(() => result ? result.accounts.filter(a => normalize(`${a.name} ${a.code} ${a.label}`).includes(normalize(query))) : [], [result, query]);

  function persistCompany(nextResult = result, nextName = company) {
    if (!nextResult) return;
    const id = companyId || `c-${Date.now()}`;
    const record = { id, name: nextName, updatedAt: new Date().toISOString(), result: nextResult };
    const next = [record, ...companies.filter(c => c.id !== id)];
    setCompanyId(id); setCompanies(next); saveCompanies(next);
  }

  function loadRows(rows, name = company, columns = null) {
    try {
      const processed = processTrialBalance(rows, learned, columns);
      const invalidNames = processed.accounts.filter(a => !a.name || normalize(a.name) === normalize(a.code));
      if (invalidNames.length) throw new Error('عمود اسم الحساب غير صحيح. اختر عمود الاسم من شاشة المطابقة.');
      setResult(processed); setCompany(name); setError(''); setShowMapping(false); setView('workspace'); setActiveTab('review');
      setTimeout(() => persistCompany(processed, name), 0);
    } catch (e) { setError(e.message || 'تعذر تحليل الملف.'); }
  }

  async function handleFile(event) {
    const file = event.target.files?.[0]; if (!file) return;
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
      if (!rows.length) throw new Error('الملف لا يحتوي على بيانات.');
      const detected = detectColumns(rows);
      setRawRows(rows); setHeaders(Object.keys(rows[0] || {})); setMapping(detected); setError('');
      setShowMapping(true); // التأكيد إلزامي لمنع اختيار رقم الحساب مكان الاسم
    } catch (e) { setError(e.message || 'حدث خطأ أثناء قراءة الملف.'); }
  }

  function applyMapping() {
    if (!mapping.name) return setError('يجب اختيار عمود اسم الحساب.');
    if (mapping.name === mapping.code) return setError('لا يمكن أن يكون رقم الحساب واسم الحساب من العمود نفسه.');
    if (!mapping.balance && !mapping.debit && !mapping.credit) return setError('اختر الرصيد أو المدين والدائن.');
    loadRows(rawRows, company, mapping);
  }

  function changeClassification(id, section) {
    const option = CLASSIFICATIONS.find(x => x.section === section); if (!option || !result) return;
    const accounts = result.accounts.map(a => a.id === id ? { ...a, ...option, confidence: 1, source: 'manual' } : a);
    const account = accounts.find(a => a.id === id);
    localStorage.setItem(`statmanets-learned-${companyId || 'default'}`, JSON.stringify({ ...learned, [normalize(`${account.code}|${account.name}`)]: section }));
    const next = recalculate(accounts, result.columns); setResult(next); persistCompany(next);
  }

  function newCompany() { setCompanyId(''); setCompany('شركة جديدة'); setResult(null); setRawRows([]); setError(''); setView('workspace'); }
  function openCompany(id) { const c = companies.find(x => x.id === id); if (!c) return; setCompanyId(c.id); setCompany(c.name); setResult(c.result); setView('workspace'); setActiveTab('review'); }
  function deleteCompany(id) { const next = companies.filter(c => c.id !== id); setCompanies(next); saveCompanies(next); }
  function renameCompany(name) { setCompany(name); if (result) persistCompany(result, name); }

  function exportStatements() {
    if (!statements) return;
    const wb = XLSX.utils.book_new(), i = statements.incomeStatement, p = statements.financialPosition;
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['قائمة الدخل', company], ['البند', 'المبلغ'], ['الإيرادات', i.revenue], ['تكلفة الإيرادات', i.cost], ['مجمل الربح', i.grossProfit], ['المصروفات التشغيلية', i.expenses], ['صافي الربح', i.netProfit]]), 'قائمة الدخل');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['قائمة المركز المالي', company], ['الأصول', 'المبلغ'], ...p.assetLines.map(x => [x.label, x.amount]), ['إجمالي الأصول', p.assets], [], ['الالتزامات وحقوق الملكية', 'المبلغ'], ...p.liabilityLines.map(x => [x.label, x.amount]), ...p.equityLines.map(x => [x.label, x.amount]), ['صافي ربح الفترة', i.netProfit]]), 'المركز المالي');
    XLSX.writeFile(wb, `القوائم-المالية-${company}.xlsx`);
  }

  const MappingSelect = ({ field, label, optional }) => <label className="mapping-field"><span>{label}{optional ? ' (اختياري)' : ''}</span><select value={mapping[field] || ''} onChange={e => setMapping({ ...mapping, [field]: e.target.value })}><option value="">— اختر العمود —</option>{headers.map(h => <option key={h} value={h}>{h}</option>)}</select></label>;

  return <main className="app-shell">
    <nav className="topbar"><button className="brand" onClick={() => setView('home')}><span>S</span><b>Statmanets</b></button><div className="nav-actions"><button onClick={() => setView('home')}>الرئيسية</button><button onClick={newCompany}>شركة جديدة</button><button onClick={() => setView('companies')}>الشركات</button><button onClick={() => setView('guide')}>الدليل</button></div></nav>

    {view === 'home' && <header className="hero"><div><span className="eyebrow">SMART FINANCIAL STATEMENTS</span><h1>قوائم مالية من ميزان المراجعة</h1><p>ارفع Excel، اختر الأعمدة بوضوح، راجع أسماء الحسابات والتصنيف، ثم صدّر القوائم.</p><div className="hero-buttons"><button className="primary" onClick={newCompany}>ابدأ الآن</button><button className="secondary" onClick={() => loadRows(SAMPLE_ROWS, 'شركة تجريبية')}>تجربة مباشرة</button></div></div></header>}

    {view === 'companies' && <section className="page"><h1>الشركات المحفوظة</h1><button className="primary" onClick={newCompany}>+ شركة جديدة</button><div className="company-grid">{companies.map(c => <article className="company-card" key={c.id}><div><b>{c.name}</b><small>{new Date(c.updatedAt).toLocaleString('ar-JO')}</small></div><div><button className="secondary" onClick={() => openCompany(c.id)}>فتح</button><button className="danger" onClick={() => deleteCompany(c.id)}>حذف</button></div></article>)}</div></section>}

    {view === 'guide' && <section className="page"><h1>دليل الاستخدام</h1><div className="guide-grid"><article><b>1. ارفع الملف</b><p>اختر ملف ميزان المراجعة.</p></article><article><b>2. أكد الأعمدة</b><p>اختر اسم الحساب ورقمه والمدين والدائن.</p></article><article><b>3. راجع البطاقات</b><p>كل حساب يظهر باسمه ورقمه ورصيده.</p></article><article><b>4. صدّر</b><p>أنشئ القوائم ونزّلها إلى Excel.</p></article></div></section>}

    {view === 'workspace' && <section className="workspace"><aside className="upload-card"><label>اسم الشركة</label><input value={company} onChange={e => renameCompany(e.target.value)} /><label className="drop-zone"><strong>رفع ميزان المراجعة</strong><span>XLSX أو XLS</span><input type="file" accept=".xlsx,.xls" onChange={handleFile} /></label><button className="primary wide" onClick={() => loadRows(SAMPLE_ROWS, 'شركة تجريبية')}>تشغيل المثال</button>{error && <div className="error">{error}</div>}</aside>
      <section className="content-card">{!result ? <div className="empty-state"><h2>ارفع ملف Excel</h2><p>سيطلب التطبيق منك تأكيد عمود اسم الحساب قبل عرض النتائج.</p></div> : <><div className="report-head"><div><small>الشركة</small><h2>{company}</h2></div><span className={result.balanced ? 'status ok' : 'status warn'}>{result.balanced ? 'الميزان متوازن' : 'يوجد فرق'}</span></div>
        <div className="kpis"><article><small>المدين</small><b>{money(result.totals.debit)}</b></article><article><small>الدائن</small><b>{money(result.totals.credit)}</b></article><article><small>الحسابات</small><b>{result.accounts.length}</b></article><article><small>غير مصنف</small><b>{statements.unmapped.length}</b></article></div>
        <div className="tabs"><button className={activeTab === 'review' ? 'active' : ''} onClick={() => setActiveTab('review')}>الحسابات</button><button className={activeTab === 'statements' ? 'active' : ''} onClick={() => setActiveTab('statements')}>القوائم المالية</button></div>
        {activeTab === 'review' ? <><div className="toolbar"><input placeholder="ابحث باسم الحساب أو رقمه" value={query} onChange={e => setQuery(e.target.value)} /><button className="secondary" onClick={() => setShowMapping(true)}>إعادة مطابقة الأعمدة</button></div><div className="accounts-grid">{filtered.map(a => <article className={`account-card ${a.section === 'unmapped' ? 'needs-review' : ''}`} key={a.id}><div className="account-title"><div><span className="field-label">اسم الحساب</span><h3>{a.name || 'اسم الحساب غير متوفر'}</h3></div><span className={a.confidence >= .7 ? 'confidence high' : 'confidence low'}>{Math.round(a.confidence * 100)}%</span></div><div className="account-details"><div><span className="field-label">رقم الحساب</span><b>{a.code || '—'}</b></div><div><span className="field-label">الرصيد</span><b>{money(a.balance)}</b></div></div><label className="classification-field"><span className="field-label">التصنيف</span><select value={a.section} onChange={e => changeClassification(a.id, e.target.value)}>{CLASSIFICATIONS.map(c => <option key={c.section} value={c.section}>{c.label}</option>)}</select></label></article>)}</div></> : <div className="statements"><div className="statement-actions"><h3>القوائم المالية</h3><button className="primary" onClick={exportStatements}>تصدير Excel</button></div><div className="statement-grid"><article className="statement-card"><h3>قائمة الدخل</h3>{[['الإيرادات', statements.incomeStatement.revenue], ['تكلفة الإيرادات', statements.incomeStatement.cost], ['مجمل الربح', statements.incomeStatement.grossProfit], ['المصروفات', statements.incomeStatement.expenses], ['صافي الربح', statements.incomeStatement.netProfit]].map(x => <div className="statement-line" key={x[0]}><span>{x[0]}</span><b>{money(x[1])}</b></div>)}</article><article className="statement-card"><h3>المركز المالي</h3>{[...statements.financialPosition.assetLines, ...statements.financialPosition.liabilityLines, ...statements.financialPosition.equityLines].map(x => <div className="statement-line" key={x.section}><span>{x.label}</span><b>{money(x.amount)}</b></div>)}</article></div></div>}
      </>}</section></section>}

    {showMapping && <div className="modal-overlay"><div className="mapping-card"><h2>تأكيد أعمدة Excel</h2><p>اختر عمود <b>اسم الحساب</b> بنفسك. لن يعرض التطبيق النتائج قبل هذا التأكيد.</p><div className="mapping-grid"><MappingSelect field="name" label="اسم الحساب" /><MappingSelect field="code" label="رقم الحساب" optional /><MappingSelect field="debit" label="مدين" optional /><MappingSelect field="credit" label="دائن" optional /><MappingSelect field="balance" label="الرصيد الصافي" optional /></div><h3>معاينة البيانات</h3><div className="preview-list">{rawRows.slice(0, 5).map((row, i) => <div key={i}><b>{mapping.name ? String(row[mapping.name] || '—') : 'اختر عمود الاسم'}</b><span>{mapping.code ? String(row[mapping.code] || '—') : ''}</span></div>)}</div>{error && <div className="error">{error}</div>}<div className="modal-actions"><button className="secondary" onClick={() => setShowMapping(false)}>إلغاء</button><button className="primary" onClick={applyMapping}>اعتماد ومتابعة</button></div></div></div>}
  </main>;
}
