import { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { processTrialBalance } from './services/trialBalanceEngine';

const money = value => new Intl.NumberFormat('ar', { maximumFractionDigits: 2 }).format(value || 0);

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
  { 'رقم الحساب': '701001', 'اسم الحساب': 'مسحوبات المالك', مدين: 0, دائن: 17000 },
];

const EMPTY_ROWS = [
  { 'رقم الحساب': '', 'اسم الحساب': '', مدين: '', دائن: '' },
  { 'رقم الحساب': '', 'اسم الحساب': '', مدين: '', دائن: '' },
  { 'رقم الحساب': '', 'اسم الحساب': '', مدين: '', دائن: '' },
  { 'رقم الحساب': '', 'اسم الحساب': '', مدين: '', دائن: '' },
];

const GUIDE_SECTIONS = [
  { id: 'start', title: 'البداية السريعة', icon: '01', text: 'اكتب اسم الشركة، ثم ارفع ميزان المراجعة أو جرّب الملف الجاهز. سيقرأ التطبيق الحسابات ويعرض نتيجة الفحص والتصنيف مباشرة.' },
  { id: 'prepare', title: 'تجهيز ملف Excel', icon: '02', text: 'ضع كل حساب في صف مستقل. يجب أن يحتوي الملف على اسم الحساب، ومعه مدين ودائن أو رصيد صافٍ. تجنب الخلايا المدمجة وكلمات المرور.' },
  { id: 'columns', title: 'اكتشاف الأعمدة', icon: '03', text: 'يحاول التطبيق اكتشاف رقم الحساب واسمه والمدين والدائن تلقائيًا، حتى لو كانت أسماء الأعمدة بالعربية أو الإنجليزية.' },
  { id: 'review', title: 'مراجعة التصنيف', icon: '04', text: 'راجع الحسابات ذات الثقة المنخفضة. لاحقًا سيتمكن المستخدم من تعديل التصنيف وحفظه للشركة كي يتعلم التطبيق منه.' },
  { id: 'reports', title: 'إنشاء القوائم', icon: '05', text: 'بعد اعتماد التصنيف، ينشئ النظام قائمة المركز المالي وقائمة الدخل، ثم التدفقات النقدية عند توافر بيانات فترتين.' },
];

function makeWorkbook(rows, fileName, includeInstructions = true) {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows);
  sheet['!cols'] = [{ wch: 16 }, { wch: 30 }, { wch: 16 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(workbook, sheet, 'ميزان المراجعة');

  if (includeInstructions) {
    const instructions = [
      ['دليل تعبئة النموذج'],
      ['1', 'اكتب كل حساب في صف مستقل.'],
      ['2', 'لا تدمج الخلايا داخل جدول الحسابات.'],
      ['3', 'أدخل المبالغ في عمود مدين أو دائن.'],
      ['4', 'تأكد من تساوي إجمالي المدين والدائن.'],
      ['5', 'يمكن استخدام مسميات عربية أو إنجليزية.'],
    ];
    const helpSheet = XLSX.utils.aoa_to_sheet(instructions);
    helpSheet['!cols'] = [{ wch: 8 }, { wch: 70 }];
    XLSX.utils.book_append_sheet(workbook, helpSheet, 'اقرأني أولًا');
  }

  XLSX.writeFile(workbook, fileName);
}

export default function App() {
  const [result, setResult] = useState(null);
  const [error, setError] = useState('');
  const [company, setCompany] = useState('شركة تجريبية');
  const [view, setView] = useState('home');
  const [guideId, setGuideId] = useState('start');
  const [tourStep, setTourStep] = useState(() => localStorage.getItem('statmanets-tour-done') ? -1 : 0);

  const summary = useMemo(() => {
    if (!result) return [];
    return Object.entries(result.totals.bySection).sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]));
  }, [result]);

  function loadRows(rows, name = company) {
    setError('');
    try {
      const processed = processTrialBalance(rows);
      setResult(processed);
      setCompany(name);
      setView('workspace');
    } catch (e) {
      setResult(null);
      setError(e.message || 'تعذر تحليل البيانات.');
    }
  }

  async function handleFile(event) {
    setError('');
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      if (!rows.length) throw new Error('الملف لا يحتوي على بيانات قابلة للقراءة.');
      loadRows(rows);
    } catch (e) {
      setResult(null);
      setError(e.message || 'حدث خطأ أثناء قراءة الملف.');
    }
  }

  function closeTour() {
    localStorage.setItem('statmanets-tour-done', '1');
    setTourStep(-1);
  }

  const activeGuide = GUIDE_SECTIONS.find(item => item.id === guideId);
  const tour = [
    { title: 'مرحبًا بك في Statmanets', text: 'حوّل ميزان المراجعة إلى قوائم مالية منظمة من ملف Excel واحد.' },
    { title: 'لا تحتاج إلى نموذج ثابت', text: 'يتعرف التطبيق على أسماء الأعمدة والحسابات العربية والإنجليزية.' },
    { title: 'راجع قبل الاعتماد', text: 'يعرض التطبيق درجة الثقة ويبرز الحسابات التي تحتاج إلى مراجعة.' },
    { title: 'ابدأ بالتجربة', text: 'استخدم الملف الجاهز الآن أو نزّل النموذج الفارغ واملأه ببيانات شركتك.' },
  ];

  return (
    <main className="app-shell">
      <nav className="topbar">
        <button className="brand" onClick={() => setView('home')}><span>S</span><b>Statmanets</b></button>
        <div className="nav-actions">
          <button className={view === 'home' ? 'active' : ''} onClick={() => setView('home')}>الرئيسية</button>
          <button className={view === 'workspace' ? 'active' : ''} onClick={() => setView('workspace')}>مساحة العمل</button>
          <button className={view === 'guide' ? 'active' : ''} onClick={() => setView('guide')}>دليل الاستخدام</button>
        </div>
      </nav>

      {view === 'home' && (
        <>
          <header className="hero landing-hero">
            <div>
              <span className="eyebrow">FINANCIAL STATEMENTS, SIMPLIFIED</span>
              <h1>أنشئ قوائمك المالية تلقائيًا من ميزان المراجعة</h1>
              <p>ارفع ملف Excel، ودع التطبيق يكتشف الأعمدة ويصنف الحسابات ويكشف الفروقات، ثم راجع النتيجة قبل إعداد القوائم.</p>
              <div className="hero-buttons">
                <button className="primary" onClick={() => setView('workspace')}>ابدأ بملفي</button>
                <button className="secondary" onClick={() => loadRows(SAMPLE_ROWS, 'شركة الأفق التجارية')}>جرّب ملفًا تجريبيًا</button>
                <button className="ghost" onClick={() => setView('guide')}>كيف أستخدم التطبيق؟</button>
              </div>
            </div>
            <div className="hero-panel">
              <div className="mini-file"><span>XLSX</span><b>ميزان المراجعة.xlsx</b><small>جاهز للتحليل</small></div>
              <div className="flow-arrow">↓</div>
              <div className="mini-result"><b>اكتشاف وتصنيف تلقائي</b><span>أصول · التزامات · إيرادات · مصروفات</span></div>
            </div>
          </header>

          <section className="steps-section">
            <div className="section-head"><span>أربع خطوات فقط</span><h2>من ملف Excel إلى نتيجة واضحة</h2></div>
            <div className="steps-grid">
              {GUIDE_SECTIONS.slice(0, 4).map(item => <article key={item.id}><i>{item.icon}</i><h3>{item.title}</h3><p>{item.text}</p></article>)}
            </div>
          </section>

          <section className="sample-section">
            <div><span className="eyebrow dark">ابدأ دون تجهيز ملف</span><h2>جرّب التطبيق ببيانات جاهزة</h2><p>شاهد كيف يقرأ التطبيق ميزان مراجعة متوازنًا، ثم نزّل النموذج واستخدمه كنقطة بداية لشركتك.</p></div>
            <div className="sample-actions">
              <button className="primary" onClick={() => loadRows(SAMPLE_ROWS, 'شركة الأفق التجارية')}>تشغيل الملف التجريبي</button>
              <button className="secondary" onClick={() => makeWorkbook(SAMPLE_ROWS, 'ميزان-مراجعة-تجريبي.xlsx')}>تنزيل الملف التجريبي</button>
              <button className="secondary" onClick={() => makeWorkbook(EMPTY_ROWS, 'نموذج-ميزان-مراجعة-فارغ.xlsx')}>تنزيل نموذج فارغ</button>
            </div>
          </section>
        </>
      )}

      {view === 'guide' && (
        <section className="guide-page">
          <div className="guide-intro"><span className="eyebrow dark">MANUAL</span><h1>دليل استخدام التطبيق</h1><p>تعلم طريقة تجهيز ملفك ومراجعته وإعداد القوائم المالية خطوة بخطوة.</p></div>
          <div className="guide-layout">
            <aside className="guide-menu">
              {GUIDE_SECTIONS.map(item => <button key={item.id} className={guideId === item.id ? 'active' : ''} onClick={() => setGuideId(item.id)}><span>{item.icon}</span>{item.title}</button>)}
            </aside>
            <article className="guide-content">
              <span className="guide-number">{activeGuide.icon}</span>
              <h2>{activeGuide.title}</h2>
              <p>{activeGuide.text}</p>
              {guideId === 'prepare' && <div className="checklist"><b>قبل رفع الملف تأكد من:</b><span>✓ اسم حساب واحد في كل صف</span><span>✓ وجود مدين ودائن أو رصيد</span><span>✓ عدم وجود خلايا مدمجة</span><span>✓ عدم حماية الملف بكلمة مرور</span></div>}
              {guideId === 'start' && <div className="guide-buttons"><button className="primary" onClick={() => loadRows(SAMPLE_ROWS, 'شركة الأفق التجارية')}>ابدأ بالتجربة</button><button className="secondary" onClick={() => makeWorkbook(EMPTY_ROWS, 'نموذج-ميزان-مراجعة-فارغ.xlsx')}>نزّل النموذج</button></div>}
            </article>
          </div>
          <div className="faq"><h2>أسئلة شائعة</h2><details><summary>هل يجب استخدام نموذج محدد؟</summary><p>لا. يحاول التطبيق اكتشاف الأعمدة تلقائيًا، والنموذج المرفق مجرد وسيلة لتسهيل البداية.</p></details><details><summary>ماذا لو لم يتعرف التطبيق على حساب؟</summary><p>يظهر الحساب ضمن العناصر التي تحتاج إلى مراجعة، وسيضاف لاحقًا خيار تعديل التصنيف وحفظه.</p></details><details><summary>هل تعمل البيانات دون إنترنت؟</summary><p>التصميم مهيأ ليعمل محليًا، وتتم قراءة ملف Excel داخل الجهاز.</p></details></div>
        </section>
      )}

      {view === 'workspace' && (
        <section className="workspace-page">
          <div className="workspace-title"><div><span className="eyebrow dark">WORKSPACE</span><h1>استيراد ميزان المراجعة</h1></div><button className="help-button" onClick={() => setView('guide')}>؟ دليل الاستخدام</button></div>
          <section className="workspace">
            <aside className="upload-card">
              <label>اسم الشركة</label>
              <input value={company} onChange={e => setCompany(e.target.value)} />
              <label className="drop-zone"><strong>رفع ميزان المراجعة</strong><span>Excel: XLSX أو XLS</span><input type="file" accept=".xlsx,.xls" onChange={handleFile} /></label>
              <button className="demo-inline" onClick={() => loadRows(SAMPLE_ROWS, 'شركة الأفق التجارية')}>أو استخدم الملف التجريبي</button>
              <div className="feature-list"><span>✓ اكتشاف الأعمدة تلقائيًا</span><span>✓ مسميات عربية وإنجليزية</span><span>✓ درجة ثقة لكل حساب</span><span>✓ كشف عدم توازن الميزان</span></div>
              {error && <div className="error">{error}</div>}
            </aside>

            <section className="content-card">
              {!result ? <div className="empty-state"><div className="empty-icon">↥</div><h2>ابدأ برفع الملف</h2><p>يجب أن يحتوي الملف على اسم الحساب، ومعه مدين ودائن أو عمود رصيد.</p><button className="secondary" onClick={() => makeWorkbook(EMPTY_ROWS, 'نموذج-ميزان-مراجعة-فارغ.xlsx')}>تنزيل نموذج Excel</button></div> : <>
                <div className="report-head"><div><small>ملف الشركة</small><h2>{company}</h2></div><span className={result.balanced ? 'status ok' : 'status warn'}>{result.balanced ? 'الميزان متوازن' : 'يوجد فرق في الميزان'}</span></div>
                <div className="kpis"><article><small>إجمالي المدين</small><b>{money(result.totals.debit)}</b></article><article><small>إجمالي الدائن</small><b>{money(result.totals.credit)}</b></article><article><small>الحسابات</small><b>{result.accounts.length}</b></article><article><small>تحتاج مراجعة</small><b>{result.accounts.filter(a => a.confidence < .7).length}</b></article></div>
                <h3>ملخص التصنيف</h3><div className="summary-grid">{summary.map(([section, value]) => <article key={section}><span>{section}</span><b>{money(value)}</b></article>)}</div>
                <h3>مراجعة الحسابات</h3><div className="table-wrap"><table><thead><tr><th>الحساب</th><th>التصنيف</th><th>الرصيد</th><th>الثقة</th></tr></thead><tbody>{result.accounts.slice(0, 100).map(a => <tr key={a.id}><td><b>{a.name}</b><small>{a.code}</small></td><td>{a.label}</td><td>{money(a.balance)}</td><td><span className={a.confidence >= .7 ? 'confidence high' : 'confidence low'}>{Math.round(a.confidence * 100)}%</span></td></tr>)}</tbody></table></div>
              </>}
            </section>
          </section>
        </section>
      )}

      {tourStep >= 0 && <div className="tour-backdrop"><section className="tour-card"><div className="tour-progress">{tour.map((_, i) => <span key={i} className={i <= tourStep ? 'active' : ''} />)}</div><div className="tour-visual">{tourStep + 1}</div><h2>{tour[tourStep].title}</h2><p>{tour[tourStep].text}</p><div className="tour-actions"><button className="ghost" onClick={closeTour}>تخطي</button>{tourStep < tour.length - 1 ? <button className="primary" onClick={() => setTourStep(tourStep + 1)}>التالي</button> : <button className="primary" onClick={() => { closeTour(); loadRows(SAMPLE_ROWS, 'شركة الأفق التجارية'); }}>جرّب الآن</button>}</div></section></div>}
    </main>
  );
}
