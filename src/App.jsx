import { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { processTrialBalance } from './services/trialBalanceEngine';

const money = value => new Intl.NumberFormat('ar', { maximumFractionDigits: 2 }).format(value || 0);

export default function App() {
  const [result, setResult] = useState(null);
  const [error, setError] = useState('');
  const [company, setCompany] = useState('شركة جديدة');

  const summary = useMemo(() => {
    if (!result) return [];
    return Object.entries(result.totals.bySection).sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]));
  }, [result]);

  async function handleFile(event) {
    setError('');
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      if (!rows.length) throw new Error('الملف لا يحتوي على بيانات قابلة للقراءة.');
      setResult(processTrialBalance(rows));
    } catch (e) {
      setResult(null);
      setError(e.message || 'حدث خطأ أثناء قراءة الملف.');
    }
  }

  return (
    <main className="app-shell">
      <header className="hero">
        <div>
          <span className="eyebrow">STATMANETS</span>
          <h1>القوائم المالية تبدأ من ملف Excel واحد</h1>
          <p>ارفع ميزان المراجعة، راجع التصنيف الذكي، ثم جهّز قوائمك المالية دون التقيد بمسميات حسابات ثابتة.</p>
        </div>
        <div className="platforms">
          <span> iPhone / PWA</span><span>▣ Computer</span><span>Android APK</span>
        </div>
      </header>

      <section className="workspace">
        <aside className="upload-card">
          <label>اسم الشركة</label>
          <input value={company} onChange={e => setCompany(e.target.value)} />
          <label className="drop-zone">
            <strong>رفع ميزان المراجعة</strong>
            <span>Excel: XLSX أو XLS</span>
            <input type="file" accept=".xlsx,.xls" onChange={handleFile} />
          </label>
          <div className="feature-list">
            <span>✓ اكتشاف الأعمدة تلقائيًا</span>
            <span>✓ مسميات عربية وإنجليزية</span>
            <span>✓ درجة ثقة لكل حساب</span>
            <span>✓ كشف عدم توازن الميزان</span>
          </div>
          {error && <div className="error">{error}</div>}
        </aside>

        <section className="content-card">
          {!result ? (
            <div className="empty-state">
              <div className="empty-icon">↥</div>
              <h2>ابدأ برفع الملف</h2>
              <p>يجب أن يحتوي الملف على اسم الحساب، ومعه مدين ودائن أو عمود رصيد.</p>
            </div>
          ) : (
            <>
              <div className="report-head">
                <div><small>ملف الشركة</small><h2>{company}</h2></div>
                <span className={result.balanced ? 'status ok' : 'status warn'}>{result.balanced ? 'الميزان متوازن' : 'يوجد فرق في الميزان'}</span>
              </div>
              <div className="kpis">
                <article><small>إجمالي المدين</small><b>{money(result.totals.debit)}</b></article>
                <article><small>إجمالي الدائن</small><b>{money(result.totals.credit)}</b></article>
                <article><small>الحسابات</small><b>{result.accounts.length}</b></article>
                <article><small>تحتاج مراجعة</small><b>{result.accounts.filter(a => a.confidence < .7).length}</b></article>
              </div>
              <h3>ملخص التصنيف</h3>
              <div className="summary-grid">
                {summary.map(([section, value]) => <article key={section}><span>{section}</span><b>{money(value)}</b></article>)}
              </div>
              <h3>مراجعة الحسابات</h3>
              <div className="table-wrap"><table><thead><tr><th>الحساب</th><th>التصنيف</th><th>الرصيد</th><th>الثقة</th></tr></thead><tbody>
                {result.accounts.slice(0, 100).map(a => <tr key={a.id}><td><b>{a.name}</b><small>{a.code}</small></td><td>{a.label}</td><td>{money(a.balance)}</td><td><span className={a.confidence >= .7 ? 'confidence high' : 'confidence low'}>{Math.round(a.confidence * 100)}%</span></td></tr>)}
              </tbody></table></div>
            </>
          )}
        </section>
      </section>
    </main>
  );
}
