export const CLASSIFICATIONS = [
  ['assets.current.cash', 'النقد وما في حكمه'],
  ['assets.current.receivables', 'الذمم المدينة'],
  ['assets.current.inventory', 'المخزون'],
  ['assets.current.other', 'أصول متداولة أخرى'],
  ['assets.nonCurrent.ppe', 'الممتلكات والمعدات'],
  ['assets.nonCurrent.accumulatedDepreciation', 'مجمع الإهلاك'],
  ['assets.nonCurrent.other', 'أصول غير متداولة أخرى'],
  ['liabilities.current.payables', 'الذمم الدائنة'],
  ['liabilities.current.accruals', 'مصروفات مستحقة'],
  ['liabilities.current.loans', 'قروض قصيرة الأجل'],
  ['liabilities.nonCurrent.loans', 'قروض طويلة الأجل'],
  ['liabilities.other', 'التزامات أخرى'],
  ['equity.capital', 'رأس المال'],
  ['equity.retained', 'الأرباح المحتجزة'],
  ['equity.reserves', 'الاحتياطيات'],
  ['income.revenue', 'الإيرادات'],
  ['income.cost', 'تكلفة الإيرادات'],
  ['income.expense', 'المصروفات التشغيلية'],
  ['income.other', 'إيرادات ومصروفات أخرى'],
  ['unmapped', 'غير مصنف'],
].map(([section, label]) => ({ section, label }));

const RULES = [
  { section: 'assets.nonCurrent.accumulatedDepreciation', label: 'مجمع الإهلاك', terms: ['accumulated depreciation','مجمع اهلاك','مجمع استهلاك'] },
  { section: 'assets.current.cash', label: 'النقد وما في حكمه', terms: ['cash','bank','petty cash','نقد','صندوق','بنك','مصرف','خزينه'] },
  { section: 'assets.current.receivables', label: 'الذمم المدينة', terms: ['receivable','debtor','customer','trade debtors','ذمم مدينه','عملاء','مدينون'] },
  { section: 'assets.current.inventory', label: 'المخزون', terms: ['inventory','stock','merchandise','مخزون','بضاعه','مواد خام','انتاج تام'] },
  { section: 'assets.nonCurrent.ppe', label: 'الممتلكات والمعدات', terms: ['property','plant','equipment','fixed asset','vehicle','furniture','اصول ثابته','معدات','اثاث','سيارات','مباني','اراضي','الات'] },
  { section: 'liabilities.current.payables', label: 'الذمم الدائنة', terms: ['payable','creditor','supplier','trade creditors','ذمم دائنه','موردون','دائنون'] },
  { section: 'liabilities.current.accruals', label: 'مصروفات مستحقة', terms: ['accrued','accrual','مستحق','مصاريف مستحقه'] },
  { section: 'liabilities.nonCurrent.loans', label: 'قروض طويلة الأجل', terms: ['long term loan','long-term loan','قرض طويل','قروض طويله'] },
  { section: 'liabilities.current.loans', label: 'قروض قصيرة الأجل', terms: ['loan','borrowing','finance lease','قرض','قروض','تمويل','تسهيلات'] },
  { section: 'equity.capital', label: 'رأس المال', terms: ['capital','share capital','owner equity','راس المال','حقوق الملكيه'] },
  { section: 'equity.retained', label: 'الأرباح المحتجزة', terms: ['retained earnings','accumulated profit','ارباح محتجزه','ارباح مدوره'] },
  { section: 'equity.reserves', label: 'الاحتياطيات', terms: ['reserve','reserves','احتياطي','احتياطيات'] },
  { section: 'income.cost', label: 'تكلفة الإيرادات', terms: ['cost of sales','cost of goods','cogs','تكلفه المبيعات','تكلفه البضاعه','تكلفه الايراد'] },
  { section: 'income.revenue', label: 'الإيرادات', terms: ['revenue','sales','service income','turnover','ايراد','ايرادات','مبيعات'] },
  { section: 'income.expense', label: 'المصروفات التشغيلية', terms: ['expense','salary','rent','depreciation','utilities','مصروف','رواتب','ايجار','استهلاك','اهلاك','كهرباء','اجور'] },
];

export const normalize = value => String(value ?? '')
  .toLowerCase().replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي')
  .replace(/[^\p{L}\p{N}\s.-]/gu, ' ').replace(/\s+/g, ' ').trim();

const number = value => {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const cleaned = String(value ?? '').replace(/,/g, '').replace(/\s/g, '');
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
};

export function classifyAccount(name, code = '', learned = {}) {
  const key = normalize(`${code}|${name}`);
  if (learned[key]) {
    const found = CLASSIFICATIONS.find(item => item.section === learned[key]);
    if (found) return { ...found, confidence: 1, source: 'learned' };
  }
  const text = normalize(`${name} ${code}`);
  let best = { section: 'unmapped', label: 'غير مصنف', confidence: 0, source: 'automatic' };
  for (const rule of RULES) {
    const hits = rule.terms.filter(term => text.includes(normalize(term))).length;
    const confidence = hits ? Math.min(0.98, 0.62 + hits * 0.12) : 0;
    if (confidence > best.confidence) best = { ...rule, confidence, source: 'automatic' };
  }
  return best;
}

export function detectColumns(rows) {
  const keys = Object.keys(rows[0] || {});
  const pick = terms => keys.find(k => terms.some(t => normalize(k).includes(normalize(t))));
  return {
    code: pick(['account code','code','رقم الحساب','رقم']),
    name: pick(['account name','name','description','اسم الحساب','البيان','الحساب']),
    debit: pick(['debit','مدين']), credit: pick(['credit','دائن']),
    balance: pick(['balance','net','الرصيد','صافي'])
  };
}

export function recalculate(accounts, columns = {}) {
  const totals = accounts.reduce((acc, account) => {
    acc.debit += account.debit; acc.credit += account.credit; acc.balance += account.balance;
    acc.bySection[account.section] = (acc.bySection[account.section] || 0) + account.balance;
    return acc;
  }, { debit: 0, credit: 0, balance: 0, bySection: {} });
  return { columns, accounts, totals, balanced: Math.abs(totals.debit - totals.credit) < 0.01 };
}

export function processTrialBalance(rows, learned = {}) {
  const columns = detectColumns(rows);
  if (!columns.name) throw new Error('تعذر اكتشاف عمود اسم الحساب. تأكد أن الصف الأول يحتوي على أسماء الأعمدة.');
  if (!columns.balance && !columns.debit && !columns.credit) throw new Error('لم يتم العثور على أعمدة المدين والدائن أو الرصيد.');
  const accounts = rows.map((row, index) => {
    const debit = number(columns.debit ? row[columns.debit] : 0);
    const credit = number(columns.credit ? row[columns.credit] : 0);
    const balance = columns.balance ? number(row[columns.balance]) : debit - credit;
    const name = String(row[columns.name] ?? '').trim();
    const code = columns.code ? String(row[columns.code] ?? '').trim() : '';
    return { id: index + 1, code, name, debit, credit, balance, ...classifyAccount(name, code, learned) };
  }).filter(a => a.name && (a.debit || a.credit || a.balance));
  if (!accounts.length) throw new Error('لم يتم العثور على حسابات ذات أرصدة داخل الملف.');
  return recalculate(accounts, columns);
}

const total = (accounts, prefix) => accounts.filter(a => a.section.startsWith(prefix)).reduce((s, a) => s + a.balance, 0);
const lines = (accounts, prefix, creditNature = false) => CLASSIFICATIONS
  .filter(item => item.section.startsWith(prefix) && item.section !== 'unmapped')
  .map(item => ({ ...item, amount: accounts.filter(a => a.section === item.section).reduce((s, a) => s + (creditNature ? -a.balance : a.balance), 0) }))
  .filter(item => Math.abs(item.amount) > 0.0001);

export function buildFinancialStatements(accounts) {
  const revenue = -total(accounts, 'income.revenue');
  const cost = total(accounts, 'income.cost');
  const expenses = total(accounts, 'income.expense');
  const other = total(accounts, 'income.other');
  const grossProfit = revenue - cost;
  const netProfit = grossProfit - expenses - other;
  const assetLines = lines(accounts, 'assets.');
  const liabilityLines = lines(accounts, 'liabilities.', true);
  const equityLines = lines(accounts, 'equity.', true);
  const assets = assetLines.reduce((s, x) => s + x.amount, 0);
  const liabilities = liabilityLines.reduce((s, x) => s + x.amount, 0);
  const recordedEquity = equityLines.reduce((s, x) => s + x.amount, 0);
  return {
    incomeStatement: { revenue, cost, grossProfit, expenses, other, netProfit },
    financialPosition: {
      assetLines, liabilityLines, equityLines,
      assets, liabilities, recordedEquity, equityWithProfit: recordedEquity + netProfit,
      difference: assets - liabilities - recordedEquity - netProfit,
    },
    unmapped: accounts.filter(a => a.section === 'unmapped'),
  };
}
