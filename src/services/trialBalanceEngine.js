const RULES = [
  { section: 'assets.current.cash', label: 'النقد وما في حكمه', terms: ['cash','bank','petty cash','نقد','صندوق','بنك','مصرف'] },
  { section: 'assets.current.receivables', label: 'الذمم المدينة', terms: ['receivable','debtor','customer','trade debtors','ذمم مدينة','عملاء','مدينون'] },
  { section: 'assets.current.inventory', label: 'المخزون', terms: ['inventory','stock','merchandise','مخزون','بضاعة'] },
  { section: 'assets.nonCurrent.ppe', label: 'الممتلكات والمعدات', terms: ['property','plant','equipment','fixed asset','vehicle','furniture','أصول ثابتة','معدات','أثاث','سيارات','مباني','أراضي'] },
  { section: 'liabilities.current.payables', label: 'الذمم الدائنة', terms: ['payable','creditor','supplier','trade creditors','ذمم دائنة','موردون','دائنون'] },
  { section: 'liabilities.current.accruals', label: 'مصروفات مستحقة', terms: ['accrued','accrual','مستحق','مصاريف مستحقة'] },
  { section: 'liabilities.nonCurrent.loans', label: 'قروض طويلة الأجل', terms: ['loan','borrowing','finance lease','قرض','قروض','تمويل'] },
  { section: 'equity.capital', label: 'رأس المال', terms: ['capital','share capital','owner equity','رأس المال','حقوق الملكية'] },
  { section: 'equity.retained', label: 'الأرباح المحتجزة', terms: ['retained earnings','accumulated profit','أرباح محتجزة','أرباح مدورة'] },
  { section: 'income.revenue', label: 'الإيرادات', terms: ['revenue','sales','service income','turnover','إيراد','إيرادات','مبيعات'] },
  { section: 'income.cost', label: 'تكلفة الإيرادات', terms: ['cost of sales','cost of goods','cogs','تكلفة المبيعات','تكلفة البضاعة'] },
  { section: 'income.expense', label: 'المصروفات التشغيلية', terms: ['expense','salary','rent','depreciation','utilities','مصروف','رواتب','إيجار','استهلاك','كهرباء'] }
];

const normalize = value => String(value ?? '')
  .toLowerCase()
  .replace(/[أإآ]/g, 'ا')
  .replace(/ة/g, 'ه')
  .replace(/ى/g, 'ي')
  .replace(/[^\p{L}\p{N}\s]/gu, ' ')
  .replace(/\s+/g, ' ')
  .trim();

export function classifyAccount(name, code = '') {
  const text = normalize(`${name} ${code}`);
  let best = { section: 'unmapped', label: 'غير مصنف', confidence: 0 };
  for (const rule of RULES) {
    const hits = rule.terms.filter(term => text.includes(normalize(term))).length;
    const confidence = hits ? Math.min(0.98, 0.62 + hits * 0.12) : 0;
    if (confidence > best.confidence) best = { ...rule, confidence };
  }
  return best;
}

export function detectColumns(rows) {
  const keys = Object.keys(rows[0] || {});
  const pick = terms => keys.find(k => terms.some(t => normalize(k).includes(normalize(t))));
  return {
    code: pick(['account code','code','رقم الحساب','رقم']),
    name: pick(['account name','name','description','اسم الحساب','البيان','الحساب']),
    debit: pick(['debit','مدين']),
    credit: pick(['credit','دائن']),
    balance: pick(['balance','net','الرصيد','صافي'])
  };
}

export function processTrialBalance(rows) {
  const columns = detectColumns(rows);
  if (!columns.name) throw new Error('تعذر اكتشاف عمود اسم الحساب.');

  const accounts = rows.map((row, index) => {
    const debit = Number(row[columns.debit] || 0);
    const credit = Number(row[columns.credit] || 0);
    const balance = columns.balance ? Number(row[columns.balance] || 0) : debit - credit;
    const classification = classifyAccount(row[columns.name], columns.code ? row[columns.code] : '');
    return {
      id: index + 1,
      code: columns.code ? row[columns.code] : '',
      name: row[columns.name],
      debit,
      credit,
      balance,
      ...classification
    };
  }).filter(a => a.name);

  const totals = accounts.reduce((acc, account) => {
    acc.debit += account.debit;
    acc.credit += account.credit;
    acc.balance += account.balance;
    acc.bySection[account.section] = (acc.bySection[account.section] || 0) + account.balance;
    return acc;
  }, { debit: 0, credit: 0, balance: 0, bySection: {} });

  return { columns, accounts, totals, balanced: Math.abs(totals.debit - totals.credit) < 0.01 };
}
