export const CLASSIFICATIONS = [
  ['assets.current.cash', 'النقد وما في حكمه'], ['assets.current.receivables', 'الذمم المدينة'],
  ['assets.current.inventory', 'المخزون'], ['assets.current.other', 'أصول متداولة أخرى'],
  ['assets.nonCurrent.ppe', 'الممتلكات والمعدات'], ['assets.nonCurrent.accumulatedDepreciation', 'مجمع الإهلاك'],
  ['assets.nonCurrent.other', 'أصول غير متداولة أخرى'], ['liabilities.current.payables', 'الذمم الدائنة'],
  ['liabilities.current.accruals', 'مصروفات مستحقة'], ['liabilities.current.loans', 'قروض قصيرة الأجل'],
  ['liabilities.nonCurrent.loans', 'قروض طويلة الأجل'], ['liabilities.other', 'التزامات أخرى'],
  ['equity.capital', 'رأس المال'], ['equity.retained', 'الأرباح المحتجزة'], ['equity.reserves', 'الاحتياطيات'],
  ['income.revenue', 'الإيرادات'], ['income.cost', 'تكلفة الإيرادات'], ['income.expense', 'المصروفات التشغيلية'],
  ['income.other', 'إيرادات ومصروفات أخرى'], ['unmapped', 'غير مصنف'],
].map(([section, label]) => ({ section, label }));

const RULES = [
  { section:'assets.nonCurrent.accumulatedDepreciation',label:'مجمع الإهلاك',terms:['accumulated depreciation','مجمع اهلاك','مجمع استهلاك'] },
  { section:'assets.current.cash',label:'النقد وما في حكمه',terms:['cash','bank','petty cash','نقد','صندوق','بنك','مصرف','خزينه'] },
  { section:'assets.current.receivables',label:'الذمم المدينة',terms:['receivable','debtor','customer','ذمم مدينه','عملاء','مدينون'] },
  { section:'assets.current.inventory',label:'المخزون',terms:['inventory','stock','merchandise','مخزون','بضاعه','مواد خام','انتاج تام'] },
  { section:'assets.nonCurrent.ppe',label:'الممتلكات والمعدات',terms:['property','plant','equipment','fixed asset','vehicle','furniture','اصول ثابته','معدات','اثاث','سيارات','مباني','اراضي','الات'] },
  { section:'liabilities.current.payables',label:'الذمم الدائنة',terms:['payable','creditor','supplier','ذمم دائنه','موردون','دائنون'] },
  { section:'liabilities.current.accruals',label:'مصروفات مستحقة',terms:['accrued','accrual','مستحق','مصاريف مستحقه'] },
  { section:'liabilities.nonCurrent.loans',label:'قروض طويلة الأجل',terms:['long term loan','long-term loan','قرض طويل','قروض طويله'] },
  { section:'liabilities.current.loans',label:'قروض قصيرة الأجل',terms:['loan','borrowing','قرض','قروض','تمويل','تسهيلات'] },
  { section:'equity.capital',label:'رأس المال',terms:['capital','share capital','owner equity','راس المال','حقوق الملكيه'] },
  { section:'equity.retained',label:'الأرباح المحتجزة',terms:['retained earnings','accumulated profit','ارباح محتجزه','ارباح مدوره'] },
  { section:'equity.reserves',label:'الاحتياطيات',terms:['reserve','reserves','احتياطي','احتياطيات'] },
  { section:'income.cost',label:'تكلفة الإيرادات',terms:['cost of sales','cost of goods','cogs','تكلفه المبيعات','تكلفه البضاعه'] },
  { section:'income.revenue',label:'الإيرادات',terms:['revenue','sales','service income','turnover','ايراد','ايرادات','مبيعات'] },
  { section:'income.expense',label:'المصروفات التشغيلية',terms:['expense','salary','rent','depreciation','utilities','مصروف','رواتب','ايجار','استهلاك','اهلاك','كهرباء','اجور'] },
];

export const normalize = value => String(value ?? '').toLowerCase().replace(/[أإآ]/g,'ا').replace(/ة/g,'ه').replace(/ى/g,'ي').replace(/[^\p{L}\p{N}\s.-]/gu,' ').replace(/\s+/g,' ').trim();
const number = value => { if (typeof value === 'number') return Number.isFinite(value) ? value : 0; const n = Number(String(value ?? '').replace(/,/g,'').replace(/\s/g,'')); return Number.isFinite(n) ? n : 0; };
const isMostlyNumeric = values => {
  const useful = values.filter(v => String(v ?? '').trim() !== '').slice(0, 20);
  if (!useful.length) return false;
  return useful.filter(v => /^[-+]?\d[\d,./-]*$/.test(String(v).trim())).length / useful.length >= 0.7;
};

export function classifyAccount(name, code='', learned={}) {
  const key = normalize(`${code}|${name}`);
  if (learned[key]) { const found = CLASSIFICATIONS.find(x=>x.section===learned[key]); if(found) return {...found,confidence:1,source:'learned'}; }
  const text = normalize(`${name} ${code}`); let best={section:'unmapped',label:'غير مصنف',confidence:0,source:'automatic'};
  for(const rule of RULES){ const hits=rule.terms.filter(t=>text.includes(normalize(t))).length; const confidence=hits?Math.min(.98,.62+hits*.12):0; if(confidence>best.confidence) best={...rule,confidence,source:'automatic'}; }
  return best;
}

export function detectColumns(rows){
  const keys = Object.keys(rows[0] || {});
  const normalizedKeys = keys.map(key => ({ key, normalized: normalize(key) }));
  const valuesFor = key => rows.slice(0, 20).map(row => row[key]);
  const score = (key, exactTerms, partialTerms, preferText=false, preferNumeric=false) => {
    const normalized = normalize(key);
    let points = 0;
    exactTerms.forEach(term => { if (normalized === normalize(term)) points += 100; });
    partialTerms.forEach(term => { if (normalized.includes(normalize(term))) points += 20; });
    const numeric = isMostlyNumeric(valuesFor(key));
    if (preferText) points += numeric ? -45 : 25;
    if (preferNumeric) points += numeric ? 25 : -20;
    return points;
  };
  const best = (exact, partial, preferText=false, preferNumeric=false, excluded=[]) => normalizedKeys
    .filter(item => !excluded.includes(item.key))
    .map(item => ({ key:item.key, score:score(item.key,exact,partial,preferText,preferNumeric) }))
    .filter(item => item.score > 0)
    .sort((a,b) => b.score-a.score)[0]?.key;

  const code = best(
    ['رقم الحساب','كود الحساب','account code','account number','gl code','code'],
    ['رقم','كود','code','number','no'], false, true
  );
  const name = best(
    ['اسم الحساب','اسم البند','account name','account description','description','البيان','الحساب'],
    ['اسم الحساب','اسم','description','account name','البيان'], true, false, code ? [code] : []
  );
  const debit = best(['مدين','debit','debit balance'],['مدين','debit'],false,true);
  const credit = best(['دائن','credit','credit balance'],['دائن','credit'],false,true,debit?[debit]:[]);
  const balance = best(['الرصيد','صافي الرصيد','balance','net balance'],['الرصيد','balance','net'],false,true,[debit,credit].filter(Boolean));
  return { code, name, debit, credit, balance };
}

export function recalculate(accounts,columns={}){
  const totals=accounts.reduce((a,x)=>{a.debit+=x.debit;a.credit+=x.credit;a.balance+=x.balance;a.bySection[x.section]=(a.bySection[x.section]||0)+x.balance;return a;},{debit:0,credit:0,balance:0,bySection:{}});
  return {columns,accounts,totals,balanced:Math.abs(totals.debit-totals.credit)<.01};
}

export function processTrialBalance(rows, learned={}, overrideColumns=null){
  const columns=overrideColumns||detectColumns(rows);
  if(!columns.name) throw new Error('تعذر اكتشاف عمود اسم الحساب. استخدم شاشة مطابقة الأعمدة.');
  if(columns.name === columns.code) throw new Error('تم اختيار العمود نفسه لرقم الحساب واسم الحساب. افتح مطابقة الأعمدة وحدد عمود الاسم الصحيح.');
  if(!columns.balance&&!columns.debit&&!columns.credit) throw new Error('حدد عمود الرصيد أو أعمدة المدين والدائن.');
  const accounts=rows.map((row,index)=>{ const debit=number(columns.debit?row[columns.debit]:0); const credit=number(columns.credit?row[columns.credit]:0); const balance=columns.balance?number(row[columns.balance]):debit-credit; const name=String(row[columns.name]??'').trim(); const code=columns.code?String(row[columns.code]??'').trim():''; return {id:index+1,code,name,debit,credit,balance,...classifyAccount(name,code,learned)}; }).filter(a=>a.name&&(a.debit||a.credit||a.balance));
  if(!accounts.length) throw new Error('لم يتم العثور على حسابات ذات أرصدة داخل الملف.');
  return recalculate(accounts,columns);
}

const total=(accounts,prefix)=>accounts.filter(a=>a.section.startsWith(prefix)).reduce((s,a)=>s+a.balance,0);
const lines=(accounts,prefix,creditNature=false)=>CLASSIFICATIONS.filter(i=>i.section.startsWith(prefix)&&i.section!=='unmapped').map(i=>({...i,amount:accounts.filter(a=>a.section===i.section).reduce((s,a)=>s+(creditNature?-a.balance:a.balance),0)})).filter(i=>Math.abs(i.amount)>.0001);
export function buildFinancialStatements(accounts){
  const revenue=-total(accounts,'income.revenue'),cost=total(accounts,'income.cost'),expenses=total(accounts,'income.expense'),other=total(accounts,'income.other');
  const grossProfit=revenue-cost,netProfit=grossProfit-expenses-other; const assetLines=lines(accounts,'assets.'),liabilityLines=lines(accounts,'liabilities.',true),equityLines=lines(accounts,'equity.',true);
  const assets=assetLines.reduce((s,x)=>s+x.amount,0),liabilities=liabilityLines.reduce((s,x)=>s+x.amount,0),recordedEquity=equityLines.reduce((s,x)=>s+x.amount,0);
  return {incomeStatement:{revenue,cost,grossProfit,expenses,other,netProfit},financialPosition:{assetLines,liabilityLines,equityLines,assets,liabilities,recordedEquity,equityWithProfit:recordedEquity+netProfit,difference:assets-liabilities-recordedEquity-netProfit},unmapped:accounts.filter(a=>a.section==='unmapped')};
}
