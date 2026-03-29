/**
 * 人生ゲーム v2 バランス検証シミュレーション
 * 3戦略 x 100回ずつ試行して平均・中央値を確認
 */

const TRIALS = 100;

// ===== Game Constants =====
// 調整案v3: 1期やさしめ、2期から本番、3期は厳しい
const BASE_RATES = {
  1: { churn: 15, turnover: 15, newDeal: 55 },
  2: { churn: 30, turnover: 25, newDeal: 40 },
  3: { churn: 40, turnover: 35, newDeal: 25 },
};

const CAP_TABLE = { 1: 0.5, 2: 1.0, 3: 1.2, 4: 1.5, 5: 2.0 };
const AI_CAP_BONUS = [0, 0.3, 0.5, 0.8];
const AI_MAINTENANCE = [0, 200000, 400000, 700000];
const AI_COST = [0, 5000000, 8000000, 15000000];

const HIRE_CHANNELS = {
  media:    { cost: 500000,  baseRate: 80, lv: 1, salary: 300000 },
  agent:    { cost: 2000000, baseRate: 70, lv: 2, salary: 400000 },
  referral: { cost: 0,       baseRate: 30, lv: 3, salary: 500000 },
};

const OFFICES = [
  { name: '社長宅',     ratio: 0.7, rent: 0,       initial: 0 },
  { name: 'コワーキング', ratio: 0.9, rent: 200000,  initial: 400000 },
  { name: '小規模',     ratio: 1.1, rent: 600000,  initial: 3000000 },
  { name: '中規模',     ratio: 1.2, rent: 1200000, initial: 7000000 },
  { name: '大規模',     ratio: 1.4, rent: 2200000, initial: 15000000 },
];

// ===== State =====
function newGame() {
  return {
    period: 1, month: 1,
    cash: 20000000,
    aiLevel: 0,
    office: 1, // index into OFFICES
    employees: [
      { role: 'MGR', lv: 4, salary: 600000, status: 'active', ramp: 0, trainRemain: 0 },
      { role: 'sales', lv: 2, salary: 400000, status: 'active', ramp: 0, trainRemain: 0 },
      { role: 'sales', lv: 2, salary: 400000, status: 'active', ramp: 0, trainRemain: 0 },
    ],
    deals: [
      { mrr: 1000000, upsellCount: 0, churnMod: 0, acquiredMonth: 0 },
      { mrr: 1000000, upsellCount: 0, churnMod: 0, acquiredMonth: 0 },
    ],
    newBizUsed: false,
    negMonths: 0,
    gameOver: false,
    totalRevenue: 0,
    totalCost: 0,
  };
}

function getMgrCount(g) { return g.employees.filter(e => e.role === 'MGR').length; }
function getActiveReps(g) { return g.employees.filter(e => e.role === 'sales' && (e.status === 'active' || e.status === 'training')).length; }
function getLv4Plus(g) { return g.employees.filter(e => e.lv >= 4 && e.role === 'sales').length; }
function hasCxo(g) { return g.employees.some(e => e.role === 'CXO'); }

function getTotalCap(g) {
  let cap = 0;
  const aiBonus = AI_CAP_BONUS[g.aiLevel] || 0;
  g.employees.forEach(e => {
    if (e.status === 'active') cap += CAP_TABLE[e.lv] + aiBonus;
    else if (e.status === 'training') cap += CAP_TABLE[e.lv] * 0.5 + aiBonus;
  });
  return cap;
}

function getChurnRate(g) {
  const rates = BASE_RATES[g.period];
  let r = rates.churn;
  if (getMgrCount(g) > 0) r -= 10;
  if (g.aiLevel >= 3) r -= 10;
  // MGR penalty
  const reps = getActiveReps(g);
  const mgrs = getMgrCount(g);
  const limit = reps <= 4 ? 4 : 3;
  if (reps > mgrs * limit) r += (reps >= 9 ? 15 : 10);
  if (g.deals.length > getTotalCap(g)) r += 30;
  return Math.max(0, Math.min(r, 100));
}

function getTurnoverRate(g) {
  const rates = BASE_RATES[g.period];
  let r = rates.turnover;
  if (getMgrCount(g) > 0) r -= 10;
  const reps = getActiveReps(g);
  const mgrs = getMgrCount(g);
  const limit = reps <= 4 ? 4 : 3;
  if (reps > mgrs * limit) r += (reps >= 9 ? 15 : 10);
  return Math.max(0, Math.min(r, 100));
}

function getNewDealRate(g) {
  const rates = BASE_RATES[g.period];
  let r = rates.newDeal;
  r += getMgrCount(g) * 10;
  if (hasCxo(g)) r += 30;
  if (g.aiLevel >= 2) r += 15;
  if (g.aiLevel >= 3) r += 10;
  r += Math.min(getLv4Plus(g) * 5, 15);
  return Math.min(r, 90);
}

function getHireRate(g, channel) {
  let r = HIRE_CHANNELS[channel].baseRate;
  r += getMgrCount(g) * 10;
  if (g.aiLevel >= 1) r += 5;
  if (hasCxo(g)) r += 30;
  if (g.aiLevel >= 3) r += 10;
  return Math.min(r, 90);
}

function diceCheck(rate) {
  return Math.random() * 100 < rate;
}

function monthlyPL(g) {
  const office = OFFICES[g.office];
  const totalMrr = g.deals.reduce((s, d) => s + d.mrr, 0);
  const revenue = Math.round(totalMrr * office.ratio);
  const salaries = g.employees.reduce((s, e) => s + e.salary, 0);
  const cost = salaries + office.rent + AI_MAINTENANCE[g.aiLevel];
  return { revenue, cost, profit: revenue - cost };
}

// ===== Actions =====
function doHire(g, channel) {
  const ch = HIRE_CHANNELS[channel];
  const rate = getHireRate(g, channel);
  if (!diceCheck(rate)) return false;

  let ramp = 2;
  if (getMgrCount(g) > 0) ramp = 1;
  if (getMgrCount(g) > 0 && g.aiLevel >= 1) ramp = 0;

  g.employees.push({
    role: 'sales', lv: ch.lv, salary: ch.salary,
    status: ramp > 0 ? 'ramp' : 'active', ramp, trainRemain: 0,
  });
  g.cash -= ch.cost;
  return true;
}

function doNewDeal(g) {
  const rate = getNewDealRate(g);
  const tries = 1 + Math.floor(getActiveReps(g) / 3);
  let gained = 0;
  for (let i = 0; i < tries; i++) {
    if (diceCheck(rate)) {
      g.deals.push({ mrr: 1000000, upsellCount: 0, churnMod: 0, acquiredMonth: g.period * 100 + g.month });
      gained++;
    }
  }
  return gained;
}

function doUpsell(g) {
  const hasLv3 = g.employees.some(e => e.role === 'sales' && e.lv >= 3 && e.status === 'active');
  if (!hasLv3) return false;
  const deal = g.deals.find(d => d.upsellCount < 2);
  if (!deal) return false;
  const mult = (g.period >= 3 && g.aiLevel < 2) ? 1.15 : 1.2;
  deal.mrr = Math.round(deal.mrr * mult);
  deal.upsellCount++;
  deal.churnMod -= 10;
  const cost = g.aiLevel >= 2 ? 300000 : 500000;
  g.cash -= cost;
  return true;
}

function doTraining(g) {
  const mgrs = getMgrCount(g);
  if (mgrs === 0) return false;
  let trained = 0;
  for (const e of g.employees) {
    if (trained >= mgrs) break;
    if (e.role === 'sales' && e.status === 'active' && e.lv < 5) {
      e.status = 'training';
      e.trainRemain = g.aiLevel >= 3 ? 1 : 2;
      g.cash -= 1000000;
      trained++;
    }
  }
  return trained > 0;
}

function doAI(g) {
  if (g.aiLevel >= 3) return false;
  const cost = AI_COST[g.aiLevel + 1];
  g.aiLevel++;
  g.cash -= cost;
  return true;
}

function doOfficeUpgrade(g) {
  if (g.office >= 4) return false;
  const next = g.office + 1;
  g.cash -= OFFICES[next].initial;
  g.office = next;
  return true;
}

// ===== Month End =====
function processMonthEnd(g) {
  // Churn (skip deals acquired this month or last month)
  const churnRate = getChurnRate(g);
  const currentM = g.period * 100 + g.month;
  g.deals = g.deals.filter(d => {
    if (d.acquiredMonth && (currentM - d.acquiredMonth) < 2) return true; // 免除
    const effectiveRate = Math.max(0, Math.min(churnRate + d.churnMod, 100));
    return !diceCheck(effectiveRate);
  });

  // Turnover
  const turnoverRate = getTurnoverRate(g);
  if (diceCheck(turnoverRate)) {
    // Remove latest sales
    for (let i = g.employees.length - 1; i >= 0; i--) {
      if (g.employees[i].role === 'sales') { g.employees.splice(i, 1); break; }
    }
    // Chain 50%
    if (diceCheck(50)) {
      for (let i = g.employees.length - 1; i >= 0; i--) {
        if (g.employees[i].role === 'sales') { g.employees.splice(i, 1); break; }
      }
    }
  }

  // Ace poaching (period 3)
  if (g.period >= 3) {
    g.employees = g.employees.filter(e => {
      if (e.lv >= 4 && e.role === 'sales') return !diceCheck(10); // 0 on d10 = 10%
      return true;
    });
  }

  // Advance ramp/training
  g.employees.forEach(e => {
    if (e.status === 'ramp') {
      e.ramp--;
      if (e.ramp <= 0) e.status = 'active';
    }
    if (e.status === 'training') {
      e.trainRemain--;
      if (e.trainRemain <= 0) {
        e.lv = Math.min(e.lv + 1, 5);
        e.status = 'active';
      }
    }
  });

  // Natural growth (every 12 months)
  // Simplified: skip for sim

  // PL
  const pl = monthlyPL(g);
  g.cash += pl.profit;
  g.totalRevenue += pl.revenue;
  g.totalCost += pl.cost;

  // Cash check
  if (g.cash < 0) {
    g.negMonths++;
    if (g.negMonths >= 2) g.gameOver = true;
  } else {
    g.negMonths = 0;
  }

  // Advance month
  g.month++;
  if (g.month > 12) {
    g.month = 1;
    g.period++;
    if (g.period > 3) g.gameOver = true;
  }
}

// ===== Strategies =====

function strategyHeadcount(g) {
  // Brute force: hire cheap people, get deals
  const cap = getTotalCap(g);
  const remainCap = cap - g.deals.length;

  if (g.employees.length < 3) {
    doHire(g, 'media');
  } else if (remainCap >= 1 && g.deals.length < 10) {
    doNewDeal(g);
  } else if (g.cash > 3000000) {
    doHire(g, 'media');
  } else {
    doNewDeal(g);
  }

  // Office upgrade when affordable
  if (g.period === 2 && g.month === 1 && g.office < 3 && g.cash > 10000000) doOfficeUpgrade(g);
}

function strategyAIPlusTraining(g) {
  const cap = getTotalCap(g);
  const remainCap = cap - g.deals.length;
  const hasLv3 = g.employees.some(e => e.role === 'sales' && e.lv >= 3 && e.status === 'active');

  // Priority: AI first, then train, then deals
  if (g.aiLevel === 0 && g.cash > 8000000 && g.period >= 1 && g.month >= 3) {
    doAI(g);
  } else if (g.aiLevel === 1 && g.cash > 12000000 && g.period >= 2) {
    doAI(g);
  } else if (g.aiLevel === 2 && g.cash > 20000000 && g.period >= 2 && g.month >= 6) {
    doAI(g);
  } else if (getMgrCount(g) > 0 && g.employees.some(e => e.role === 'sales' && e.lv < 4 && e.status === 'active') && g.cash > 5000000) {
    doTraining(g);
  } else if (hasLv3 && g.deals.some(d => d.upsellCount < 2)) {
    doUpsell(g);
  } else if (remainCap >= 1) {
    doNewDeal(g);
  } else if (g.employees.length < 6 && g.cash > 5000000) {
    doHire(g, 'agent');
  } else {
    doNewDeal(g);
  }

  if (g.period === 2 && g.month === 1 && g.office < 3 && g.cash > 10000000) doOfficeUpgrade(g);
}

function strategyBalanced(g) {
  const cap = getTotalCap(g);
  const remainCap = cap - g.deals.length;
  const hasLv3 = g.employees.some(e => e.role === 'sales' && e.lv >= 3 && e.status === 'active');

  // Mix of everything
  const monthInPeriod = (g.period - 1) * 12 + g.month;

  if (g.aiLevel === 0 && g.cash > 8000000 && monthInPeriod >= 4) {
    doAI(g);
  } else if (g.employees.length < 4 && g.cash > 3000000 && monthInPeriod <= 6) {
    doHire(g, 'agent');
  } else if (remainCap >= 1 && g.deals.length < 5) {
    doNewDeal(g);
  } else if (hasLv3 && g.deals.some(d => d.upsellCount < 2)) {
    doUpsell(g);
  } else if (getMgrCount(g) > 0 && g.employees.some(e => e.role === 'sales' && e.lv < 4 && e.status === 'active') && g.cash > 5000000) {
    doTraining(g);
  } else if (g.aiLevel < 3 && g.cash > 15000000) {
    doAI(g);
  } else if (g.employees.length < 6 && g.cash > 5000000) {
    doHire(g, 'agent');
  } else if (remainCap >= 1) {
    doNewDeal(g);
  } else {
    // Do nothing useful, just try new deal
    doNewDeal(g);
  }

  if (g.period >= 2 && g.month === 1 && g.office < 3 && g.cash > 10000000) doOfficeUpgrade(g);
}

// ===== Run Simulation =====
function simulate(strategyFn, name) {
  const results = [];
  for (let t = 0; t < TRIALS; t++) {
    const g = newGame();
    let months = 0;
    while (!g.gameOver && months < 36) {
      strategyFn(g);
      processMonthEnd(g);
      months++;
    }
    results.push({
      cash: g.cash,
      totalRevenue: g.totalRevenue,
      totalCost: g.totalCost,
      totalProfit: g.totalRevenue - g.totalCost,
      deals: g.deals.length,
      employees: g.employees.length,
      aiLevel: g.aiLevel,
      months,
      gameOver: months < 36,
      survived: months >= 36 || g.period > 3,
    });
  }

  // Stats
  const survived = results.filter(r => r.survived).length;
  const cashArr = results.map(r => r.cash).sort((a, b) => a - b);
  const profitArr = results.map(r => r.totalProfit).sort((a, b) => a - b);
  const avgCash = Math.round(cashArr.reduce((a, b) => a + b, 0) / TRIALS);
  const medCash = cashArr[Math.floor(TRIALS / 2)];
  const avgProfit = Math.round(profitArr.reduce((a, b) => a + b, 0) / TRIALS);
  const medProfit = profitArr[Math.floor(TRIALS / 2)];
  const avgDeals = Math.round(results.reduce((s, r) => s + r.deals, 0) / TRIALS * 10) / 10;
  const avgEmps = Math.round(results.reduce((s, r) => s + r.employees, 0) / TRIALS * 10) / 10;
  const avgAI = Math.round(results.reduce((s, r) => s + r.aiLevel, 0) / TRIALS * 10) / 10;
  const avgMonths = Math.round(results.reduce((s, r) => s + r.months, 0) / TRIALS * 10) / 10;
  const gameOvers = results.filter(r => r.gameOver).length;

  console.log(`\n===== ${name} (${TRIALS}回試行) =====`);
  console.log(`生存率: ${survived}/${TRIALS} (${Math.round(survived/TRIALS*100)}%)`);
  console.log(`ゲームオーバー: ${gameOvers}回`);
  console.log(`平均キャッシュ: ${(avgCash/10000).toFixed(0)}万 / 中央値: ${(medCash/10000).toFixed(0)}万`);
  console.log(`平均累計利益: ${(avgProfit/10000).toFixed(0)}万 / 中央値: ${(medProfit/10000).toFixed(0)}万`);
  console.log(`平均案件数: ${avgDeals} / 平均社員数: ${avgEmps} / 平均AI Lv: ${avgAI}`);
  console.log(`平均到達月数: ${avgMonths}/36`);
  console.log(`最終キャッシュ: 最悪${(cashArr[0]/10000).toFixed(0)}万 / 最良${(cashArr[TRIALS-1]/10000).toFixed(0)}万`);
}

console.log('='.repeat(60));
console.log('人生ゲーム v2 バランス検証シミュレーション');
console.log('='.repeat(60));

simulate(strategyHeadcount, '① 人数ゴリ押し戦略');
simulate(strategyAIPlusTraining, '② AI＋育成戦略');
simulate(strategyBalanced, '③ バランス戦略');
