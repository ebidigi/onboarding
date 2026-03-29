/**
 * 人生ゲーム v2 - GAS バックエンド
 * スプレッドシートID: 1VK6y-FdargJ037BkWz3n37PpVixDOPvowrjxlmPqtOU
 */

const SS_ID = '1Z0VyC3XdOY-phKkx5Oo__Y0ZoLP7cmAhRIKel5ZFkGI';
const DB_SHEET = 'v2_データベース';

// ===== HtmlService dispatcher (called from google.script.run) =====

function gasCallDispatch(action, args) {
  switch (action) {
    case 'createTeam':
      return JSON.stringify(createTeam(args[0]));
    case 'getTeamState':
      return JSON.stringify(getTeamState(args[0]));
    case 'listTeams':
      return JSON.stringify(listTeams());
    case 'executeUpsell':
      saveSnapshot(args[0]);
      return JSON.stringify(executeEvent(args[0], { eventType: 'upsell', dealIndex: args[1] }));
    case 'executeNewDeal':
      saveSnapshot(args[0]);
      return JSON.stringify(executeEvent(args[0], { eventType: 'newdeal', diceValue: args[1] }));
    case 'executeTraining':
      saveSnapshot(args[0]);
      return JSON.stringify(executeEvent(args[0], { eventType: 'training', employeeIndices: args[1] }));
    case 'executeAI':
      saveSnapshot(args[0]);
      return JSON.stringify(executeEvent(args[0], { eventType: 'ai' }));
    case 'executeNewBiz':
      saveSnapshot(args[0]);
      return JSON.stringify(executeEvent(args[0], { eventType: 'newbiz', bizType: args[1] }));
    case 'monthEndChurnBatch':
      saveSnapshot(args[0]);
      return JSON.stringify(monthEndChurn(args[0], { diceValues: args[1] }));
    case 'monthEndTurnoverSingle':
      return JSON.stringify(monthEndTurnover(args[0], { diceValue: args[1], chainDiceValue: args[2] }));
    case 'advanceMonth':
      return JSON.stringify(advanceMonth(args[0]));
    case 'moveOffice':
      saveSnapshot(args[0]);
      return JSON.stringify(moveOffice(args[0], args[1]));
    case 'raisePayOrPromote':
      saveSnapshot(args[0]);
      return JSON.stringify(raisePayOrPromote(args[0], { employeeIndex: args[1], newSalary: args[2], promote: args[3] }));
    case 'undoLastAction':
      return JSON.stringify(restoreSnapshot(args[0]));
    case 'bailout':
      return JSON.stringify(executeBailout(args[0]));
    case 'getEmployeeList':
      return JSON.stringify(getEmployeeList());
    case 'setInitialNames':
      return JSON.stringify(setInitialNames(args[0], args[1]));
    case 'executeHire':
      saveSnapshot(args[0]);
      return JSON.stringify(executeEvent(args[0], { eventType: 'hire', channel: args[1], diceValue: args[2], employeeName: args[3] }));
    case 'savePlan':
      return JSON.stringify(savePlan(args[0], args[1], args[2]));
    case 'getPlan':
      return JSON.stringify(getPlan(args[0], args[1]));
    case 'getPlHistory':
      return JSON.stringify(getPlHistory(args[0], args[1]));
    default:
      return JSON.stringify({ error: 'Unknown action: ' + action });
  }
}

// ===== Snapshot (Undo) =====

function saveSnapshot(teamName) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  if (!sheet) return;

  // Save entire sheet data to PropertiesService
  var data = sheet.getDataRange().getValues();
  var props = PropertiesService.getScriptProperties();
  props.setProperty('snapshot_' + teamName, JSON.stringify(data));
}

function restoreSnapshot(teamName) {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('snapshot_' + teamName);
  if (!raw) return { error: '取り消し可能なデータがありません' };

  var data = JSON.parse(raw);
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  if (!sheet) return { error: 'チームシートが見つかりません' };

  // Clear and restore
  sheet.getDataRange().clearContent();
  if (data.length > 0 && data[0].length > 0) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }

  // Remove used snapshot
  props.deleteProperty('snapshot_' + teamName);

  return { success: true, message: '直前のアクションを取り消しました' };
}

// ===== Web App Endpoints =====

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'ui';
  const team = (e && e.parameter && e.parameter.team) || '';

  if (action === 'getState' && team) {
    const state = getTeamState(team);
    return jsonResponse(state);
  }

  if (action === 'listTeams') {
    const teams = listTeams();
    return jsonResponse({ teams });
  }

  // Default: serve the dice UI
  const html = HtmlService.createHtmlOutputFromFile('dice')
    .setTitle('DigiMan 経営シミュレーション')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const team = data.team;

    switch (action) {
      case 'createTeam':
        return jsonResponse(createTeam(team));
      case 'executeEvent':
        return jsonResponse(executeEvent(team, data));
      case 'monthEndChurn':
        return jsonResponse(monthEndChurn(team, data));
      case 'monthEndTurnover':
        return jsonResponse(monthEndTurnover(team, data));
      case 'monthEndPoach':
        return jsonResponse(monthEndPoach(team, data));
      case 'advanceMonth':
        return jsonResponse(advanceMonth(team));
      case 'moveOffice':
        return jsonResponse(moveOffice(team, data.office));
      case 'raisePayOrPromote':
        return jsonResponse(raisePayOrPromote(team, data));
      default:
        return jsonResponse({ error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.message, stack: err.stack });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== Team Management =====

function listTeams() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheets = ss.getSheets();
  const teams = [];
  sheets.forEach(function(s) {
    const name = s.getName();
    if (name.indexOf('team_') === 0) {
      teams.push(name.replace('team_', ''));
    }
  });
  return teams;
}

function createTeam(teamName) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheetName = 'team_' + teamName;

  // Check if already exists
  if (ss.getSheetByName(sheetName)) {
    return { error: 'チーム「' + teamName + '」は既に存在します' };
  }

  // Copy template
  const template = ss.getSheetByName('v2_テンプレート');
  if (!template) return { error: 'テンプレートシートが見つかりません' };

  const newSheet = template.copyTo(ss);
  newSheet.setName(sheetName);
  newSheet.getRange('B1').setValue(teamName);

  return { success: true, team: teamName };
}

function getEmployeeList() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('社員マスター');
  if (!sheet) return { error: '社員マスターが見つかりません' };
  var data = sheet.getRange('A2:C' + sheet.getLastRow()).getValues();
  var employees = [];
  data.forEach(function(row) {
    if (row[0] && row[0] !== '') {
      employees.push({ name: row[0], type: row[1], hireDate: row[2] });
    }
  });
  return employees;
}

function setInitialNames(teamName, names) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  if (!sheet) return { error: 'チームが見つかりません' };
  // names = [MGR名, 営業1名, 営業2名]
  if (names[0]) sheet.getRange(12, 1).setValue(names[0]);
  if (names[1]) sheet.getRange(13, 1).setValue(names[1]);
  if (names[2]) sheet.getRange(14, 1).setValue(names[2]);
  return { success: true };
}

// ===== State Reading =====

function getTeamState(teamName) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('team_' + teamName);
  if (!sheet) return { error: 'チームが見つかりません' };

  // Basic state (A1:B9) + bailout flag (D1)
  const basic = sheet.getRange('A1:D9').getValues();
  const state = {
    team: basic[0][1],
    period: basic[1][1],
    month: basic[2][1],
    aiLevel: basic[3][1],
    office: basic[4][1],
    cash: basic[5][1],
    newBizUsed: basic[6][1] === 'Y',
    consecutiveNegativeMonths: basic[7][1],
    revenueMultiplier: basic[8][1],
    bailoutUsed: basic[0][3] === 'Y',
  };

  // Employees (A12:I25)
  const empData = sheet.getRange('A12:I25').getValues();
  state.employees = [];
  empData.forEach(function(row) {
    if (row[0] && row[0] !== '') {
      state.employees.push({
        id: row[0],
        role: row[1],
        skillLv: row[2],
        salary: row[3],
        status: row[4],
        rampRemaining: row[5],
        trainingRemaining: row[6],
        hireMonth: row[7],
        capContribution: row[8],
      });
    }
  });

  // Deals (A27:E39) — E列=取得月(期*100+月)
  const dealData = sheet.getRange('A27:E39').getValues();
  state.deals = [];
  dealData.forEach(function(row) {
    if (row[0] && row[0] !== '') {
      state.deals.push({
        id: row[0],
        mrr: row[1],
        upsellCount: row[2],
        churnModifier: row[3],
        acquiredMonth: row[4] || 0,
      });
    }
  });

  // Calculate derived values from data (no spreadsheet formulas needed)
  var aiCapBonus = [0, 0.3, 0.5, 0.8][state.aiLevel] || 0;
  var totalCap = 0;
  var mgrCount = 0;
  var hasCxo = false;
  var activeRepCount = 0;
  var lv4PlusCount = 0;

  state.employees.forEach(function(emp) {
    if (emp.role === 'MGR') mgrCount++;
    if (emp.role === 'CXO') hasCxo = true;
    if (emp.role === '営業' || emp.role === 'MGR') {
      if (emp.status === '稼働') {
        totalCap += (emp.capContribution || getCapContribution(emp.skillLv)) + aiCapBonus;
        if (emp.role === '営業') activeRepCount++;
      } else if (emp.status === '研修') {
        totalCap += ((emp.capContribution || getCapContribution(emp.skillLv)) * 0.5) + aiCapBonus;
        if (emp.role === '営業') activeRepCount++;
      }
      // ランプ中: cap = 0, not counted as active
    }
    if (emp.skillLv >= 4 && emp.role !== 'MGR' && emp.role !== 'CXO') lv4PlusCount++;
  });

  state.totalCapacity = Math.round(totalCap * 10) / 10;
  state.dealCount = state.deals.length;
  state.remainingCapacity = Math.round((totalCap - state.deals.length) * 10) / 10;
  state.mgrCount = mgrCount;
  state.hasCxo = hasCxo;
  state.activeRepCount = activeRepCount;
  state.lv4PlusCount = lv4PlusCount;

  // Calculate rates for current period
  state.rates = getBaseRates(state.period);
  state.calculatedRates = calculateRates(state);

  return state;
}

function getBaseRates(period) {
  var rates = { churn: 15, turnover: 15, newDeal: 55 };
  if (period === 2) { rates = { churn: 30, turnover: 25, newDeal: 40 }; }
  if (period >= 3) { rates = { churn: 40, turnover: 35, newDeal: 25 }; }
  return rates;
}

function calculateRates(state) {
  var rates = state.rates;
  var aiLv = state.aiLevel;
  var mgrCount = state.mgrCount;
  var hasCxo = state.hasCxo;
  var lv4Count = state.lv4PlusCount;
  var activeReps = state.activeRepCount;

  // MGR penalty check
  var mgrPenalty = getMgrPenalty(activeReps, mgrCount);

  // Churn rate
  var churnRate = rates.churn;
  if (mgrCount > 0) churnRate -= 10;
  if (aiLv >= 3) churnRate -= 10;
  churnRate += mgrPenalty.churn;
  if (state.dealCount > state.totalCapacity) churnRate += 30;
  churnRate = Math.max(0, Math.min(churnRate, 100));

  // Turnover rate
  var turnoverRate = rates.turnover;
  if (mgrCount > 0) turnoverRate -= 10;
  turnoverRate += mgrPenalty.turnover;
  turnoverRate = Math.max(0, Math.min(turnoverRate, 100));

  // New deal rate
  var newDealRate = rates.newDeal;
  newDealRate += mgrCount * 10;
  if (hasCxo) newDealRate += 30;
  if (aiLv >= 2) newDealRate += 15;
  if (aiLv >= 3) newDealRate += 10;
  newDealRate += Math.min(lv4Count * 5, 15);
  newDealRate = Math.min(newDealRate, 90);

  // Hire rates per channel
  var hireBase = { media: 80, agent: 70, referral: 30, cxo: 20 };
  var hireRates = {};
  Object.keys(hireBase).forEach(function(ch) {
    var r = hireBase[ch];
    r += mgrCount * 10;
    if (aiLv >= 1) r += 5;
    if (hasCxo) r += 30;
    if (aiLv >= 3) r += 10;
    hireRates[ch] = Math.min(r, 90);
  });

  // Try count for new deals
  var tryCount = 1 + Math.floor(activeReps / 3);

  // Upsell cost
  var upsellCost = (aiLv >= 2) ? 300000 : 500000;

  return {
    churnRate: churnRate,
    turnoverRate: turnoverRate,
    newDealRate: newDealRate,
    hireRates: hireRates,
    tryCount: tryCount,
    upsellCost: upsellCost,
    mgrPenalty: mgrPenalty,
  };
}

function getMgrPenalty(activeReps, mgrCount) {
  var limit, penChurn, penTurnover, penNewDeal;

  if (activeReps <= 4) {
    limit = 4;
    penChurn = 10; penTurnover = 10; penNewDeal = 0;
  } else if (activeReps <= 8) {
    limit = 3;
    penChurn = 10; penTurnover = 10; penNewDeal = 0;
  } else {
    limit = 3;
    penChurn = 15; penTurnover = 15; penNewDeal = -10;
  }

  var manageable = mgrCount * limit;
  if (activeReps > manageable) {
    return { churn: penChurn, turnover: penTurnover, newDeal: penNewDeal, active: true };
  }
  return { churn: 0, turnover: 0, newDeal: 0, active: false };
}

// ===== Event Execution =====

function executeEvent(teamName, data) {
  var eventType = data.eventType;

  switch (eventType) {
    case 'upsell':
      return executeUpsell(teamName, data);
    case 'newdeal':
      return executeNewDeal(teamName, data);
    case 'hire':
      return executeHire(teamName, data);
    case 'training':
      return executeTraining(teamName, data);
    case 'ai':
      return executeAI(teamName, data);
    case 'newbiz':
      return executeNewBiz(teamName, data);
    default:
      return { error: 'Unknown event type' };
  }
}

function executeUpsell(teamName, data) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  var state = getTeamState(teamName);

  // Check Lv3+ exists
  var hasLv3 = state.employees.some(function(e) {
    return e.role === '営業' && e.skillLv >= 3 && e.status === '稼働';
  });
  if (!hasLv3) return { error: 'Lv3以上の社員が必要です' };

  var dealIdx = data.dealIndex; // 0-based index in deals array
  if (dealIdx < 0 || dealIdx >= state.deals.length) return { error: '案件が見つかりません' };

  var deal = state.deals[dealIdx];
  if (deal.upsellCount >= 2) return { error: 'この案件はアップセル上限（2回）に達しています' };

  // Calculate new MRR
  var multiplier = 1.2;
  if (state.period >= 3 && state.aiLevel < 2) multiplier = 1.15;
  var newMrr = Math.round(deal.mrr * multiplier);

  // Update deal
  var dealRow = 27 + dealIdx;
  sheet.getRange(dealRow, 2).setValue(newMrr); // MRR
  sheet.getRange(dealRow, 3).setValue(deal.upsellCount + 1); // upsell count
  sheet.getRange(dealRow, 4).setValue(deal.churnModifier - 10); // churn -10pt

  // Deduct cost
  var cost = (state.aiLevel >= 2) ? 300000 : 500000;
  var newCash = state.cash - cost;
  sheet.getRange('B6').setValue(newCash);

  return {
    success: true,
    event: 'upsell',
    dealId: deal.id,
    oldMrr: deal.mrr,
    newMrr: newMrr,
    cost: cost,
    multiplier: multiplier,
  };
}

function executeNewDeal(teamName, data) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  var state = getTeamState(teamName);

  var diceValue = data.diceValue; // 0-9
  var rate = state.calculatedRates.newDealRate;
  var threshold = Math.ceil(rate / 10); // success if diceValue < threshold
  var success = diceValue < threshold;

  if (success) {
    // Add new deal
    var nextDealRow = 27 + state.deals.length;
    if (nextDealRow > 39) return { error: '案件枠が上限です' };

    var dealId = 'D' + (state.deals.length + 1);
    var mrr = 1000000;
    // Check year event: 景気後退
    sheet.getRange(nextDealRow, 1).setValue(dealId);
    sheet.getRange(nextDealRow, 2).setValue(mrr);
    sheet.getRange(nextDealRow, 3).setValue(0);
    sheet.getRange(nextDealRow, 4).setValue(0);
    sheet.getRange(nextDealRow, 5).setValue(state.period * 100 + state.month); // 取得月
  }

  return {
    success: success,
    event: 'newdeal',
    diceValue: diceValue,
    rate: rate,
    threshold: threshold,
  };
}

function executeHire(teamName, data) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  var state = getTeamState(teamName);

  var channel = data.channel; // 'media', 'agent', 'referral', 'cxo'
  var diceValue = data.diceValue;

  // CXO requires MGR1+ and deals 3+
  if (channel === 'cxo') {
    if (state.mgrCount < 1) return { success: false, error: 'CXO採用にはMGRが1人以上必要です' };
    if (state.dealCount < 3) return { success: false, error: 'CXO採用には案件が3件以上必要です' };
  }

  var rate = state.calculatedRates.hireRates[channel];
  var threshold = Math.ceil(rate / 10);
  var success = diceValue < threshold;

  if (success) {
    var channelInfo = {
      media:    { lv: 1, salary: 300000, cost: 500000, role: '営業' },
      agent:    { lv: 2, salary: 400000, cost: 2000000, role: '営業' },
      referral: { lv: 3, salary: 500000, cost: 0, role: '営業' },
      cxo:      { lv: 0, salary: 2000000, cost: 5000000, role: 'CXO' },
    };
    var info = channelInfo[channel];

    // Determine ramp
    var ramp = 2; // default: no MGR
    if (state.mgrCount > 0) ramp = 1;
    if (state.mgrCount > 0 && state.aiLevel >= 1) ramp = 0;
    if (channel === 'cxo') ramp = 0; // CXO has no ramp

    // Find next employee slot
    var nextEmpRow = 12 + state.employees.length;
    if (nextEmpRow > 25) return { success: true, error: '社員枠が上限です' };

    var empId = data.employeeName || ('E' + (state.employees.length + 1));
    var status = ramp > 0 ? 'ランプ' : '稼働';
    var capContrib = ramp > 0 ? 0 : getCapContribution(info.lv);

    sheet.getRange(nextEmpRow, 1).setValue(empId);
    sheet.getRange(nextEmpRow, 2).setValue(info.role);
    sheet.getRange(nextEmpRow, 3).setValue(info.lv);
    sheet.getRange(nextEmpRow, 4).setValue(info.salary);
    sheet.getRange(nextEmpRow, 5).setValue(status);
    sheet.getRange(nextEmpRow, 6).setValue(ramp);
    sheet.getRange(nextEmpRow, 7).setValue(0);
    sheet.getRange(nextEmpRow, 8).setValue(state.period + '期' + state.month + '月');
    sheet.getRange(nextEmpRow, 9).setValue(capContrib);

    // Deduct hire cost
    var newCash = state.cash - info.cost;
    sheet.getRange('B6').setValue(newCash);

    return {
      success: true,
      event: 'hire',
      channel: channel,
      diceValue: diceValue,
      rate: rate,
      threshold: threshold,
      hired: { id: empId, role: info.role, lv: info.lv, salary: info.salary, ramp: ramp },
      cost: info.cost,
    };
  }

  return {
    success: false,
    event: 'hire',
    channel: channel,
    diceValue: diceValue,
    rate: rate,
    threshold: threshold,
  };
}

function executeTraining(teamName, data) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  var state = getTeamState(teamName);

  var empIndices = data.employeeIndices; // array of 0-based indices
  if (!empIndices || empIndices.length === 0) return { error: '研修対象を選択してください' };

  // Check MGR limit
  if (empIndices.length > state.mgrCount) {
    return { error: 'MGR数(' + state.mgrCount + ')を超える人数は同時に研修できません' };
  }

  var trained = [];
  var totalCost = 0;

  empIndices.forEach(function(idx) {
    var emp = state.employees[idx];
    if (!emp) return;
    if (emp.status === 'ランプ') return;
    if (emp.skillLv >= 5) return;
    if (emp.role === 'CXO') return;

    var empRow = 12 + idx;
    var trainingMonths = (state.aiLevel >= 3) ? 1 : 2;

    sheet.getRange(empRow, 5).setValue('研修'); // status
    sheet.getRange(empRow, 7).setValue(trainingMonths); // training remaining
    // Cap contribution halved during training
    var baseCap = getCapContribution(emp.skillLv);
    sheet.getRange(empRow, 9).setValue(baseCap * 0.5);

    totalCost += 1000000;
    trained.push({ id: emp.id, lv: emp.skillLv, trainingMonths: trainingMonths });
  });

  // Deduct cost
  var newCash = state.cash - totalCost;
  sheet.getRange('B6').setValue(newCash);

  return {
    success: true,
    event: 'training',
    trained: trained,
    totalCost: totalCost,
  };
}

function executeAI(teamName, data) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  var state = getTeamState(teamName);

  var currentAI = state.aiLevel;
  if (currentAI >= 3) return { error: 'AI Lv3が最大です' };

  var costs = [0, 5000000, 8000000, 15000000];
  var cost = costs[currentAI + 1];

  var newAI = currentAI + 1;
  sheet.getRange('B4').setValue(newAI);

  // Deduct cost (maintenance starts next month, handled in advanceMonth)
  var newCash = state.cash - cost;
  sheet.getRange('B6').setValue(newCash);

  return {
    success: true,
    event: 'ai',
    oldLevel: currentAI,
    newLevel: newAI,
    cost: cost,
  };
}

function executeNewBiz(teamName, data) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  var state = getTeamState(teamName);

  if (state.newBizUsed) return { error: '新規事業は全期間で1回のみです' };
  if (state.remainingCapacity < 3) return { error: 'キャパ空き3以上が必要です' };

  var bizType = data.bizType; // 'saas', 'school', 'resident'
  var bizInfo = {
    saas: { cost: 10000000, mrr: 1000000 },
    school: { cost: 5000000, mrr: 600000 },
    resident: { cost: 2000000, mrr: 800000 },
  };
  var info = bizInfo[bizType];
  if (!info) return { error: '不明な事業タイプ' };

  sheet.getRange('B7').setValue('Y');
  var newCash = state.cash - info.cost;
  sheet.getRange('B6').setValue(newCash);

  return {
    success: true,
    event: 'newbiz',
    bizType: bizType,
    cost: info.cost,
    mrr: info.mrr,
  };
}

// ===== Month-End Processing =====

function monthEndChurn(teamName, data) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  var state = getTeamState(teamName);

  var diceValues = data.diceValues; // array of dice values, one per deal
  var churnRate = state.calculatedRates.churnRate;
  var results = [];

  var currentMonth = state.period * 100 + state.month;

  for (var i = state.deals.length - 1; i >= 0; i--) {
    var deal = state.deals[i];

    // 取得当月の案件は解約チェックをスキップ（2ヶ月目以降のみ対象）
    var monthsSinceAcquired = currentMonth - (deal.acquiredMonth || 0);
    if (deal.acquiredMonth > 0 && monthsSinceAcquired < 2) {
      results.push({
        dealId: deal.id,
        diceValue: -1,
        rate: 0,
        threshold: 0,
        churned: false,
        skipped: true,
      });
      continue;
    }

    var effectiveRate = churnRate + deal.churnModifier;
    effectiveRate = Math.max(0, Math.min(effectiveRate, 100));
    var threshold = Math.ceil(effectiveRate / 10);
    var churned = diceValues[i] < threshold;

    results.push({
      dealId: deal.id,
      diceValue: diceValues[i],
      rate: effectiveRate,
      threshold: threshold,
      churned: churned,
    });

    if (churned) {
      var dealRow = 27 + i;
      sheet.getRange(dealRow, 1, 1, 5).clearContent();
    }
  }

  // Compact deal rows (move non-empty up)
  compactDeals(sheet);

  return { success: true, results: results };
}

function monthEndTurnover(teamName, data) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  var state = getTeamState(teamName);

  var diceValue = data.diceValue;
  var turnoverRate = state.calculatedRates.turnoverRate;
  var threshold = Math.ceil(turnoverRate / 10);
  var turnedOver = diceValue < threshold;

  var result = {
    diceValue: diceValue,
    rate: turnoverRate,
    threshold: threshold,
    turnedOver: turnedOver,
    removedEmployees: [],
  };

  if (turnedOver) {
    // Remove most recent non-MGR employee
    var removed = removeLatestEmployee(sheet, state);
    if (removed) result.removedEmployees.push(removed);

    // 50% chance of chain turnover
    if (data.chainDiceValue !== undefined) {
      var chainTurnover = data.chainDiceValue < 5; // 50%
      result.chainCheck = { diceValue: data.chainDiceValue, triggered: chainTurnover };
      if (chainTurnover) {
        var state2 = getTeamState(teamName);
        var removed2 = removeLatestEmployee(sheet, state2);
        if (removed2) result.removedEmployees.push(removed2);
      }
    }
  }

  return { success: true, result: result };
}

function monthEndPoach(teamName, data) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  var state = getTeamState(teamName);

  if (state.period < 3) return { success: true, results: [] };

  var diceValues = data.diceValues; // one per Lv4+ employee
  var results = [];
  var lv4Employees = [];

  state.employees.forEach(function(emp, idx) {
    if (emp.skillLv >= 4 && emp.role !== 'MGR' && emp.role !== 'CXO') {
      lv4Employees.push({ emp: emp, idx: idx });
    }
  });

  // Process in reverse to avoid index shifting
  var toRemove = [];
  for (var i = 0; i < lv4Employees.length; i++) {
    var poached = diceValues[i] === 0;
    results.push({
      empId: lv4Employees[i].emp.id,
      diceValue: diceValues[i],
      poached: poached,
    });
    if (poached) toRemove.push(lv4Employees[i].idx);
  }

  // Remove in reverse order
  toRemove.sort(function(a, b) { return b - a; });
  toRemove.forEach(function(idx) {
    var empRow = 12 + idx;
    sheet.getRange(empRow, 1, 1, 9).clearContent();
  });

  if (toRemove.length > 0) compactEmployees(sheet);

  return { success: true, results: results };
}

function advanceMonth(teamName) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  var state = getTeamState(teamName);

  var newMonth = state.month + 1;
  var newPeriod = state.period;
  var gameOver = false;

  if (newMonth > 12) {
    newMonth = 1;
    newPeriod += 1;
    if (newPeriod > 3) {
      return { success: true, gameEnd: true, finalCash: state.cash };
    }
  }

  // Process ramp/training countdown
  state.employees.forEach(function(emp, idx) {
    var empRow = 12 + idx;

    if (emp.status === 'ランプ' && emp.rampRemaining > 0) {
      var newRamp = emp.rampRemaining - 1;
      sheet.getRange(empRow, 6).setValue(newRamp);
      if (newRamp === 0) {
        sheet.getRange(empRow, 5).setValue('稼働');
        sheet.getRange(empRow, 9).setValue(getCapContribution(emp.skillLv));
      }
    }

    if (emp.status === '研修' && emp.trainingRemaining > 0) {
      var newTraining = emp.trainingRemaining - 1;
      sheet.getRange(empRow, 7).setValue(newTraining);
      if (newTraining === 0) {
        var newLv = Math.min(emp.skillLv + 1, 5);
        sheet.getRange(empRow, 3).setValue(newLv);
        sheet.getRange(empRow, 5).setValue('稼働');
        sheet.getRange(empRow, 9).setValue(getCapContribution(newLv));
      }
    }
  });

  // AI maintenance cost
  var aiMaintenance = [0, 200000, 400000, 700000];
  var cost = aiMaintenance[state.aiLevel];

  // Total monthly cost: salaries + rent + AI maintenance
  var totalSalaries = 0;
  state.employees.forEach(function(emp) { totalSalaries += emp.salary; });
  var rentLookup = {
    '社長宅': 0,
    'コワーキングスペース': 200000,
    '小規模オフィス': 600000,
    '中規模オフィス': 1200000,
    '大規模オフィス': 2200000,
  };
  var rent = rentLookup[state.office] || 0;

  // Revenue
  var totalMrr = 0;
  state.deals.forEach(function(d) { totalMrr += d.mrr; });
  var revenue = Math.round(totalMrr * state.revenueMultiplier);

  var monthlyCost = totalSalaries + rent + cost;
  var profit = revenue - monthlyCost;
  var newCash = state.cash + profit;

  // Cash check
  var negMonths = state.consecutiveNegativeMonths;
  if (newCash < 0) {
    negMonths += 1;
    if (negMonths >= 2) gameOver = true;
  } else {
    negMonths = 0;
  }

  // Record PL history before advancing
  // PL rows: period 1 = row 59-65, period 2 = row 76-82, period 3 = row 93-99
  var plBaseRows = { 1: 59, 2: 76, 3: 93 };
  var plBase = plBaseRows[state.period];
  var col = state.month + 1; // column B=month1, C=month2, etc.
  if (plBase && col >= 2 && col <= 13) {
    sheet.getRange(plBase, col).setValue(revenue);      // 実績売上
    sheet.getRange(plBase + 1, col).setValue(monthlyCost); // 実績コスト
    sheet.getRange(plBase + 2, col).setValue(profit);    // 実績利益
    sheet.getRange(plBase + 3, col).setValue(newCash);   // 残キャッシュ
    sheet.getRange(plBase + 4, col).setValue(state.dealCount); // 案件数
    sheet.getRange(plBase + 5, col).setValue(state.employees.length); // 社員数
  }

  sheet.getRange('B2').setValue(newPeriod);
  sheet.getRange('B3').setValue(newMonth);
  sheet.getRange('B6').setValue(newCash);
  sheet.getRange('B8').setValue(negMonths);

  return {
    success: true,
    period: newPeriod,
    month: newMonth,
    revenue: revenue,
    cost: monthlyCost,
    profit: profit,
    cash: newCash,
    gameOver: gameOver,
    consecutiveNegativeMonths: negMonths,
  };
}

// ===== Office =====

function moveOffice(teamName, officeName) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  var state = getTeamState(teamName);

  var offices = {
    '社長宅': { ratio: 0.7, rent: 0, initial: 0 },
    'コワーキングスペース': { ratio: 0.9, rent: 200000, initial: 400000 },
    '小規模オフィス': { ratio: 1.1, rent: 600000, initial: 3000000 },
    '中規模オフィス': { ratio: 1.2, rent: 1200000, initial: 7000000 },
    '大規模オフィス': { ratio: 1.4, rent: 2200000, initial: 15000000 },
  };

  var office = offices[officeName];
  if (!office) return { error: '不明なオフィス' };

  // Deduct initial cost only if moving to a different office
  var cost = (state.office !== officeName) ? office.initial : 0;
  var newCash = state.cash - cost;

  sheet.getRange('B5').setValue(officeName);
  sheet.getRange('B6').setValue(newCash);
  sheet.getRange('B9').setValue(office.ratio);

  return {
    success: true,
    office: officeName,
    cost: cost,
    newRatio: office.ratio,
  };
}

// ===== Raise / Promote =====

function raisePayOrPromote(teamName, data) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  var state = getTeamState(teamName);

  var empIdx = data.employeeIndex;
  var newSalary = data.newSalary;
  var promote = data.promote; // true if promoting to MGR

  var emp = state.employees[empIdx];
  if (!emp) return { error: '社員が見つかりません' };
  if (newSalary < emp.salary) return { error: '給与は下げられません' };

  var empRow = 12 + empIdx;
  sheet.getRange(empRow, 4).setValue(newSalary);

  if (promote && emp.skillLv >= 4 && newSalary >= 600000) {
    sheet.getRange(empRow, 2).setValue('MGR');
  }

  return { success: true, empId: emp.id, newSalary: newSalary, promoted: promote };
}

// ===== Plan & PL History =====

function savePlan(teamName, period, planData) {
  // planData = { revenue: [12 values], cost: [12 values], profit: [12 values], memo: [12 values] }
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  if (!sheet) return { error: 'チームが見つかりません' };

  // Plan rows: period 1 = row 52-55, period 2 = row 69-72, period 3 = row 86-89
  var planBaseRows = { 1: 52, 2: 69, 3: 86 };
  var base = planBaseRows[period];
  if (!base) return { error: '不正な期' };

  // Write 12 months (columns B-M) + total (column N)
  if (planData.revenue) {
    var revTotal = planData.revenue.reduce(function(a,b){return a+b;}, 0);
    sheet.getRange(base, 2, 1, 12).setValues([planData.revenue]);
    sheet.getRange(base, 14).setValue(revTotal);
  }
  if (planData.cost) {
    var costTotal = planData.cost.reduce(function(a,b){return a+b;}, 0);
    sheet.getRange(base + 1, 2, 1, 12).setValues([planData.cost]);
    sheet.getRange(base + 1, 14).setValue(costTotal);
  }
  if (planData.profit) {
    var profitTotal = planData.profit.reduce(function(a,b){return a+b;}, 0);
    sheet.getRange(base + 2, 2, 1, 12).setValues([planData.profit]);
    sheet.getRange(base + 2, 14).setValue(profitTotal);
  }
  if (planData.memo) {
    sheet.getRange(base + 3, 2, 1, 12).setValues([planData.memo]);
  }
  if (planData.actions) {
    sheet.getRange(base + 4, 2, 1, 12).setValues([planData.actions]);
  }

  return { success: true };
}

function getPlan(teamName, period) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  if (!sheet) return { error: 'チームが見つかりません' };

  var planBaseRows = { 1: 52, 2: 69, 3: 86 };
  var base = planBaseRows[period];
  if (!base) return { error: '不正な期' };

  var data = sheet.getRange(base, 2, 5, 12).getValues();
  return {
    revenue: data[0],
    cost: data[1],
    profit: data[2],
    memo: data[3],
    actions: data[4],
  };
}

function getPlHistory(teamName, period) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  if (!sheet) return { error: 'チームが見つかりません' };

  var plBaseRows = { 1: 59, 2: 76, 3: 93 };
  var base = plBaseRows[period];
  if (!base) return { error: '不正な期' };

  var data = sheet.getRange(base, 2, 7, 12).getValues();
  return {
    revenue: data[0],
    cost: data[1],
    profit: data[2],
    cash: data[3],
    deals: data[4],
    employees: data[5],
    events: data[6],
  };
}

// ===== Bailout (救済) =====

function executeBailout(teamName) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('team_' + teamName);
  var state = getTeamState(teamName);

  if (state.bailoutUsed) return { error: '救済は1ゲームに1回のみです' };

  var penalties = [];

  // 1. 緊急融資 1,000万
  var newCash = state.cash + 10000000;
  sheet.getRange('B6').setValue(newCash);
  sheet.getRange('B8').setValue(0); // reset negative months

  // 2. オフィス1段階降格
  var officeNames = ['社長宅','コワーキングスペース','小規模オフィス','中規模オフィス','大規模オフィス'];
  var currentIdx = officeNames.indexOf(state.office);
  if (currentIdx > 0) {
    var newOffice = officeNames[currentIdx - 1];
    var ratios = [0.7, 0.9, 1.1, 1.2, 1.4];
    sheet.getRange('B5').setValue(newOffice);
    sheet.getRange('B9').setValue(ratios[currentIdx - 1]);
    penalties.push('オフィス降格: ' + state.office + ' → ' + newOffice);
  }

  // 3. 最もLvが高い営業が退職
  var highestIdx = -1;
  var highestLv = 0;
  state.employees.forEach(function(emp, idx) {
    if (emp.role === '営業' && emp.skillLv > highestLv) {
      highestLv = emp.skillLv;
      highestIdx = idx;
    }
  });
  if (highestIdx >= 0) {
    var removed = state.employees[highestIdx];
    var empRow = 12 + highestIdx;
    sheet.getRange(empRow, 1, 1, 9).clearContent();
    compactEmployees(sheet);
    penalties.push('エース退職: ' + removed.id + ' (Lv' + removed.skillLv + ')');
  }

  // 4. AI投資レベル -1
  if (state.aiLevel > 0) {
    var newAI = state.aiLevel - 1;
    sheet.getRange('B4').setValue(newAI);
    penalties.push('AI Lv低下: Lv' + state.aiLevel + ' → Lv' + newAI);
  }

  // 5. 救済フラグを立てる
  sheet.getRange('D1').setValue('Y');

  return {
    success: true,
    event: 'bailout',
    cashAdded: 10000000,
    penalties: penalties,
  };
}

// ===== Helpers =====

function getCapContribution(lv) {
  var table = { 1: 0.5, 2: 1.0, 3: 1.2, 4: 1.5, 5: 2.0 };
  return table[lv] || 0;
}

function removeLatestEmployee(sheet, state) {
  // Find latest non-MGR, non-CXO employee
  for (var i = state.employees.length - 1; i >= 0; i--) {
    var emp = state.employees[i];
    if (emp.role !== 'MGR' && emp.role !== 'CXO') {
      var empRow = 12 + i;
      sheet.getRange(empRow, 1, 1, 9).clearContent();
      compactEmployees(sheet);
      return emp;
    }
  }
  return null;
}

function compactDeals(sheet) {
  var data = sheet.getRange('A27:E39').getValues();
  var nonEmpty = data.filter(function(row) { return row[0] !== '' && row[0] !== null; });
  sheet.getRange('A27:E39').clearContent();
  if (nonEmpty.length > 0) {
    sheet.getRange(27, 1, nonEmpty.length, 5).setValues(nonEmpty);
  }
}

function compactEmployees(sheet) {
  var data = sheet.getRange('A12:I25').getValues();
  var nonEmpty = data.filter(function(row) { return row[0] !== '' && row[0] !== null; });
  sheet.getRange('A12:I25').clearContent();
  if (nonEmpty.length > 0) {
    sheet.getRange(12, 1, nonEmpty.length, 9).setValues(nonEmpty);
  }
}
