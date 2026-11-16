/**
 * 후원주점 서빙 시스템 — 백엔드 (Google Apps Script, 구글 시트 연결형)
 *
 * [설치 순서]
 *  1. 구글 시트를 열고 [확장 프로그램] > [Apps Script] 선택
 *  2. 기본 Code.gs 내용을 모두 지우고 이 파일 전체를 붙여넣은 뒤 저장
 *  3. 상단 함수 목록에서 'setup' 선택 후 [실행] (최초 1회 권한 승인 필요)
 *  4. [배포] > [새 배포] > 유형: 웹 앱
 *       - 다음 사용자 인증 정보로 실행: 나
 *       - 액세스 권한이 있는 사용자: 모든 사용자
 *     → 발급된 웹 앱 URL(https://script.google.com/macros/s/.../exec)을 config.js에 입력
 *  ※ 코드를 수정한 경우 [배포] > [배포 관리] > 편집 > 버전: 새 버전 으로 재배포해야 반영됩니다.
 */

/* ───────────────────────── 기본 정의 ───────────────────────── */

var SHEET = {
  CONFIG: '설정',
  MENU: '메뉴',
  TABLES: '테이블',
  ORDERS: '주문',
  PAYMENTS: '결제',
  SESSIONS: '이용내역'
};

var HEADERS = {
  '설정': ['항목', '값', '설명'],
  '메뉴': ['분류', '메뉴명', '가격', '상태', '비고'],
  '테이블': ['테이블', '세션ID', '상태', '인원', '수저세팅', '착석시각', '메모', '수정시각'],
  '주문': ['주문ID', '주문시각', '테이블', '세션ID', '분류', '메뉴명', '단가', '수량', '금액', '상태', '처리시각', '입력자', '요청ID', '비고'],
  '결제': ['결제ID', '결제시각', '테이블', '세션ID', '금액', '수단', '입력자', '비고'],
  '이용내역': ['세션ID', '테이블', '인원', '착석시각', '정리시각', '주문합계', '결제합계', '차액', '메모']
};

var DEFAULT_CONFIG = [
  ['행사명', '후원주점', '화면 상단에 표시되는 행사 이름'],
  ['테이블수', 30, '운영 테이블 수 (1번부터 순서대로 표시)'],
  ['접속비밀번호', '1234', '웹앱 접속 시 입력하는 비밀번호 (행사 전 반드시 변경, 비워두면 비밀번호 없이 접속)']
];

var SAMPLE_MENU = [
  ['안주', '해물파전', 15000, '판매', '예시 메뉴 — 실제 메뉴로 수정'],
  ['안주', '김치전', 12000, '판매', '예시'],
  ['안주', '두부김치', 13000, '판매', '예시'],
  ['안주', '어묵탕', 13000, '판매', '예시'],
  ['안주', '골뱅이무침', 15000, '판매', '예시'],
  ['안주', '계란말이', 10000, '판매', '예시'],
  ['안주', '마른안주', 10000, '판매', '예시'],
  ['주류', '소주', 5000, '판매', '예시'],
  ['주류', '맥주', 5000, '판매', '예시'],
  ['주류', '막걸리', 5000, '판매', '예시'],
  ['음료', '콜라', 2000, '판매', '예시'],
  ['음료', '사이다', 2000, '판매', '예시']
];

var ORDER_STATUS = ['대기', '서빙완료', '취소'];
var MENU_STATUS = ['판매', '품절', '숨김'];

/* ───────────────────────── 시트 메뉴·초기 설정 ───────────────────────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('후원주점')
    .addItem('① 초기 설정 (시트 생성)', 'setup')
    .addSeparator()
    .addItem('행사 데이터 초기화 (주문·결제·이용내역 삭제)', 'resetEventData')
    .addToUi();
}

/** 시트에서 직접 메뉴·설정을 수정하면 웹앱 화면에 바로 반영되도록 갱신번호를 올림 */
function onEdit(e) {
  try { bumpRev_(); } catch (err) {}
}

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(function (name) {
    var s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    var h = HEADERS[name];
    if (s.getLastRow() === 0) s.getRange(1, 1, 1, h.length).setValues([h]);
    s.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#f1f3f4');
    s.setFrozenRows(1);
  });

  var cfg = ss.getSheetByName(SHEET.CONFIG);
  cfg.getRange('B:B').setNumberFormat('@'); // 비밀번호 앞자리 0 보존
  if (cfg.getLastRow() < 2) cfg.getRange(2, 1, DEFAULT_CONFIG.length, 3).setValues(DEFAULT_CONFIG);

  var menu = ss.getSheetByName(SHEET.MENU);
  if (menu.getLastRow() < 2) menu.getRange(2, 1, SAMPLE_MENU.length, 5).setValues(SAMPLE_MENU);
  menu.getRange('C2:C').setNumberFormat('#,##0');
  menu.getRange('D2:D').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(MENU_STATUS, true).build()
  );

  ss.getSheetByName(SHEET.ORDERS).getRange('G2:I').setNumberFormat('#,##0');
  ss.getSheetByName(SHEET.PAYMENTS).getRange('E2:E').setNumberFormat('#,##0');
  ss.getSheetByName(SHEET.SESSIONS).getRange('F2:H').setNumberFormat('#,##0');

  ensureTables_();

  // 기본으로 생성된 빈 시트(시트1, Sheet1 등) 정리
  ss.getSheets().forEach(function (s) {
    if (!HEADERS[s.getName()] && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  bumpRev_();
  try {
    SpreadsheetApp.getUi().alert('초기 설정이 완료되었습니다.\n「설정」 시트에서 행사명·테이블수·접속비밀번호를, 「메뉴」 시트에서 메뉴와 가격을 확인해 주십시오.');
  } catch (err) {}
}

function resetEventData() {
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (err) {}
  if (ui) {
    var res = ui.alert('행사 데이터 초기화', '주문·결제·이용내역이 모두 삭제되고 모든 테이블이 빈자리로 바뀝니다.\n(메뉴와 설정은 유지됩니다)\n계속하시겠습니까?', ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    [SHEET.ORDERS, SHEET.PAYMENTS, SHEET.SESSIONS, SHEET.TABLES].forEach(function (name) {
      var s = sh_(name);
      if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
    });
    ensureTables_();
    bumpRev_();
  } finally {
    lock.releaseLock();
  }
  if (ui) ui.alert('초기화가 완료되었습니다.');
}

/* ───────────────────────── 웹앱 진입점 ───────────────────────── */

function doGet(e) {
  return handle_((e && e.parameter) || {});
}

function doPost(e) {
  var p;
  try {
    p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: '요청 형식이 올바르지 않습니다.' });
  }
  return handle_(p);
}

var WRITE_ACTIONS = {
  updateTable: actUpdateTable_,
  order: actOrder_,
  setOrderStatus: actSetOrderStatus_,
  pay: actPay_,
  clearTable: actClearTable_,
  moveTable: actMoveTable_,
  setMenuStatus: actSetMenuStatus_
};

function handle_(p) {
  try {
    var cfg = getConfig_();
    if (cfg.pin && String(p.pin || '').trim() !== cfg.pin) {
      return json_({ ok: false, code: 'PIN', error: '접속 비밀번호가 올바르지 않습니다.' });
    }
    var action = p.action || 'state';
    if (action === 'state') return json_({ ok: true, state: getStateCached_() });
    if (action === 'summary') return json_({ ok: true, summary: getSummary_() });

    var fn = WRITE_ACTIONS[action];
    if (!fn) throw new Error('알 수 없는 요청입니다: ' + action);

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    var result;
    try {
      result = fn(p) || {};
      bumpRev_();
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    var out = { ok: true, state: buildState_() };
    Object.keys(result).forEach(function (k) { out[k] = result[k]; });
    return json_(out);
  } catch (err) {
    return json_({ ok: false, code: (err && err.code) || 'ERR', error: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ───────────────────────── 쓰기 처리 ───────────────────────── */

/** 인원·수저세팅·메모 변경 (빈 테이블이면 착석 처리) */
function actUpdateTable_(p) {
  var t = getTable_(p.table);
  if (!isActive_(t)) {
    if (p.memo == null && p.spoon == null && !(Number(p.people) > 0)) {
      throw new Error('빈 테이블입니다. 인원을 입력하거나 주문을 넣으면 착석 처리됩니다.');
    }
    t = ensureSession_(p.table);
  }
  applyTableFields_(t, p);
  return {};
}

function actOrder_(p) {
  var items = (p.items || [])
    .map(function (i) { return { name: String(i.name || '').trim(), qty: Math.floor(Number(i.qty)) }; })
    .filter(function (i) { return i.name && i.qty > 0; });
  if (!items.length) throw new Error('주문할 메뉴를 선택해 주십시오.');
  items.forEach(function (i) { if (i.qty > 99) throw new Error('한 메뉴의 수량은 99개를 넘을 수 없습니다.'); });

  if (p.requestId) {
    var dup = readObjects_(SHEET.ORDERS).some(function (o) { return String(o['요청ID']) === String(p.requestId); });
    if (dup) return { duplicate: true };
  }

  var menu = getMenuMap_();
  items.forEach(function (i) {
    var m = menu[i.name];
    if (!m) throw new Error('「' + i.name + '」 메뉴를 찾을 수 없습니다.');
    if (m.status !== '판매') throw new Error('「' + i.name + '」은(는) 현재 ' + m.status + ' 상태입니다.');
    i.m = m;
  });

  var t = ensureSession_(p.table);
  if (p.people != null || p.spoon != null) applyTableFields_(t, p);

  var now = new Date();
  var by = cleanText_(p.by, 20);
  var note = cleanText_(p.note, 100);
  var rows = items.map(function (i) {
    return {
      '주문ID': newId_('O'),
      '주문시각': now,
      '테이블': Number(p.table),
      '세션ID': t['세션ID'],
      '분류': i.m.cat,
      '메뉴명': i.name,
      '단가': i.m.price,
      '수량': i.qty,
      '금액': i.m.price * i.qty,
      '상태': '대기',
      '처리시각': '',
      '입력자': by,
      '요청ID': p.requestId || '',
      '비고': note
    };
  });
  appendObjects_(SHEET.ORDERS, rows);
  return { added: rows.length };
}

function actSetOrderStatus_(p) {
  var status = String(p.status || '');
  if (ORDER_STATUS.indexOf(status) < 0) throw new Error('주문 상태 값이 올바르지 않습니다.');
  var ids = {};
  (p.ids || []).forEach(function (id) { ids[String(id)] = true; });
  var rows = readObjects_(SHEET.ORDERS).filter(function (o) { return ids[String(o['주문ID'])]; });
  if (!rows.length) throw new Error('해당 주문을 찾을 수 없습니다.');
  var now = new Date();
  rows.forEach(function (o) {
    updateRow_(SHEET.ORDERS, o, { '상태': status, '처리시각': status === '대기' ? '' : now });
  });
  return { updated: rows.length };
}

function actPay_(p) {
  var t = getTable_(p.table);
  if (!isActive_(t)) throw new Error('이용 중인 테이블이 아닙니다.');
  var amount = Math.round(Number(p.amount));
  if (!amount || !isFinite(amount)) throw new Error('결제 금액을 입력해 주십시오.');
  var method = p.method === '계좌이체' ? '계좌이체' : '현금';
  appendObjects_(SHEET.PAYMENTS, [{
    '결제ID': newId_('P'),
    '결제시각': new Date(),
    '테이블': Number(p.table),
    '세션ID': t['세션ID'],
    '금액': amount,
    '수단': method,
    '입력자': cleanText_(p.by, 20),
    '비고': cleanText_(p.note, 100)
  }]);
  return {};
}

function actClearTable_(p) {
  var t = getTable_(p.table);
  if (!isActive_(t)) throw new Error('이미 빈 테이블입니다.');
  var sid = String(t['세션ID']);
  var orders = readObjects_(SHEET.ORDERS).filter(function (o) { return String(o['세션ID']) === sid && o['상태'] !== '취소'; });
  var total = sum_(orders, '금액');
  var paid = sum_(readObjects_(SHEET.PAYMENTS).filter(function (x) { return String(x['세션ID']) === sid; }), '금액');
  var pending = orders.filter(function (o) { return o['상태'] === '대기'; }).length;
  if (!p.force && (total !== paid || pending > 0)) {
    var err = new Error('미결제 금액 또는 미서빙 주문이 남아 있습니다.');
    err.code = 'CONFIRM';
    throw err;
  }
  appendObjects_(SHEET.SESSIONS, [{
    '세션ID': sid,
    '테이블': Number(t['테이블']),
    '인원': Number(t['인원']) || 0,
    '착석시각': t['착석시각'],
    '정리시각': new Date(),
    '주문합계': total,
    '결제합계': paid,
    '차액': total - paid,
    '메모': t['메모'] || ''
  }]);
  updateRow_(SHEET.TABLES, t, {
    '세션ID': '', '상태': '빈자리', '인원': '', '수저세팅': '', '착석시각': '', '메모': '', '수정시각': new Date()
  });
  return {};
}

function actMoveTable_(p) {
  var from = getTable_(p.table);
  var to = getTable_(p.to);
  if (!isActive_(from)) throw new Error('이동할 테이블이 이용 중이 아닙니다.');
  if (isActive_(to)) throw new Error(Number(p.to) + '번 테이블은 이미 이용 중입니다.');
  var sid = String(from['세션ID']);
  var toNo = Number(to['테이블']);
  updateRow_(SHEET.TABLES, to, {
    '세션ID': sid, '상태': '이용중', '인원': from['인원'], '수저세팅': from['수저세팅'],
    '착석시각': from['착석시각'], '메모': from['메모'], '수정시각': new Date()
  });
  updateRow_(SHEET.TABLES, from, {
    '세션ID': '', '상태': '빈자리', '인원': '', '수저세팅': '', '착석시각': '', '메모': '', '수정시각': new Date()
  });
  [SHEET.ORDERS, SHEET.PAYMENTS].forEach(function (name) {
    readObjects_(name).forEach(function (o) {
      if (String(o['세션ID']) === sid) updateRow_(name, o, { '테이블': toNo });
    });
  });
  return {};
}

function actSetMenuStatus_(p) {
  var status = String(p.status || '');
  if (MENU_STATUS.indexOf(status) < 0) throw new Error('메뉴 상태 값이 올바르지 않습니다.');
  var row = readObjects_(SHEET.MENU).filter(function (m) { return String(m['메뉴명']).trim() === String(p.name); })[0];
  if (!row) throw new Error('메뉴를 찾을 수 없습니다.');
  updateRow_(SHEET.MENU, row, { '상태': status });
  return {};
}

/* ───────────────────────── 조회 ───────────────────────── */

function getStateCached_() {
  var rev = getRev_();
  var cache = CacheService.getScriptCache();
  var c = cache.get('state');
  if (c) {
    try {
      var s = JSON.parse(c);
      if (s.rev === rev) { s.serverTime = Date.now(); return s; }
    } catch (err) {}
  }
  var st = buildState_(rev);
  try { cache.put('state', JSON.stringify(st), 15); } catch (err) {}
  return st;
}

function buildState_(rev) {
  if (rev == null) rev = getRev_();
  var cfg = getConfig_();
  var menu = readMenu_().filter(function (m) { return m.status !== '숨김'; });
  var tRows = readObjects_(SHEET.TABLES);

  var activeBySid = {};
  var byNo = {};
  tRows.forEach(function (t) {
    var no = Number(t['테이블']);
    if (!no) return;
    byNo[no] = t;
    if (isActive_(t)) activeBySid[String(t['세션ID'])] = no;
  });

  var ordersBySid = {};
  readObjects_(SHEET.ORDERS).forEach(function (o) {
    var sid = String(o['세션ID']);
    if (activeBySid[sid] == null) return;
    (ordersBySid[sid] = ordersBySid[sid] || []).push({
      id: String(o['주문ID']),
      time: toMs_(o['주문시각']),
      cat: String(o['분류'] || ''),
      name: String(o['메뉴명']),
      price: Number(o['단가']) || 0,
      qty: Number(o['수량']) || 0,
      amount: Number(o['금액']) || 0,
      status: String(o['상태'] || '대기'),
      doneAt: toMs_(o['처리시각']),
      by: String(o['입력자'] || ''),
      note: String(o['비고'] || '')
    });
  });

  var paysBySid = {};
  readObjects_(SHEET.PAYMENTS).forEach(function (x) {
    var sid = String(x['세션ID']);
    if (activeBySid[sid] == null) return;
    (paysBySid[sid] = paysBySid[sid] || []).push({
      time: toMs_(x['결제시각']),
      amount: Number(x['금액']) || 0,
      method: String(x['수단'] || '현금'),
      by: String(x['입력자'] || '')
    });
  });

  var maxNo = cfg.tableCount;
  Object.keys(byNo).forEach(function (no) {
    if (isActive_(byNo[no]) && Number(no) > maxNo) maxNo = Number(no);
  });

  var tables = [];
  for (var no = 1; no <= maxNo; no++) {
    var t = byNo[no];
    if (!t || !isActive_(t)) {
      tables.push({ no: no, active: false, orders: [], payments: [], total: 0, paid: 0, balance: 0, pending: 0, people: 0, spoon: false, memo: '' });
      continue;
    }
    var sid = String(t['세션ID']);
    var orders = ordersBySid[sid] || [];
    var pays = paysBySid[sid] || [];
    var valid = orders.filter(function (o) { return o.status !== '취소'; });
    var total = valid.reduce(function (a, o) { return a + o.amount; }, 0);
    var paid = pays.reduce(function (a, x) { return a + x.amount; }, 0);
    tables.push({
      no: no,
      active: true,
      sessionId: sid,
      people: Number(t['인원']) || 0,
      spoon: t['수저세팅'] === '완료',
      seatedAt: toMs_(t['착석시각']),
      memo: String(t['메모'] || ''),
      orders: orders,
      payments: pays,
      total: total,
      paid: paid,
      balance: total - paid,
      pending: valid.filter(function (o) { return o.status === '대기'; }).length
    });
  }

  return {
    rev: rev,
    eventName: cfg.eventName,
    tableCount: cfg.tableCount,
    menu: menu,
    tables: tables,
    serverTime: Date.now()
  };
}

function getSummary_() {
  var cfg = getConfig_();
  var menuList = readMenu_();
  var sessions = {};

  readObjects_(SHEET.SESSIONS).forEach(function (s) {
    var sid = String(s['세션ID']);
    sessions[sid] = {
      sessionId: sid, table: Number(s['테이블']), people: Number(s['인원']) || 0,
      seatedAt: toMs_(s['착석시각']), closedAt: toMs_(s['정리시각']), open: false,
      total: 0, paid: 0, items: 0
    };
  });
  readObjects_(SHEET.TABLES).forEach(function (t) {
    if (!isActive_(t)) return;
    var sid = String(t['세션ID']);
    sessions[sid] = {
      sessionId: sid, table: Number(t['테이블']), people: Number(t['인원']) || 0,
      seatedAt: toMs_(t['착석시각']), closedAt: null, open: true,
      total: 0, paid: 0, items: 0
    };
  });
  function sess(sid, table) {
    if (!sessions[sid]) sessions[sid] = { sessionId: sid, table: Number(table), people: 0, seatedAt: null, closedAt: null, open: false, total: 0, paid: 0, items: 0 };
    return sessions[sid];
  }

  var menuAgg = {};
  menuList.forEach(function (m, i) { menuAgg[m.name] = { cat: m.cat, name: m.name, price: m.price, qty: 0, amount: 0, order: i }; });

  var orderTotal = 0, cancelCount = 0, cancelAmount = 0, pendingCount = 0, itemsCount = 0;
  readObjects_(SHEET.ORDERS).forEach(function (o) {
    var amt = Number(o['금액']) || 0;
    var qty = Number(o['수량']) || 0;
    if (o['상태'] === '취소') { cancelCount++; cancelAmount += amt; return; }
    var os = sessions[String(o['세션ID'])];
    if (o['상태'] === '대기' && os && os.open) pendingCount++;
    orderTotal += amt;
    itemsCount += qty;
    var s = sess(String(o['세션ID']), o['테이블']);
    s.total += amt;
    s.items += qty;
    var name = String(o['메뉴명']);
    if (!menuAgg[name]) menuAgg[name] = { cat: String(o['분류'] || ''), name: name, price: Number(o['단가']) || 0, qty: 0, amount: 0, order: 9999 };
    menuAgg[name].qty += qty;
    menuAgg[name].amount += amt;
  });

  var paidTotal = 0, byMethod = { '현금': 0, '계좌이체': 0 }, payCount = 0;
  readObjects_(SHEET.PAYMENTS).forEach(function (x) {
    var amt = Number(x['금액']) || 0;
    paidTotal += amt;
    payCount++;
    var m = x['수단'] === '계좌이체' ? '계좌이체' : '현금';
    byMethod[m] += amt;
    sess(String(x['세션ID']), x['테이블']).paid += amt;
  });

  var list = Object.keys(sessions).map(function (k) {
    var s = sessions[k];
    s.balance = s.total - s.paid;
    return s;
  }).filter(function (s) { return s.open || s.total || s.paid; });
  list.sort(function (a, b) { return (a.seatedAt || 0) - (b.seatedAt || 0); });

  var outstanding = 0, overpaid = 0, peopleTotal = 0;
  list.forEach(function (s) {
    if (s.balance > 0) outstanding += s.balance;
    if (s.balance < 0) overpaid += -s.balance;
    peopleTotal += s.people;
  });

  var menuRows = Object.keys(menuAgg).map(function (k) { return menuAgg[k]; })
    .sort(function (a, b) { return a.order - b.order; });

  return {
    eventName: cfg.eventName,
    generatedAt: Date.now(),
    orderTotal: orderTotal,
    itemsCount: itemsCount,
    pendingCount: pendingCount,
    cancelCount: cancelCount,
    cancelAmount: cancelAmount,
    paidTotal: paidTotal,
    payCount: payCount,
    byMethod: byMethod,
    outstanding: outstanding,
    overpaid: overpaid,
    sessionCount: list.length,
    openCount: list.filter(function (s) { return s.open; }).length,
    peopleTotal: peopleTotal,
    menu: menuRows,
    sessions: list
  };
}

/* ───────────────────────── 시트 접근 도우미 ───────────────────────── */

function sh_(name) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('「' + name + '」 시트가 없습니다. 구글 시트 메뉴 [후원주점] > [① 초기 설정]을 먼저 실행해 주십시오.');
  return s;
}

function readObjects_(name) {
  var s = sh_(name);
  var last = s.getLastRow();
  if (last < 2) return [];
  var H = HEADERS[name];
  var vals = s.getRange(2, 1, last - 1, H.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var o = { _row: i + 2 };
    for (var j = 0; j < H.length; j++) o[H[j]] = vals[i][j];
    out.push(o);
  }
  return out;
}

function appendObjects_(name, objs) {
  if (!objs.length) return;
  var s = sh_(name);
  var H = HEADERS[name];
  var rows = objs.map(function (o) { return H.map(function (h) { return o[h] === undefined || o[h] === null ? '' : o[h]; }); });
  s.getRange(s.getLastRow() + 1, 1, rows.length, H.length).setValues(rows);
}

function updateRow_(name, obj, changes) {
  Object.keys(changes).forEach(function (k) { obj[k] = changes[k]; });
  var H = HEADERS[name];
  var row = H.map(function (h) { return obj[h] === undefined || obj[h] === null ? '' : obj[h]; });
  sh_(name).getRange(obj._row, 1, 1, H.length).setValues([row]);
}

function getConfig_() {
  var map = {};
  readObjects_(SHEET.CONFIG).forEach(function (r) { map[String(r['항목']).trim()] = r['값']; });
  var count = Math.floor(Number(map['테이블수'])) || 20;
  return {
    eventName: String(map['행사명'] || '후원주점'),
    tableCount: Math.min(Math.max(count, 1), 200),
    pin: String(map['접속비밀번호'] == null ? '' : map['접속비밀번호']).trim()
  };
}

function readMenu_() {
  return readObjects_(SHEET.MENU)
    .filter(function (m) { return String(m['메뉴명']).trim(); })
    .map(function (m) {
      var st = String(m['상태'] || '판매').trim();
      return {
        cat: String(m['분류'] || '기타').trim() || '기타',
        name: String(m['메뉴명']).trim(),
        price: Number(m['가격']) || 0,
        status: MENU_STATUS.indexOf(st) >= 0 ? st : '판매'
      };
    });
}

function getMenuMap_() {
  var map = {};
  readMenu_().forEach(function (m) { map[m.name] = m; });
  return map;
}

function ensureTables_() {
  var cfg = getConfig_();
  var have = {};
  readObjects_(SHEET.TABLES).forEach(function (t) { have[Number(t['테이블'])] = true; });
  var add = [];
  for (var no = 1; no <= cfg.tableCount; no++) {
    if (!have[no]) add.push({ '테이블': no, '상태': '빈자리' });
  }
  appendObjects_(SHEET.TABLES, add);
}

function getTable_(no) {
  no = Number(no);
  if (!(no >= 1) || Math.floor(no) !== no) throw new Error('테이블 번호가 올바르지 않습니다.');
  var t = readObjects_(SHEET.TABLES).filter(function (r) { return Number(r['테이블']) === no; })[0];
  if (t) return t;
  appendObjects_(SHEET.TABLES, [{ '테이블': no, '상태': '빈자리' }]);
  return readObjects_(SHEET.TABLES).filter(function (r) { return Number(r['테이블']) === no; })[0];
}

function isActive_(t) {
  return t && t['상태'] === '이용중' && String(t['세션ID'] || '') !== '';
}

function ensureSession_(no) {
  var t = getTable_(no);
  if (isActive_(t)) return t;
  var now = new Date();
  updateRow_(SHEET.TABLES, t, {
    '세션ID': 'T' + Number(no) + '-' + now.getTime().toString(36).toUpperCase(),
    '상태': '이용중',
    '인원': 0,
    '수저세팅': '미완료',
    '착석시각': now,
    '메모': '',
    '수정시각': now
  });
  return t;
}

function applyTableFields_(t, p) {
  var ch = {};
  if (p.people != null && p.people !== '') ch['인원'] = Math.max(0, Math.min(99, Math.floor(Number(p.people)) || 0));
  if (p.spoon != null) ch['수저세팅'] = p.spoon ? '완료' : '미완료';
  if (p.memo != null) ch['메모'] = cleanText_(p.memo, 200);
  ch['수정시각'] = new Date();
  updateRow_(SHEET.TABLES, t, ch);
}

function getRev_() {
  return Number(PropertiesService.getScriptProperties().getProperty('rev')) || 0;
}

function bumpRev_() {
  var props = PropertiesService.getScriptProperties();
  var old = Number(props.getProperty('rev')) || 0;
  props.setProperty('rev', String(Math.max(Date.now(), old + 1)));
}

function newId_(prefix) {
  return prefix + Date.now().toString(36).toUpperCase() + Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
}

function toMs_(v) {
  if (v === '' || v == null) return null;
  if (typeof v.getTime === 'function') return v.getTime();
  var d = new Date(v).getTime();
  return isNaN(d) ? null : d;
}

function sum_(rows, key) {
  return rows.reduce(function (a, r) { return a + (Number(r[key]) || 0); }, 0);
}

function cleanText_(v, max) {
  return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
}
