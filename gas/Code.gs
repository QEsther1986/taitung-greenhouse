/**
 * 台東老屋・會員小管家 — Google Apps Script 後端
 * 已對齊試算表「老屋專屬數位管家 - 後台資料庫」的實際欄位結構。
 * 使用方式：整份貼到試算表的「擴充功能 → Apps Script」，部署為網頁應用程式。
 */

const SHEET_ID = '1fLMEEOuqD9xsCVAWC6XDHVEXv7FM1-supUbjIkEzKTI'; // 老屋專屬數位管家 - 後台資料庫

// 空房管理日曆：留空字串 = 使用您 Google 帳號的預設日曆（零設定）。
// 若之後建立專用日曆，把日曆 ID 貼進來即可。
const CALENDAR_ID = 'be25872ecdbdbb55673dce571584645d609b5dcb486aa6eaf6a81994d81ea4c1@group.calendar.google.com';

/* ═══════════ 路由 ═══════════ */

function doGet(e) {
  try {
    if (e.parameter.action === 'getMember')       return json(getMember(e.parameter.phone));
    if (e.parameter.action === 'getAvailability') return json(getAvailability(e.parameter.from, e.parameter.to));
    if (e.parameter.action === 'getRooms')        return json(getRooms());
    return json({ ok: false, message: 'unknown action' });
  } catch (err) {
    return json({ ok: false, message: '系統錯誤：' + err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'register')      return json(registerMember(body.data));
    if (body.action === 'submitTopup')   return json(submitTopup(body.data));
    if (body.action === 'submitBooking') return json(submitBooking(body.data));
    return json({ ok: false, message: 'unknown action' });
  } catch (err) {
    return json({ ok: false, message: '系統錯誤：' + err.message });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ═══════════ 工具 ═══════════ */

// 用第一格標題找分頁，不怕分頁改名或順序調動
function findSheet(headerA1) {
  const sheets = SpreadsheetApp.openById(SHEET_ID).getSheets();
  for (const s of sheets) {
    if (String(s.getRange(1, 1).getValue()).trim() === headerA1) return s;
  }
  throw new Error('找不到標題為「' + headerA1 + '」的分頁');
}
const memberSheet  = () => findSheet('會員編號');
const bookingSheet = () => findSheet('申請時間');

// 電話正規化：只留數字，0980-123456 與 0980123456 視為相同
const digits = v => String(v).replace(/\D/g, '');

function fmtDate(v) {
  return v instanceof Date
    ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd') : String(v);
}

function fmtMonth(v) {
  return v instanceof Date
    ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM') : String(v);
}

function calendar() {
  return CALENDAR_ID
    ? CalendarApp.getCalendarById(CALENDAR_ID)
    : CalendarApp.getDefaultCalendar();
}

/* ═══════════ 會員查詢 ═══════════ */
// 會員總表欄位（A~K）：
// 會員編號/真實姓名/LINE暱稱/聯絡電話/匯款帳號後五碼/會費核對狀態/儲值餘額/家庭人數與房型偏好/特殊備註/加入月份/到期月份

function getMember(phone) {
  const rows = memberSheet().getDataRange().getValues();
  const target = digits(phone);
  if (!target) return { ok: false, message: '請提供手機號碼' };

  for (let i = 1; i < rows.length; i++) {
    if (digits(rows[i][3]) === target) {
      return { ok: true, member: {
        memberId:    String(rows[i][0]),
        name:        String(rows[i][1]),
        lineName:    String(rows[i][2]),
        // 一般會員（免費註冊）/ 待確認（會費對帳中）/ 已開通（訂閱制生效）
        status:      String(rows[i][5]) || '一般會員',
        balance:     Number(rows[i][6]) || 0,
        expiryMonth: fmtMonth(rows[i][10]),
        records:     getRecords(String(rows[i][1])),
      }};
    }
  }
  return { ok: false, message: '查無此帳號，請確認號碼，或先免費註冊會員。' };
}

/* ═══════════ 免費註冊（一般會員） ═══════════ */

function registerMember(d) {
  const sheet = memberSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (digits(rows[i][3]) === digits(d.phone)) {
      return { ok: false, message: '此手機號碼已註冊過，請直接以手機號碼登入。' };
    }
  }
  const nextId = 'M' + String(rows.length).padStart(3, '0');
  sheet.appendRow([
    nextId, d.name, d.lineName, "'" + d.phone, '',
    '一般會員',                     // 未訂閱；匯款回報後改為 待確認 → 已開通
    0, '', '',
    Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM'),
    '',                             // 到期月份：訂閱開通時由管家填入
  ]);
  return { ok: true, member: {
    memberId: nextId, name: d.name, lineName: d.lineName,
    status: '一般會員', balance: 0, expiryMonth: '', records: [],
  }};
}

// 訂房紀錄欄位（A~J）：
// 申請時間/預約會員姓名/預計入住日期/預計退房日期/訂購房型/人數(幾位大人)/人數(幾位小孩)/本次扣款金額/對帳與確認狀態/密碼派發狀態

function getRecords(name) {
  const rows = bookingSheet().getDataRange().getValues();
  return rows.slice(1)
    .filter(r => String(r[1]).trim() === name.trim())
    .slice(-5)                                   // 最近 5 筆
    .map(r => ({
      checkIn:  fmtDate(r[2]),
      checkOut: fmtDate(r[3]),
      party:    `${String(r[4]).split('｜')[0]}・${r[5]}大${r[6]}小`,
      amount:   Number(r[7]) || 0,
      status:   String(r[8]) || '待確認',
    }));
}

/* ═══════════ 房型開放設定 ═══════════ */
// 管理者在試算表「房型設定」分頁把狀態改成「關閉」，該房型即暫停開放。
// 分頁不存在時第一次呼叫會自動建立（預設全部開放）。

const ALL_ROOMS = ['和室', '套房', '雙人房・雙人床', '雙人房・兩單床', '四人房', '包棟'];

function getRooms() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = null;
  for (const s of ss.getSheets()) {
    if (String(s.getRange(1, 1).getValue()).trim() === '房型') { sheet = s; break; }
  }
  if (!sheet) {
    sheet = ss.insertSheet('房型設定');
    sheet.appendRow(['房型', '狀態（開放/關閉）', '備註']);
    ALL_ROOMS.forEach(r => sheet.appendRow([r, '開放', '']));
    sheet.setFrozenRows(1);
  }
  const rows = sheet.getDataRange().getValues();
  const closed = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).indexOf('關') !== -1) closed.push(String(rows[i][0]).trim());
  }
  return { ok: true, closed: closed };
}

/* ═══════════ 空房查詢（Google 日曆，每房各自計算） ═══════════ */
// 事件標題格式：【訂房|房型】姓名 人數，例：【訂房|和室】陳美惠 2大1國小
// 回傳 { ok: true, booked: { 房型: [日期...], ... } }

function getAvailability(from, to) {
  const events = calendar().getEvents(new Date(from), new Date(to + 'T23:59:59'));
  const booked = {};
  events.forEach(ev => {
    // 看【訂房|房型】與【關房|房型】格式的事件（關房 = 管理者臨時封房，
    // 例如維修週在日曆建「【關房|和室】」整日事件即可），其他私人行程不受影響
    const m = ev.getTitle().match(/^【(?:訂房|關房)\|([^】]+)】/);
    if (!m) return;
    const room = m[1];
    // 事件每跨一晚就標成已訂；退房日早上不佔房
    for (let d = new Date(ev.getStartTime()); d < ev.getEndTime(); d.setDate(d.getDate() + 1)) {
      const iso = Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd');
      (booked[room] = booked[room] || {})[iso] = true;
    }
  });
  const out = {};
  Object.keys(booked).forEach(r => { out[r] = Object.keys(booked[r]).sort(); });
  return { ok: true, booked: out };
}

/* ═══════════ 送出訂房 ═══════════ */

function submitBooking(d) {
  // 房型欄記錄計價身分（訂閱8折 / 一般會員原價）
  const tierLabel = { sub: '訂閱8折', expired: '訂閱到期原價', pending: '對帳中原價', basic: '一般會員原價' };
  const roomInfo = d.roomType + '｜' + (tierLabel[d.memberTier] || '原價');

  // ── 新增:負餘額(欠款中)擋訂房 ──
  if (memberRow === -1) {
    return { ok: false, message: '查無會員資料,請重新登入。' };
  }
  if (balanceBefore < 0) {
    return { ok: false, message: `您有未補足的房費 ${-balanceBefore} 元,請先補款後再訂房。` };
  }

  // 餘額為正,正常扣款(可扣成負數)
  balanceAfter = balanceBefore - (Number(d.amount) || 0);
  sheet.getRange(memberRow + 1, 7).setValue(balanceAfter);

  bookingSheet().appendRow([
    new Date(), d.name, d.checkIn, d.checkOut, roomInfo,
    Number(d.adults) || 0, Number(d.kids) || 0,
    Number(d.amount) || 0, '待確認', '未派發',
  ]);
  calendar().createAllDayEvent(
    `【訂房|${d.roomType}】${d.name} ${d.party}`,
    new Date(d.checkIn), new Date(d.checkOut)
  );

  return {
    ok: true,
    amount: Number(d.amount) || 0,
    balanceAfter: balanceAfter,
    shortfall: balanceAfter < 0 ? -balanceAfter : 0,
  };

  /*
  bookingSheet().appendRow([
    new Date(),                                      // 申請時間
    d.name,                                          // 預約會員姓名
    d.checkIn,                                       // 預計入住日期
    d.checkOut,                                      // 預計退房日期
    roomInfo,                                        // 訂購房型
    Number(d.adults) || 0,                           // 人數(幾位大人)
    Number(d.kids) || 0,                             // 人數(幾位小孩=國小+6歲以下)
    Number(d.amount) || 0,                           // 本次扣款金額
    '待確認',                                         // 對帳與確認狀態
    '未派發',                                         // 密碼派發狀態
  ]);
  calendar().createAllDayEvent(
    `【訂房|${d.roomType}】${d.name} ${d.party}`,
    new Date(d.checkIn), new Date(d.checkOut)
  );
  return { ok: true };
  */
}

/* ═══════════ 入會 / 匯款回報 ═══════════ */

function submitTopup(d) {
  const sheet = memberSheet();
  const rows = sheet.getDataRange().getValues();

  // 同一支手機已存在 → 一般會員升級訂閱 / 訂閱會員續約補繳，不重複建檔
  for (let i = 1; i < rows.length; i++) {
    if (digits(rows[i][3]) === digits(d.phone)) {
      sheet.getRange(i + 1, 5).setValue("'" + d.last5);         // 更新後五碼
      if (String(rows[i][5]) !== '已開通') {
        sheet.getRange(i + 1, 6).setValue('待確認');            // 一般會員 → 進入對帳
      }
      sheet.getRange(i + 1, 9).setValue(
        (rows[i][8] ? rows[i][8] + '\n' : '') +
        `[${Utilities.formatDate(new Date(), 'Asia/Taipei', 'MM/dd')} 匯款回報] 後五碼 ${d.last5}｜${d.note || ''}`
      );
      notifyTopup(d, false); //通知已匯款要查帳
      return { ok: true };
    }
  }

  // 新會員：自動編號 M001, M002, ...
  const nextId = 'M' + String(rows.length).padStart(3, '0');
  const now = new Date();
  const expiry = new Date(now); expiry.setFullYear(expiry.getFullYear() + 1);

  sheet.appendRow([
    nextId,
    d.name,
    d.lineName,
    "'" + d.phone,                    // 前置 ' 保住開頭的 0
    "'" + d.last5,
    '待確認',
    0,                                // 對帳開通後由管家改為 3000
    `${d.adults}大 ${d.kids}小${d.familyNote ? '（' + d.familyNote + '）' : ''}`,
    d.note || '',
    Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM'),
    Utilities.formatDate(expiry, 'Asia/Taipei', 'yyyy-MM'),
  ]);
  return { ok: true };
}

/* ═══════════ 匯款回報通知管家 ═══════════ */
function notifyTopup(d, isNew) {
  const to = Session.getEffectiveUser().getEmail();  // 寄給你自己(部署帳號)
  const subject = `💰 匯款回報：${d.name}（後五碼 ${d.last5}）`;
  const body =
    `有一筆匯款回報,請查帳核對:\n\n` +
    `姓名:${d.name}\n` +
    `LINE:${d.lineName || '(未填)'}\n` +
    `手機:${d.phone}\n` +
    `匯款後五碼:${d.last5}\n` +
    `類型:${isNew ? '新入會' : '既有會員補款/續約'}\n` +
    `備註:${d.note || '(無)'}\n` +
    `時間:${Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm')}\n\n` +
    `請至後台資料庫核對後,將該會員狀態改為「已開通」、餘額填入。`;
  MailApp.sendEmail(to, subject, body);
}
