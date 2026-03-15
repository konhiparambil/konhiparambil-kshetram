// ═══ PASTE THIS ENTIRE FILE INTO GOOGLE APPS SCRIPT ═══

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getAll') return getAllData();
  if (action === 'getAnalytics') return getAnalyticsData();
  if (action === 'logVisit') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Analytics');
    const deviceId = e.parameter.deviceId || '';
    const ts = e.parameter.ts || Date.now();
    const ua = e.parameter.ua || '';
    const today = new Date().toISOString().split('T')[0];
    sheet.appendRow([deviceId, today, ts, ua]);
    return ContentService.createTextOutput('{"ok":true}').setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('{}').setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = JSON.parse(e.postData.contents);
  const action = data.action;
  const evSheet = ss.getSheetByName('Events');
  const devSheet = ss.getSheetByName('Devotees');
  const notSheet = ss.getSheetByName('Notices');
  const cfgSheet = ss.getSheetByName('Config');
  const pendSheet = ss.getSheetByName('Pending');
  const analyticsSheet = ss.getSheetByName('Analytics');
  let result = {};

  function findRow(sheet, id) {
    const vals = sheet.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) if (String(vals[i][0]) === String(id)) return i + 1;
    return -1;
  }

  if (action === 'addEvent') {
    const ev = data.event;
    evSheet.appendRow([ev.id, ev.nameML, ev.nameEN, ev.date, ev.time, ev.type]);
    result = {ok: true};

  } else if (action === 'updateEvent') {
    const ev = data.event;
    const row = findRow(evSheet, ev.id);
    if (row > 0) evSheet.getRange(row,1,1,6).setValues([[ev.id,ev.nameML,ev.nameEN,ev.date,ev.time,ev.type]]);
    result = {ok: true};

  } else if (action === 'deleteEvent') {
    const row = findRow(evSheet, data.id);
    if (row > 0) evSheet.deleteRow(row);
    result = {ok: true};

  } else if (action === 'addDevotee') {
    const d = data.devotee;
    devSheet.appendRow([d.id,d.name,d.phone,d.place,d.star,d.pin,d.note,d.bday,'approved']);
    result = {ok: true};

  } else if (action === 'updateDevotee') {
    const d = data.devotee;
    const row = findRow(devSheet, d.id);
    if (row > 0) devSheet.getRange(row,1,1,9).setValues([[d.id,d.name,d.phone,d.place,d.star,d.pin,d.note,d.bday,'approved']]);
    result = {ok: true};

  } else if (action === 'deleteDevotee') {
    const row = findRow(devSheet, data.id);
    if (row > 0) devSheet.deleteRow(row);
    result = {ok: true};

  } else if (action === 'addPending') {
    const d = data.devotee;
    pendSheet.appendRow([d.id,d.name,d.phone,d.place,d.star,d.pin,d.note,d.bday,'pending']);
    result = {ok: true};

  } else if (action === 'approveDevotee') {
    const d = data.devotee;
    // Remove from pending
    const pendRow = findRow(pendSheet, d.id);
    if (pendRow > 0) pendSheet.deleteRow(pendRow);
    // Add to devotees
    devSheet.appendRow([d.id,d.name,d.phone,d.place,d.star,d.pin,d.note,d.bday,'approved']);
    result = {ok: true};

  } else if (action === 'rejectPending') {
    const row = findRow(pendSheet, data.id);
    if (row > 0) pendSheet.deleteRow(row);
    result = {ok: true};

  } else if (action === 'addNotice') {
    const n = data.notice;
    notSheet.appendRow([n.id, n.ts, n.type, n.msg, n.msgML||'']);
    result = {ok: true};

  } else if (action === 'deleteNotice') {
    const row = findRow(notSheet, data.id);
    if (row > 0) notSheet.deleteRow(row);
    result = {ok: true};

  } else if (action === 'setAdminPin') {
    cfgSheet.getRange('B1').setValue(data.pin);
    result = {ok: true};

  } else if (action === 'logVisit') {
    const today = new Date().toISOString().split('T')[0];
    analyticsSheet.appendRow([data.deviceId, today, data.ts, data.ua||'']);
    result = {ok: true};

  } else if (action === 'getAnalytics') {
    const rows = analyticsSheet.getDataRange().getValues();
    const visits = rows.slice(1).map(r => ({deviceId:String(r[0]),date:String(r[1]),ts:Number(r[2])}));
    result = {visits};
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evSheet = ss.getSheetByName('Events');
  const devSheet = ss.getSheetByName('Devotees');
  const notSheet = ss.getSheetByName('Notices');
  const cfgSheet = ss.getSheetByName('Config');
  const pendSheet = ss.getSheetByName('Pending');

  function fmtDate(val) {
    if (!val) return '';
    if (val instanceof Date) {
      return val.getFullYear()+'-'+String(val.getMonth()+1).padStart(2,'0')+'-'+String(val.getDate()).padStart(2,'0');
    }
    return String(val);
  }
  function fmtTime(val) {
    if (!val) return '';
    if (val instanceof Date) return String(val.getHours()).padStart(2,'0')+':'+String(val.getMinutes()).padStart(2,'0');
    return String(val);
  }

  function getEvents() {
    const rows = evSheet.getDataRange().getValues();
    if (rows.length < 2) return [];
    return rows.slice(1).filter(r=>r[0]).map(r=>({
      id:String(r[0]), nameML:String(r[1]||''), nameEN:String(r[2]||''),
      date:fmtDate(r[3]), time:fmtTime(r[4]), type:String(r[5]||'')
    }));
  }

  function getDevotees() {
    const rows = devSheet.getDataRange().getValues();
    if (rows.length < 2) return [];
    return rows.slice(1).filter(r=>r[0]).map(r=>({
      id:String(r[0]), name:String(r[1]||''), phone:String(r[2]||''),
      place:String(r[3]||''), star:String(r[4]||''), pin:String(r[5]||''),
      note:String(r[6]||''), bday:fmtDate(r[7]), status:String(r[8]||'approved')
    }));
  }

  function getPending() {
    const rows = pendSheet.getDataRange().getValues();
    if (rows.length < 2) return [];
    return rows.slice(1).filter(r=>r[0]).map(r=>({
      id:String(r[0]), name:String(r[1]||''), phone:String(r[2]||''),
      place:String(r[3]||''), star:String(r[4]||''), pin:String(r[5]||''),
      note:String(r[6]||''), bday:fmtDate(r[7]), status:'pending'
    }));
  }

  function getNotices() {
    const rows = notSheet.getDataRange().getValues();
    if (rows.length < 2) return [];
    return rows.slice(1).filter(r=>r[0]).map(r=>({
      id:String(r[0]), ts:Number(r[1])||0, type:String(r[2]||''),
      msg:String(r[3]||''), msgML:String(r[4]||'')
    })).reverse();
  }

  return ContentService.createTextOutput(JSON.stringify({
    events: getEvents(),
    devotees: getDevotees(),
    pending: getPending(),
    notices: getNotices(),
    adminPin: String(cfgSheet.getRange('B1').getValue())
  })).setMimeType(ContentService.MimeType.JSON);
}

function getAnalyticsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Analytics');
  const rows = sheet.getDataRange().getValues();
  const visits = rows.slice(1).filter(r=>r[0]).map(r => ({
    deviceId: String(r[0]),
    date: r[1] instanceof Date ?
      r[1].getFullYear()+'-'+String(r[1].getMonth()+1).padStart(2,'0')+'-'+String(r[1].getDate()).padStart(2,'0') :
      String(r[1]),
    ts: Number(r[2]) || 0
  }));
  return ContentService.createTextOutput(JSON.stringify({visits}))
    .setMimeType(ContentService.MimeType.JSON);
}
