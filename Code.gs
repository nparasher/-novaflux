const SHEET_NAME = 'Requests';
const ADMIN_EMAIL = ''; // Optional: put your email here for notification emails.

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  const headers = ['Request ID','Created At','Name','Email','Business Type','Plan','Requirement','Status','Admin Note','Updated At'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  return 'Setup complete';
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'health';
  if (action === 'health') return json({ok:true, service:'novaflux-requests'});
  if (action === 'status') {
    const id = String(e.parameter.id || '').trim();
    if (!id) return json({ok:false,error:'Missing request ID'});
    return json(findRequest(id));
  }
  return json({ok:false,error:'Unknown action'});
}

function doPost(e) {
  try {
    const p = e.parameter || {};
    const action = String(p.action || 'create').trim();
    if (action === 'create') return createRequest(p);
    if (action === 'update') return updateRequest(p);
    return json({ok:false,error:'Unknown action'});
  } catch (err) {
    return json({ok:false,error:err.message || String(err)});
  }
}

function createRequest(p) {
  const name = clean(p.name);
  const email = clean(p.email);
  const business = clean(p.business);
  const plan = clean(p.plan);
  const message = clean(p.message);
  if (!name || !email || !business || !message) return json({ok:false,error:'Name, email, business type and requirement are required.'});
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return json({ok:false,error:'Please enter a valid email.'});

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { setup(); sheet = ss.getSheetByName(SHEET_NAME); }

  const id = 'NF-' + Utilities.getUuid().replace(/-/g,'').substring(0,8).toUpperCase();
  const now = new Date();
  sheet.appendRow([id, now, name, email, business, plan, message, 'Pending', '', now]);

  if (ADMIN_EMAIL) {
    try {
      MailApp.sendEmail(ADMIN_EMAIL, 'New Novaflux demo request: ' + id,
        'New request '+id+'\nName: '+name+'\nEmail: '+email+'\nBusiness: '+business+'\nPlan: '+plan+'\nRequirement: '+message+'\n\nOpen the Requests sheet to approve/reject.');
    } catch (_) {}
  }
  return json({ok:true, requestId:id, status:'Pending'});
}

function findRequest(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return {ok:false,error:'Request sheet not initialized'};
  const values = sheet.getDataRange().getValues();
  for (let i=1;i<values.length;i++) {
    if (String(values[i][0]).trim().toUpperCase() === id.toUpperCase()) {
      return {ok:true, requestId:values[i][0], createdAt:values[i][1], status:values[i][7] || 'Pending', adminNote:values[i][8] || ''};
    }
  }
  return {ok:false,error:'Request not found'};
}

function updateRequest(p) {
  const id = clean(p.requestId).toUpperCase();
  const status = clean(p.status);
  const note = clean(p.note);
  const allowed = ['Pending','Approved','Rejected','Contacted'];
  if (!id || allowed.indexOf(status) === -1) return json({ok:false,error:'Invalid request ID or status'});
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return json({ok:false,error:'Request sheet not initialized'});
  const ids = sheet.getRange(2,1,Math.max(sheet.getLastRow()-1,0),1).getValues();
  for (let i=0;i<ids.length;i++) {
    if (String(ids[i][0]).trim().toUpperCase() === id) {
      const row = i+2;
      sheet.getRange(row,8,1,3).setValues([[status,note,new Date()]]);
      return json({ok:true,requestId:id,status:status});
    }
  }
  return json({ok:false,error:'Request not found'});
}

function clean(v) { return String(v == null ? '' : v).trim().substring(0,5000); }
function json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
