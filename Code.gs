const CONFIG = {
  SPREADSHEET_ID: 'PASTE_GOOGLE_SHEET_ID_HERE',
  DRIVE_FOLDER_ID: 'PASTE_GOOGLE_DRIVE_FOLDER_ID_HERE',
  REGISTRATION_FEE_GHS: 300,
  TUITION_GHS: 1500,
  OTHER_EXPENSES_GHS: 200,
  // Set true only if the October-December cohort is officially tuition-waived.
  TUITION_WAIVED: false,
  CURRENCY: 'GHS',
  PROGRAM: 'Character Building'
};

const MODULES = [
  ['M01','Self Awareness and Identity Clarity',1],
  ['M02','Emotional Intelligence (EQ)',2],
  ['M03','Ethical Reasoning and Values',3],
  ['M04','Goal Setting and Direction',4],
  ['M05','Personal Productivity Systems',5],
  ['M06','Foundational Financial Literacy',6],
  ['M07','Digital Responsibility',7],
  ['M08','Resilience and Mental Stability',8],
  ['M09','Civic Engagement and National Responsibility',9],
  ['M10','Interpersonal Leadership and Conflict Resolution',10],
  ['M11','Radical Honesty and Accountability (Anti Corruption)',11],
  ['M12','Excellence, Order and Property Care',12]
];

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'verify') return verifyEmail_(e.parameter.token || '');
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Character Building LMS');
}

function setup() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const defs = {
    Students: ['id','createdAt','fullName','dob','profession','email','phone','term','username','passwordHash','emailVerified','status','paymentStatus','paymentReference','paymentAmount','paymentAt'],
    Payments: ['reference','studentId','email','amount','currency','status','provider','createdAt','verifiedAt','rawStatus'],
    Assignments: ['id','moduleCode','title','description','dueDate','fileUrl','createdAt','active'],
    Submissions: ['id','assignmentId','studentId','fileUrl','submittedAt','grade','feedback','status'],
    Announcements: ['id','title','body','createdAt','active'],
    Modules: ['code','title','week']
  };
  Object.keys(defs).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(defs[name]);
  });
  const mod = ss.getSheetByName('Modules');
  if (mod.getLastRow() <= 1) mod.getRange(2,1,MODULES.length,3).setValues(MODULES);
  return 'Setup complete';
}

function registerStudent(data) {
  validate_(data, ['fullName','dob','profession','email','phone','term']);
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName('Students');
  const email = String(data.email).trim().toLowerCase();
  if (findRow_(sh, 6, email)) throw new Error('An account with this email already exists.');
  const id = Utilities.getUuid();
  const username = uniqueUsername_(sh, data.fullName);
  const password = randomPassword_();
  const verifyToken = Utilities.getUuid();
  const row = [id,new Date(),data.fullName,data.dob,data.profession,email,data.phone,data.term,username,sha256_(password),false,'pending_payment','unpaid','',0,''];
  sh.appendRow(row);
  PropertiesService.getScriptProperties().setProperty('VERIFY_'+verifyToken, email);
  sendEmail_(email, 'Verify your Character Building email', `Hello ${data.fullName},\n\nPlease verify your email to continue your Character Building registration:\n${getAppUrl_()}?action=verify&token=${encodeURIComponent(verifyToken)}\n\nYour username will be ${username}. Your temporary password will be issued after successful payment.\n\nCharacter Building`);
  return {studentId:id, username, message:'Verification email sent.'};
}

function verifyEmail_(token) {
  const email = PropertiesService.getScriptProperties().getProperty('VERIFY_'+token);
  if (!email) return HtmlService.createHtmlOutput('<h2>Verification link expired or invalid.</h2>');
  const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Students');
  const row = findRow_(sh,6,email);
  if (!row) return HtmlService.createHtmlOutput('<h2>Account not found.</h2>');
  sh.getRange(row,11).setValue(true);
  PropertiesService.getScriptProperties().deleteProperty('VERIFY_'+token);
  return HtmlService.createHtmlOutput('<h2>Email verified.</h2><p>You can now return to the Character Building portal and make your registration payment.</p>');
}

function initiateMomoPayment(studentId) {
  const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Students');
  const row = findRow_(sh,1,studentId);
  if (!row) throw new Error('Student not found.');
  const v = sh.getRange(row,1,1,16).getValues()[0];
  if (!v[10]) throw new Error('Please verify your email first.');
  const amount = CONFIG.REGISTRATION_FEE_GHS;
  const secret = PropertiesService.getScriptProperties().getProperty('PAYSTACK_SECRET_KEY');
  if (!secret) throw new Error('Payment gateway is not configured.');
  const reference = 'CBREG-'+Utilities.getUuid().replace(/-/g,'').slice(0,24);
  const payload = {email:v[5], amount:Math.round(amount*100), currency:'GHS', mobile_money:{phone:normalizePhone_(v[6]), provider:'mtn'}, reference:reference, metadata:{studentId:studentId, purpose:'Character Building registration'}};
  const res = UrlFetchApp.fetch('https://api.paystack.co/charge', {method:'post', contentType:'application/json', headers:{Authorization:'Bearer '+secret}, payload:JSON.stringify(payload), muteHttpExceptions:true});
  const body = JSON.parse(res.getContentText());
  if (!body.status) throw new Error(body.message || 'Unable to start Mobile Money payment.');
  sh.getRange(row,14).setValue(reference);
  sh.getRange(row,13).setValue('pending');
  SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Payments').appendRow([reference,studentId,v[5],amount,'GHS',body.data?.status || 'pending','paystack',new Date(),'','']);
  return {reference:reference, status:body.data?.status || 'pending', message:body.data?.display_text || body.message || 'Approve the payment prompt on your MTN phone.'};
}

function login(username,password) {
  const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Students');
  const values = sh.getDataRange().getValues();
  for (let i=1;i<values.length;i++) {
    const r=values[i];
    if (String(r[8]).toLowerCase()===String(username).toLowerCase() && r[9]===sha256_(password)) {
      return {ok:true, studentId:r[0], name:r[2], username:r[8], term:r[7], paymentStatus:r[12], modules:MODULES};
    }
  }
  if (username==='bkoomson' && password==='BK123456') return {ok:true, role:'instructor', name:'Facilitator', username:'bkoomson'};
  throw new Error('Invalid username or password.');
}

function paystackWebhook(e) {
  const secret = PropertiesService.getScriptProperties().getProperty('PAYSTACK_SECRET_KEY');
  const signature = e && e.parameter && e.parameter['x-paystack-signature'];
  // Apps Script does not expose arbitrary headers in doPost. For production, use a small
  // Cloud Run/Node webhook or Paystack's verification endpoint from Apps Script before marking paid.
  // This function is intentionally not used to trust client-side success.
  return 'Webhook endpoint received; configure server-side signature verification.';
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const data = JSON.parse(raw);
    if (data && data.event === 'charge.success') {
      const ref = data.data && data.data.reference;
      if (ref) verifyPaystackTransaction_(ref);
    }
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function verifyPaystackTransaction_(reference) {
  const secret = PropertiesService.getScriptProperties().getProperty('PAYSTACK_SECRET_KEY');
  if (!secret) throw new Error('Missing Paystack secret');
  const res = UrlFetchApp.fetch('https://api.paystack.co/transaction/verify/'+encodeURIComponent(reference), {headers:{Authorization:'Bearer '+secret}, muteHttpExceptions:true});
  const body = JSON.parse(res.getContentText());
  if (!body.status || !body.data || body.data.status !== 'success') return false;
  const amountGhs = Number(body.data.amount)/100;
  const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Payments');
  const pRow = findRow_(sh,1,reference);
  if (!pRow) return false;
  const p = sh.getRange(pRow,1,1,10).getValues()[0];
  if (Number(p[3]) !== amountGhs || p[4] !== 'GHS') throw new Error('Payment amount/currency mismatch.');
  sh.getRange(pRow,6).setValue('success'); sh.getRange(pRow,9).setValue(new Date()); sh.getRange(pRow,10).setValue(body.data.status);
  const students = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Students');
  const sRow = findRow_(students,1,p[1]);
  if (!sRow) return false;
  const vals = students.getRange(sRow,1,1,16).getValues()[0];
  if (vals[12] === 'paid' && vals[13] === reference) return true;
  students.getRange(sRow,13).setValue('paid'); students.getRange(sRow,14).setValue(reference); students.getRange(sRow,15).setValue(amountGhs); students.getRange(sRow,16).setValue(new Date());
  const password = randomPassword_();
  students.getRange(sRow,10).setValue(sha256_(password));
  sendEmail_(vals[5], 'Character Building registration payment confirmed', `Hello ${vals[2]},\n\nYour GHS ${amountGhs.toFixed(2)} registration payment has been verified successfully.\n\nUsername: ${vals[8]}\nTemporary password: ${password}\n\nPlease sign in and change your password immediately.\n\nPayment reference: ${reference}\n\nCharacter Building`);
  return true;
}

function sendEmail_(to,subject,body) {
  MailApp.sendEmail({to:to,subject:subject,body:body,name:PropertiesService.getScriptProperties().getProperty('SENDER_NAME')||'Character Building'});
}
function getAppUrl_(){return PropertiesService.getScriptProperties().getProperty('APP_URL') || ScriptApp.getService().getUrl();}
function normalizePhone_(p){let x=String(p).replace(/\D/g,''); if(x.startsWith('0')) x='233'+x.slice(1); if(x.startsWith('+')) x=x.slice(1); return x;}
function randomPassword_(){return Utilities.getUuid().replace(/-/g,'').slice(0,10)+'!';}
function sha256_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8).map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join('');}
function validate_(o,keys){keys.forEach(k=>{if(!o[k]) throw new Error(k+' is required.');});}
function findRow_(sh,col,value){const vals=sh.getRange(1,col,Math.max(sh.getLastRow(),1),1).getValues(); for(let i=1;i<vals.length;i++) if(String(vals[i][0]).toLowerCase()===String(value).toLowerCase()) return i+1; return null;}
function uniqueUsername_(sh,name){const base=String(name).toLowerCase().replace(/[^a-z0-9]/g,'').slice(0,12)||'student'; let u=base, n=1; while(findRow_(sh,9,u)) u=base+(++n); return u;}
