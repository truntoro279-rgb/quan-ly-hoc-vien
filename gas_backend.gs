// Google Apps Script backend cho app Quản lý học viên.
// 1) Mở Google Sheet của bạn > Extensions > Apps Script.
// 2) Dán toàn bộ code này vào Code.gs.
// 3) Đổi SHEET_ID thành ID của Google Sheet.
// 4) Deploy > New deployment > Web app > Execute as Me > Anyone.
// 5) Copy URL /exec và dán vào API_URL trong index.html.

const SHEET_ID = 'DAN_ID_GOOGLE_SHEET_VAO_DAY';
const STUDENT_SHEET = 'ALL';
const LEAVE_SHEET = 'VE_QUE';
const ATT_SHEET = 'DIEM_DANH';

function sh(name){ return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name); }
function json(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e){
  try{
    const p=JSON.parse(e.postData.contents||'{}');
    if(p.action==='getAll') return json(getAll());
    if(p.action==='addLeave') { addLeave(p.row); return json({ok:true}); }
    if(p.action==='updateLeave') { updateLeave(p.id,p.status); return json({ok:true}); }
    if(p.action==='saveAttendance') { saveAttendance(p.rows||[]); return json({ok:true}); }
    return json({ok:false,error:'Unknown action'});
  }catch(err){ return json({ok:false,error:String(err)}); }
}

function rowsToObjects(sheet){
  if(!sheet) return [];
  const v=sheet.getDataRange().getValues(); if(v.length<2) return [];
  const h=v[0].map(String);
  return v.slice(1).filter(r=>r.some(x=>x!=='')).map(r=>Object.fromEntries(h.map((k,i)=>[k,String(r[i]??'')])));
}

function normalizeStudent(o){
  const get=(...ks)=>{for(const k of ks) if(o[k]!==undefined && o[k]!=='' ) return o[k]; return '';};
  return {
    code:get('Mã học viên','Mã HV','Mã HV ','Ma hoc vien'),
    name:get('Họ tên Romaji','Họ tên','Ho ten Romaji','Họ và tên'),
    katakana:get('Họ tên Katakana','Ho ten Katakana'),
    class:get('Lớp','Lop'),
    center:get('Trung tâm','Trung tam'),
    teacher:get('Giáo viên','Giao vien'),
    status:get('Trạng thái','Status'),
    birth:get('Ngày sinh','Ngay sinh')
  };
}

function getAll(){
  const students=rowsToObjects(sh(STUDENT_SHEET)).map(normalizeStudent).filter(x=>x.code||x.name);
  const leaves=rowsToObjects(sh(LEAVE_SHEET));
  const attendance=rowsToObjects(sh(ATT_SHEET)).filter(x=>x.date===Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'));
  return {students,leaves,attendance};
}

function ensureSheet(name,headers){
  let s=SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  if(!s){s=SpreadsheetApp.openById(SHEET_ID).insertSheet(name);s.appendRow(headers);}
  return s;
}
function addLeave(r){
  const s=ensureSheet(LEAVE_SHEET,['id','code','name','class','from','to','reason','phone','contact','status','createdAt']);
  s.appendRow([r.id,r.code,r.name,r.class,r.from,r.to,r.reason,r.phone,r.contact,r.status,new Date()]);
}
function updateLeave(id,status){
  const s=sh(LEAVE_SHEET); if(!s)return;
  const v=s.getDataRange().getValues();
  for(let i=1;i<v.length;i++) if(String(v[i][0])===String(id)){s.getRange(i+1,10).setValue(status);break;}
}
function saveAttendance(rows){
  const s=ensureSheet(ATT_SHEET,['code','name','status','date','savedAt']);
  const date=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const old=s.getDataRange().getValues();
  const codes=new Set(rows.map(r=>String(r.code)));
  for(let i=old.length-1;i>=1;i--) if(String(old[i][3])===date && codes.has(String(old[i][0]))) s.deleteRow(i+1);
  rows.forEach(r=>s.appendRow([r.code,r.name,r.status,r.date,new Date()]));
}
