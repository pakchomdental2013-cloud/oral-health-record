const SPREADSHEET_ID = '1LLTSUTMrGC0LB79ZsTxjJDrX5kkmu7Oo7-XMPtrdhFg';
const PROVIDERS = ['ทพ.ธรภัทร เจตนพิพัฒนพงษ์','ทพ.อุกฤษฏ์ ตนะทิพย์','ทพ.อิษยา นาคประสิทธิ์','จพ.อัญชลีพร สุริยวรรณ','จพ.สิรินันท์ ดีจริง','จพ.จุไรวรรณ แถวบุญมี','นวก.กริณากร พิศพร'];
const TOOTH_STATUS_VERSION_MARKER='[TOOTH_STATUS_VERSION_2]';
const SCHEMAS = {
  Students: ['เลขบัตรประชาชน','ชื่อ','นามสกุล','ชั้น','โรงเรียน','วันที่ลงทะเบียน'],
  Examinations: ['รหัสตรวจ','เลขบัตรประชาชน','วันที่ตรวจ','ผู้ตรวจ','โรงเรียน','หมายเหตุ','วันเวลาบันทึก','เหงือกอักเสบ','ความเร่งด่วน'],
  ToothRecords: ['รหัสผลตรวจ','รหัสตรวจ','เลขบัตรประชาชน','ซี่ฟัน FDI','ประเภทฟัน','สภาพฟันที่ตรวจพบ','แผนที่แนะนำ/ยืนยัน','สภาพวัสดุเคลือบหลุมร่องฟัน','หมายเหตุ'],
  TreatmentHistory: ['รหัสรักษา','รหัสตรวจอ้างอิง','เลขบัตรประชาชน','ซี่ฟัน FDI','วันที่รักษา','ผู้ให้บริการ','หัตถการ','ผลหลังรักษา','หมายเหตุ','วันเวลาบันทึก'],
  Providers: ['รายชื่อผู้ให้บริการ']
};

function doGet(){
  setupDatabase();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบตรวจสุขภาพช่องปากนักเรียน')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function setupDatabase(){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.keys(SCHEMAS).forEach(name=>{
    let sh=ss.getSheetByName(name);
    if(!sh) sh=ss.insertSheet(name);
    if(sh.getLastRow()===0) sh.getRange(1,1,1,SCHEMAS[name].length).setValues([SCHEMAS[name]]);
    else {
      const existingHeaders=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0].map(String);
      SCHEMAS[name].forEach((header,i)=>{ if(!existingHeaders[i]) sh.getRange(1,i+1).setValue(header); });
    }
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,SCHEMAS[name].length).setFontWeight('bold').setBackground('#dceee8');
  });
  const p=ss.getSheetByName('Providers');
  if(p.getLastRow()<2) p.getRange(2,1,PROVIDERS.length,1).setValues(PROVIDERS.map(x=>[x]));
  ss.getSheetByName('Students').getRange('A:A').setNumberFormat('@');
  return 'ตั้งค่าฐานข้อมูลเรียบร้อยแล้ว';
}

function sh_(n){ return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(n); }
function data_(n){
  const s=sh_(n);
  return s.getLastRow()<2 ? [] : s.getRange(2,1,s.getLastRow()-1,SCHEMAS[n].length).getValues();
}
function iso_(d){ return d ? Utilities.formatDate(new Date(d),Session.getScriptTimeZone(),'yyyy-MM-dd') : ''; }
function providers(){ return data_('Providers').map(r=>r[0]).filter(String); }

function normalizeCitizenId_(value){
  const thaiDigits={'๐':'0','๑':'1','๒':'2','๓':'3','๔':'4','๕':'5','๖':'6','๗':'7','๘':'8','๙':'9'};
  const raw=String(value==null?'':value).replace(/[๐-๙]/g,ch=>thaiDigits[ch]).trim().toUpperCase();
  // รหัสนักเรียนต่างชาติ เช่น G123456789 รองรับตัวอักษร G นำหน้าและตัวเลข/อักษรต่อท้าย
  const compact=raw.replace(/[\s-]/g,'');
  if(/^G[A-Z0-9]{1,19}$/.test(compact)) return compact;
  return compact.replace(/\D/g,'');
}
function isValidStudentId_(value){
  const id=normalizeCitizenId_(value);
  return /^\d{13}$/.test(id) || /^G[A-Z0-9]{1,19}$/.test(id);
}

function saveStudent(d){
  const id=normalizeCitizenId_(d.citizenId);
  if(!isValidStudentId_(id)) throw Error('กรุณากรอกเลขบัตรประชาชนไทย 13 หลัก หรือรหัสนักเรียนต่างชาติที่ขึ้นต้นด้วย G เช่น G123456789');
  const first=String(d.firstName||'').trim(), last=String(d.lastName||'').trim();
  const grade=String(d.grade||'').trim(), school=String(d.school||'').trim();
  if(!first||!last||!grade||!school) throw Error('กรุณากรอกชื่อ นามสกุล ชั้น และโรงเรียน');
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const s=sh_('Students'), rows=data_('Students');
    const i=rows.findIndex(r=>normalizeCitizenId_(r[0])===id);
    const row=[id,first,last,grade,school,i<0?new Date():rows[i][5]];
    if(i<0) s.appendRow(row);
    else s.getRange(i+2,1,1,row.length).setValues([row]);
    s.getRange('A:A').setNumberFormat('@');
    return i<0 ? 'เพิ่มทะเบียนนักเรียนแล้ว' : 'ปรับปรุงข้อมูลนักเรียนแล้ว';
  } finally { lock.releaseLock(); }
}

// แยกข้อมูลนักเรียนจากรูปแบบที่ผู้ใช้อาจคัดลอกมาได้หลายแบบ
function parseStudentImportLine_(line){
  const clean=String(line||'').replace(/\u00a0/g,' ').replace(/\u2007/g,' ').trim();
  if(!clean) return null;
  // แถวเส้นประจาก Markdown ไม่ใช่ข้อมูลนักเรียน
  if(/^\|?\s*:?-{3,}/.test(clean)) return null;

  let cells;
  if(clean.indexOf('\t')>=0){
    cells=clean.split('\t').map(x=>String(x||'').trim());
  } else if(clean.indexOf('|')>=0){
    cells=clean.split('|').map(x=>String(x||'').trim()).filter(Boolean);
    cells=cells.filter(x=>!/^:?-{3,}:?$/.test(x));
  } else {
    cells=[clean];
  }

  // รูปแบบมาตรฐาน: เลขบัตร | ชื่อ | นามสกุล | ชั้น | โรงเรียน
  if(cells.length>=5){
    return {idRaw:cells[0],first:cells[1],lastName:cells[2],grade:cells[3],school:cells.slice(4).join(' ').trim()};
  }

  // รูปแบบสี่คอลัมน์: เลขบัตร | ชื่อ-นามสกุล | ชั้น | โรงเรียน
  if(cells.length===4){
    let idRaw='',fullName='',grade='',school='';
    if(/^(?:G[A-Za-z0-9๐-๙]{1,19}|[0-9๐-๙]{13})$/i.test(cells[0].replace(/[\s-]/g,''))){
      [idRaw,fullName,grade,school]=cells;
    } else if(/^(?:G[A-Za-z0-9๐-๙]{1,19}|[0-9๐-๙]{13})$/i.test(cells[2].replace(/[\s-]/g,''))){
      [fullName,grade,idRaw,school]=cells;
    }
    if(idRaw){
      const parts=splitStudentName_(fullName);
      return {idRaw,first:parts.first,lastName:parts.lastName,grade,school};
    }
  }

  // รูปแบบข้อความคั่นด้วยช่องว่าง: รองรับทั้ง ID อยู่หน้าชื่อหรืออยู่หลังชั้น
  const idMatch=clean.match(/G[A-Za-z0-9๐-๙]{1,19}|[0-9๐-๙]{13}/i);
  const gradeMatch=clean.match(/อนุบาล\s*[1-3]|ป\.?\s*[1-6](?:\s*\/\s*\d+)?/);
  if(idMatch && gradeMatch){
    const withoutId=clean.replace(idMatch[0],' ').replace(/\s+/g,' ').trim();
    const gm=withoutId.match(/อนุบาล\s*[1-3]|ป\.?\s*[1-6](?:\s*\/\s*\d+)?/);
    if(gm){
      const before=withoutId.slice(0,gm.index).trim();
      const after=withoutId.slice(gm.index+gm[0].length).trim();
      const parts=splitStudentName_(before);
      return {idRaw:idMatch[0],first:parts.first,lastName:parts.lastName,grade:gm[0].replace(/\s+/g,' ').trim(),school:after};
    }
  }
  return {idRaw:'',first:'',lastName:'',grade:'',school:''};
}

function splitStudentName_(fullName){
  const text=String(fullName||'').replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim();
  const at=text.indexOf(' ');
  if(at<0) return {first:text,lastName:''};
  return {first:text.slice(0,at).trim(),lastName:text.slice(at+1).trim()};
}

// รับข้อมูลจาก Excel/ข้อความได้หลายพันแถว รองรับแท็บ ช่องว่าง และตาราง Markdown
function importStudentsTsv(txt){
  const raw=String(txt||'').replace(/^\uFEFF/,'').trim();
  if(!raw) throw Error('ยังไม่มีข้อมูล กรุณาคัดลอกรายชื่อนักเรียนแล้ววางในช่องนำเข้า');
  const lines=raw.split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
  if(!lines.length) throw Error('ไม่พบแถวข้อมูล');

  const parsed=[];
  lines.forEach((line,index)=>{
    if(/^\|?\s*:?-{3,}/.test(line)) return;
    const cells=line.indexOf('\t')>=0 ? line.split('\t').map(x=>x.trim()) :
      (line.indexOf('|')>=0 ? line.split('|').map(x=>x.trim()).filter(Boolean) : [line]);
    if(cells.some(cell=>/เลขบัตรประชาชน|เลขบัตร|บัตรประชาชน|นามสกุล|โรงเรียน/.test(cell))) return;
    const row=parseStudentImportLine_(line);
    if(row) { row.sourceLine=index+1; parsed.push(row); }
  });
  if(!parsed.length) throw Error('อ่านข้อมูลไม่สำเร็จ กรุณาวางข้อมูลรายชื่อในรูปแบบ เลขบัตร ชื่อ-นามสกุล ชั้น โรงเรียน');

  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const sheet=sh_('Students');
    const last=sheet.getLastRow();
    const existing=last>1 ? sheet.getRange(2,1,last-1,SCHEMAS.Students.length).getValues() : [];
    const indexById=new Map();
    existing.forEach((r,i)=>{
      const id=normalizeCitizenId_(r[0]);
      if(id) indexById.set(id,i);
    });

    let added=0, updated=0, skipped=0, invalidId=0, incomplete=0, duplicate=0;
    const seenInput=new Set();
    const newRows=[];
    parsed.forEach(r=>{
      const id=normalizeCitizenId_(r.idRaw);
      const first=String(r.first||'').trim(), lastName=String(r.lastName||'').trim();
      const grade=String(r.grade||'').trim(), school=String(r.school||'').trim();
      if(!isValidStudentId_(id)){skipped++;invalidId++;return;}
      if(!first||!lastName||!grade||!school){skipped++;incomplete++;return;}
      if(seenInput.has(id)){skipped++;duplicate++;return;}
      seenInput.add(id);
      const idx=indexById.get(id);
      if(idx===undefined){
        newRows.push([id,first,lastName,grade,school,new Date()]);
        added++;
      } else {
        existing[idx]=[id,first,lastName,grade,school,existing[idx][5]||new Date()];
        updated++;
      }
    });

    if(updated && existing.length) sheet.getRange(2,1,existing.length,SCHEMAS.Students.length).setValues(existing);
    if(newRows.length){
      const start=sheet.getLastRow()+1;
      sheet.getRange(start,1,newRows.length,SCHEMAS.Students.length).setValues(newRows);
    }
    sheet.getRange('A:A').setNumberFormat('@');
    return `นำเข้าสำเร็จ ${added} คน | ปรับปรุงข้อมูลเดิม ${updated} คน | ข้าม ${skipped} แถว (เลขบัตรไม่ถูกต้อง ${invalidId} แถว, ข้อมูลไม่ครบ ${incomplete} แถว, เลขซ้ำในไฟล์ ${duplicate} แถว)`;
  } finally { lock.releaseLock(); }
}

function searchStudents(q){
  q=String(q||'').toLowerCase().trim();
  return data_('Students').filter(r=>!q||r.slice(0,5).join(' ').toLowerCase().includes(q))
    .map(r=>({id:normalizeCitizenId_(r[0]),citizenId:normalizeCitizenId_(r[0]),firstName:r[1],lastName:r[2],grade:r[3],school:r[4]})).slice(0,100);
}
function getStudent(id){
  const key=normalizeCitizenId_(id);
  const r=data_('Students').find(x=>normalizeCitizenId_(x[0])===key);
  return r?{id:key,citizenId:key,firstName:r[1],lastName:r[2],grade:r[3],school:r[4]}:null;
}

// รายชื่อนักเรียนตามโรงเรียนและชั้น เพื่อให้ตรวจได้โดยไม่ต้องค้นหาเลขบัตรทีละคน
function getStudentRoster(school,grade){
  school=String(school||'').trim(); grade=String(grade||'').trim();
  const rows=data_('Students').filter(r=>(!school||String(r[4])===school)&&(!grade||String(r[3])===grade));
  const latest={};
  data_('Examinations').forEach(r=>{const id=normalizeCitizenId_(r[1]);const date=iso_(r[2]);if(!id)return;if(!latest[id]||date>=latest[id].date)latest[id]={date,gumStatus:r[7]||'ยังไม่ได้ประเมิน',urgency:r[8]||'ยังไม่ได้ประเมิน'};});
  return {schools:[...new Set(data_('Students').map(r=>String(r[4]||'')).filter(Boolean))].sort(),grades:[...new Set(data_('Students').filter(r=>!school||String(r[4])===school).map(r=>String(r[3]||'')).filter(Boolean))].sort(),students:rows.map(r=>{const id=normalizeCitizenId_(r[0]);return {id,citizenId:id,firstName:r[1],lastName:r[2],grade:r[3],school:r[4],lastExam:latest[id]||null};}).sort((a,b)=>(a.firstName+' '+a.lastName).localeCompare(b.firstName+' '+b.lastName,'th'))};
}

function saveExamination(d){
  const s=getStudent(d.studentId);
  if(!s) throw Error('กรุณาเลือกนักเรียน');
  if(!d.examDate) throw Error('กรุณาระบุวันที่ตรวจ');
  const records=Array.isArray(d.teeth)?d.teeth:[];
  const validRecords=records.filter(t=>t&&t.tooth&&t.status);
  const toothIds=validRecords.map(t=>String(t.tooth));
  const allowedStatuses=['ปกติ','ฟันผุ','ฟันที่มีหลุมร่องลึก','ฟันที่อุดแล้วผุต่อ','ฟันที่อุดแล้วไม่ผุ','ฟันที่ถูกถอน','ฟันน้ำนมหลุดตามธรรมชาติ','ฟันที่ผุระยะเริ่มแรก','ฟันที่ได้รับการเคลือบหลุมร่องฟันแล้ว','ฟันที่ยังไม่ขึ้น','ไม่มีฟันแต่กำเนิด'];
  if(validRecords.length!==52 || new Set(toothIds).size!==52 || !toothIds.every(validToothFDI_)){
    throw Error('ต้องระบุสถานะฟันครบทั้ง 52 ตำแหน่งและไม่ซ้ำกัน กรุณาตรวจสอบผังฟัน');
  }
  if(validRecords.some(t=>!allowedStatuses.includes(String(t.status)))){
    throw Error('พบสถานะฟันที่ไม่ถูกต้อง กรุณาเลือกสถานะจากรายการ');
  }
  if(validRecords.some(t=>t.status==='ฟันน้ำนมหลุดตามธรรมชาติ' && !isPrimaryTooth_(t.tooth))){
    throw Error('สถานะฟันน้ำนมหลุดตามธรรมชาติใช้ได้เฉพาะซี่ฟันน้ำนมเท่านั้น');
  }
  const examId=Utilities.getUuid();
  const noteText=[String(d.note||'').trim(),TOOTH_STATUS_VERSION_MARKER].filter(Boolean).join('\n');
  sh_('Examinations').appendRow([examId,s.citizenId,d.examDate,d.examiner||'',s.school,noteText,new Date(),d.gumStatus||'ยังไม่ได้ประเมิน',d.urgency||'ยังไม่ได้ประเมิน']);
  const values=validRecords.map(t=>[
    Utilities.getUuid(),examId,s.citizenId,String(t.tooth),t.toothType||'',t.status,t.plan||'',t.sealant||'',t.note||''
  ]);
  if(values.length) sh_('ToothRecords').getRange(sh_('ToothRecords').getLastRow()+1,1,values.length,SCHEMAS.ToothRecords.length).setValues(values);
  return {message:`บันทึกผลตรวจครบ ${values.length} ซี่`,examId};
}

function treatmentOutcome_(procedure){
  const p=String(procedure||'');
  if(p==='อุดฟัน') return 'ฟันที่อุดแล้วไม่ผุ';
  if(p==='PRR') return 'ได้รับการทำ PRR แล้ว';
  if(p==='ถอนฟัน') return 'ฟันที่ถูกถอน';
  if(p==='เคลือบหลุมร่องฟัน') return 'ฟันที่ได้รับการเคลือบหลุมร่องฟันแล้ว';
  if(p==='เคลือบฟลูออไรด์') return 'ได้รับการเคลือบฟลูออไรด์แล้ว';
  if(p==='รักษารากฟัน') return 'รักษารากฟันแล้ว';
  if(p==='ขูดหินปูน') return 'ได้รับการขูดหินปูนแล้ว';
  if(p==='ตรวจติดตาม') return 'ตรวจติดตามแล้ว';
  return 'บันทึกการรักษาแล้ว';
}
function validToothFDI_(tooth){
  const n=Number(tooth);
  return [11,12,13,14,15,16,17,18,21,22,23,24,25,26,27,28,31,32,33,34,35,36,37,38,41,42,43,44,45,46,47,48,51,52,53,54,55,61,62,63,64,65,71,72,73,74,75,81,82,83,84,85].includes(n);
}
function isPrimaryTooth_(tooth){
  const quadrant=Math.floor(Number(tooth)/10);
  return [5,6,7,8].includes(quadrant);
}
function saveTreatment(d){
  const student=getStudent(d.studentId);
  if(!student) throw Error('กรุณาค้นหาและเลือกนักเรียนก่อน');
  const tooth=String(d.tooth||'').trim();
  if(!validToothFDI_(tooth)) throw Error('กรุณาระบุหมายเลขซี่ฟัน FDI ให้ถูกต้อง เช่น 16 หรือ 54');
  if(!d.date) throw Error('กรุณาระบุวันที่รักษา');
  if(!d.provider) throw Error('กรุณาเลือกผู้ให้บริการ');
  if(!d.procedure) throw Error('กรุณาเลือกหัตถการที่ทำจริง');

  const history=getHistory(student.citizenId);
  const toothCurrent=(history.currentTeeth||[]).find(t=>String(t.tooth)===tooth);
  if(!toothCurrent || toothCurrent.status==='ยังไม่ได้ประเมิน'){
    throw Error('ยังไม่มีผลตรวจที่ยืนยันสถานะของซี่ '+tooth+' กรุณาบันทึกผลตรวจก่อนทำรายการรักษา');
  }
  if(toothCurrent && ['ฟันที่ถูกถอน','ฟันน้ำนมหลุดตามธรรมชาติ','ฟันที่ยังไม่ขึ้น','ไม่มีฟันแต่กำเนิด'].includes(toothCurrent.status)){
    throw Error('ซี่ '+tooth+' ถูกระบุว่า '+toothCurrent.status+' จึงไม่ควรบันทึกการรักษาเป็นซี่ฟันที่มีอยู่ กรุณาตรวจสอบสถานะผลตรวจก่อน');
  }
  const examDateById={};
  history.exams.forEach(e=>{examDateById[e.id]=e.date;});
  const matching=history.teeth.filter(t=>String(t.tooth)===tooth);
  // ผูกการรักษากับผลตรวจล่าสุดที่เกิดขึ้นก่อนหรือในวันที่รักษา
  const eligible=matching.filter(t=>!examDateById[t.examId] || examDateById[t.examId]<=d.date);
  eligible.sort((a,b)=>String(examDateById[b.examId]||'').localeCompare(String(examDateById[a.examId]||'')));
  matching.sort((a,b)=>String(examDateById[b.examId]||'').localeCompare(String(examDateById[a.examId]||'')));
  const examId=(eligible[0]||matching[0]||{}).examId||'';
  const outcome=String(d.outcome||'').trim()||treatmentOutcome_(d.procedure);
  sh_('TreatmentHistory').appendRow([Utilities.getUuid(),examId,student.citizenId,tooth,d.date,d.provider,d.procedure,outcome,d.note||'',new Date()]);
  return `บันทึกการรักษาซี่ ${tooth} เรียบร้อยแล้ว`;
}

function getHistory(studentId){
  const id=normalizeCitizenId_(studentId);
  const exams=data_('Examinations').filter(r=>normalizeCitizenId_(r[1])===id)
    .map(r=>{
      const rawNote=String(r[5]||'');
      return {id:r[0],date:iso_(r[2]),examiner:r[3],school:r[4],note:rawNote.replace(TOOTH_STATUS_VERSION_MARKER,'').trim(),gumStatus:r[7]||'ยังไม่ได้ประเมิน',urgency:r[8]||'ยังไม่ได้ประเมิน',toothStatusVerified:rawNote.includes(TOOTH_STATUS_VERSION_MARKER)};
    });
  const teeth=data_('ToothRecords').filter(r=>normalizeCitizenId_(r[2])===id)
    .map(r=>({id:r[0],examId:r[1],tooth:String(r[3]),toothType:r[4],status:r[5],plan:r[6],sealant:r[7],note:r[8]}));
  const treatments=data_('TreatmentHistory').filter(r=>normalizeCitizenId_(r[2])===id)
    .map(r=>({id:r[0],examId:r[1],tooth:String(r[3]),date:iso_(r[4]),provider:r[5],procedure:r[6],outcome:r[7],note:r[8]}));

  // สถานะปัจจุบันคำนวณแยกจากข้อมูลตรวจเดิม โดยไม่แก้ไขแถวผลตรวจของโรงเรียน
  const examDateById={},verifiedExamIds=new Set();
  exams.forEach(e=>{examDateById[e.id]=e.date;if(e.toothStatusVerified)verifiedExamIds.add(String(e.id));});
  const current={};
  teeth.forEach(t=>{
    // ผลตรวจจากเวอร์ชันเดิมอาจตั้งต้นทุกซี่เป็นปกติโดยอัตโนมัติ จึงเก็บเป็นประวัติ แต่ไม่นำมาคำนวณจำนวนฟันปัจจุบัน
    if(!verifiedExamIds.has(String(t.examId)))return;
    const key=String(t.tooth), examDate=examDateById[t.examId]||'';
    if(!current[key] || examDate>=current[key].examDate){ current[key]={...t,examDate}; }
  });
  treatments.slice().sort((a,b)=>String(a.date).localeCompare(String(b.date))).forEach(t=>{
    const key=String(t.tooth), item=current[key];
    if(!item || (t.date && item.examDate && t.date<item.examDate)) return;
    const updatedStatus=t.outcome||treatmentOutcome_(t.procedure);
    if(['ฟันที่อุดแล้วไม่ผุ','ฟันที่ถูกถอน','ฟันที่ได้รับการเคลือบหลุมร่องฟันแล้ว'].includes(updatedStatus)) item.status=updatedStatus;
    else if(t.procedure==='PRR' || updatedStatus==='ได้รับการทำ PRR แล้ว') item.status='ฟันที่อุดแล้วไม่ผุ';
    else if(t.procedure==='รักษารากฟัน') item.status='รักษารากฟันแล้ว';
    item.lastTreatment=t.date;
    item.lastProcedure=t.procedure;
    item.lastOutcome=updatedStatus;
  });
  return {exams,teeth,treatments,currentTeeth:Object.values(current)};
}

function dmftCounts_(studentId){
  const h=getHistory(studentId), current=h.currentTeeth;
  const permanent=current.filter(t=>!isPrimaryTooth_(t.tooth));
  const primary=current.filter(t=>isPrimaryTooth_(t.tooth));
  const absentStatuses=['ฟันที่ถูกถอน','ฟันน้ำนมหลุดตามธรรมชาติ','ฟันที่ยังไม่ขึ้น','ไม่มีฟันแต่กำเนิด'];
  const presentStatuses=['ปกติ','ฟันผุ','ฟันที่มีหลุมร่องลึก','ฟันที่อุดแล้วผุต่อ','ฟันที่อุดแล้วไม่ผุ','ฟันที่ผุระยะเริ่มแรก','ฟันที่ได้รับการเคลือบหลุมร่องฟันแล้ว','รักษารากฟันแล้ว'];
  function count(list){
    const D=list.filter(t=>['ฟันผุ','ฟันที่อุดแล้วผุต่อ','ฟันที่ผุระยะเริ่มแรก'].includes(t.status)).length;
    // การหลุดตามธรรมชาติของฟันน้ำนมไม่นับเป็น M
    const M=list.filter(t=>t.status==='ฟันที่ถูกถอน').length;
    const F=list.filter(t=>t.status==='ฟันที่อุดแล้วไม่ผุ').length;
    return {D,M,F,total:D+M+F};
  }
  const p=count(permanent),d=count(primary);
  const all=permanent.concat(primary);
  const presence={
    permanentPresent:permanent.filter(t=>presentStatuses.includes(t.status)).length,
    primaryPresent:primary.filter(t=>presentStatuses.includes(t.status)).length,
    totalPresent:all.filter(t=>presentStatuses.includes(t.status)).length,
    extracted:all.filter(t=>t.status==='ฟันที่ถูกถอน').length,
    naturalShedding:primary.filter(t=>t.status==='ฟันน้ำนมหลุดตามธรรมชาติ').length,
    unerupted:all.filter(t=>t.status==='ฟันที่ยังไม่ขึ้น').length,
    congenitalAbsent:all.filter(t=>t.status==='ไม่มีฟันแต่กำเนิด').length,
    unassessed:Math.max(0,52-current.length)+all.filter(t=>t.status==='ยังไม่ได้ประเมิน').length
  };
  const untreated=all.filter(t=>['ฟันผุ','ฟันที่อุดแล้วผุต่อ','ฟันที่ผุระยะเริ่มแรก'].includes(t.status)).length;
  return {permanent:p,primary:d,untreated,treated:h.treatments.length,presence};
}

function getStudentSummary(studentId){
  const st=getStudent(studentId);
  if(!st) throw Error('ไม่พบข้อมูลนักเรียน');
  return {student:st,counts:dmftCounts_(studentId)};
}
function summaryRows_(){
  return data_('Students').map(r=>{
    const id=normalizeCitizenId_(r[0]),c=dmftCounts_(id);
    return [id,r[1],r[2],r[3],r[4],c.presence.permanentPresent,c.presence.primaryPresent,c.presence.totalPresent,c.presence.extracted,c.presence.naturalShedding,c.presence.unerupted,c.presence.unassessed,c.permanent.D,c.permanent.M,c.permanent.F,c.permanent.total,c.primary.D,c.primary.M,c.primary.F,c.primary.total,c.untreated,c.treated];
  });
}
function csv_(rows){
  return '\uFEFF'+rows.map(r=>r.map(v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"').join(',')).join('\r\n');
}
function exportSummaryCsv(){
  const rows=[['เลขบัตรประชาชน','ชื่อ','นามสกุล','ชั้น','โรงเรียน','ฟันแท้ที่มีอยู่','ฟันน้ำนมที่มีอยู่','ฟันที่มีอยู่รวม','ฟันที่ถูกถอน','ฟันน้ำนมหลุดตามธรรมชาติ','ฟันที่ยังไม่ขึ้น','ตำแหน่งยังไม่ได้ประเมิน','ฟันแท้ D','ฟันแท้ M','ฟันแท้ F','ฟันแท้ DMFT','ฟันน้ำนม d','ฟันน้ำนม m','ฟันน้ำนม f','ฟันน้ำนม dmft','ซี่ที่ยังต้องรักษา','รายการรักษาที่บันทึก'],...summaryRows_()];
  return csv_(rows);
}

function getSchoolTreatmentReport(filters){
  filters=filters||{};
  const schoolFilter=String(filters.school||'').trim(), gradeFilter=String(filters.grade||'').trim();
  const students=data_('Students').filter(r=>(!schoolFilter||String(r[4])===schoolFilter)&&(!gradeFilter||String(r[3])===gradeFilter));
  const exams=data_('Examinations').map(r=>({id:String(r[0]),studentId:normalizeCitizenId_(r[1]),date:iso_(r[2]),examiner:r[3],school:r[4],gumStatus:r[7]||'ยังไม่ได้ประเมิน',urgency:r[8]||'ยังไม่ได้ประเมิน',verified:String(r[5]||'').includes(TOOTH_STATUS_VERSION_MARKER)}));
  const toothRows=data_('ToothRecords').map(r=>({examId:String(r[1]),studentId:normalizeCitizenId_(r[2]),tooth:String(r[3]),toothType:String(r[4]||''),status:String(r[5]||''),plan:String(r[6]||'')}));
  const treatmentRows=data_('TreatmentHistory').map(r=>({examId:String(r[1]),studentId:normalizeCitizenId_(r[2]),tooth:String(r[3]),date:iso_(r[4]),provider:String(r[5]||''),procedure:String(r[6]||''),outcome:String(r[7]||'')}));
  const cariesStatuses=['ฟันผุ','ฟันที่อุดแล้วผุต่อ','ฟันที่ผุระยะเริ่มแรก'];
  const absentStatuses=['ฟันที่ถูกถอน','ฟันน้ำนมหลุดตามธรรมชาติ','ฟันที่ยังไม่ขึ้น','ไม่มีฟันแต่กำเนิด','ยังไม่ได้ประเมิน'];
  const definitiveProcedures=['อุดฟัน','อุดฟันซ้ำ','PRR','รักษารากฟัน','ถอนฟัน','เคลือบหลุมร่องฟัน'];
  const rows=students.map(r=>{
    const id=normalizeCitizenId_(r[0]);
    const candidates=exams.filter(e=>e.studentId===id&&e.verified).sort((a,b)=>String(b.date).localeCompare(String(a.date)));
    const latest=candidates[0]||null;
    if(!latest) return {citizenId:id,firstName:r[1],lastName:r[2],grade:r[3],school:r[4],examDate:'',totalPresent:0,permanentPresent:0,primaryPresent:0,cariesDetected:0,cariesTeeth:[],filledTeeth:0,extractedTeeth:0,cariesFree:0,permanentD:0,permanentM:0,permanentF:0,permanentDMFT:0,primaryD:0,primaryM:0,primaryF:0,primaryDMFT:0,treatedCaries:0,untreatedCaries:0,status:'ยังไม่มีผลตรวจ',gumStatus:'ยังไม่ได้ประเมิน',urgency:'ยังไม่ได้ประเมิน',details:[]};
    const latestTeeth=toothRows.filter(t=>t.studentId===id&&t.examId===latest.id);
    const present=latestTeeth.filter(t=>!absentStatuses.includes(t.status));
    const permanent=present.filter(t=>!isPrimaryTooth_(t.tooth)), primary=present.filter(t=>isPrimaryTooth_(t.tooth));
    const countDMF=list=>({D:list.filter(t=>cariesStatuses.includes(t.status)).length,M:list.filter(t=>t.status==='ฟันที่ถูกถอน').length,F:list.filter(t=>t.status==='ฟันที่อุดแล้วไม่ผุ').length});
    const pd=countDMF(latestTeeth.filter(t=>!isPrimaryTooth_(t.tooth))), bd=countDMF(latestTeeth.filter(t=>isPrimaryTooth_(t.tooth)));
    const examTeeth=latestTeeth.filter(t=>cariesStatuses.includes(t.status));
    const details=examTeeth.map(t=>{
      const matches=treatmentRows.filter(x=>x.studentId===id&&x.examId===latest.id&&x.tooth===t.tooth).sort((a,b)=>String(a.date).localeCompare(String(b.date)));
      const completed=matches.some(x=>definitiveProcedures.includes(x.procedure)||['ฟันที่อุดแล้วไม่ผุ','ฟันที่ถูกถอน','รักษารากฟันแล้ว','ฟันที่ได้รับการเคลือบหลุมร่องฟันแล้ว'].includes(x.outcome));
      const last=matches[matches.length-1]||null;
      return {tooth:t.tooth,originalStatus:t.status,plan:t.plan,treated:completed,treatment:last?last.procedure:'',treatmentDate:last?last.date:'',provider:last?last.provider:''};
    });
    const treated=details.filter(x=>x.treated).length, detected=details.length, remaining=detected-treated;
    const cariesFree=present.filter(t=>!cariesStatuses.includes(t.status)).length;
    return {citizenId:id,firstName:r[1],lastName:r[2],grade:r[3],school:r[4],examDate:latest.date,totalPresent:present.length,permanentPresent:permanent.length,primaryPresent:primary.length,cariesDetected:detected,cariesTeeth:examTeeth.map(t=>t.tooth).sort(),filledTeeth:latestTeeth.filter(t=>t.status==='ฟันที่อุดแล้วไม่ผุ').length,extractedTeeth:latestTeeth.filter(t=>t.status==='ฟันที่ถูกถอน').length,cariesFree,permanentD:pd.D,permanentM:pd.M,permanentF:pd.F,permanentDMFT:pd.D+pd.M+pd.F,primaryD:bd.D,primaryM:bd.M,primaryF:bd.F,primaryDMFT:bd.D+bd.M+bd.F,treatedCaries:treated,untreatedCaries:remaining,status:detected===0?'ไม่พบฟันผุ':remaining===0?'รักษาครบแล้ว':'ยังรักษาไม่ครบ',gumStatus:latest.gumStatus,urgency:latest.urgency,details};
  });
  return {schools:[...new Set(data_('Students').map(r=>String(r[4]||'')).filter(Boolean))].sort(),grades:[...new Set(data_('Students').filter(r=>!schoolFilter||String(r[4])===schoolFilter).map(r=>String(r[3]||'')).filter(Boolean))].sort(),rows,totals:{students:rows.length,examined:rows.filter(x=>x.examDate).length,cariesDetected:rows.reduce((a,x)=>a+x.cariesDetected,0),treatedCaries:rows.reduce((a,x)=>a+x.treatedCaries,0),untreatedCaries:rows.reduce((a,x)=>a+x.untreatedCaries,0),fullyTreated:rows.filter(x=>x.cariesDetected>0&&x.untreatedCaries===0).length,notComplete:rows.filter(x=>x.untreatedCaries>0).length}};
}
function exportSchoolTreatmentReportCsv(school,grade){
  const report=getSchoolTreatmentReport({school:school||'',grade:grade||''});
  const headers=['เลขประจำตัว','ชื่อ','นามสกุล','ชั้น','โรงเรียน','วันที่ตรวจล่าสุด','ฟันแท้ที่มีอยู่','ฟันน้ำนมที่มีอยู่','ฟันที่มีอยู่รวม','ฟันผุ (ซี่)','เลขซี่ฟันที่ผุ','ฟันอุดแล้ว','ฟันที่ถูกถอน','ฟันที่ปราศจากฟันผุ','เหงือกอักเสบ','ความเร่งด่วน','ฟันแท้ D','ฟันแท้ M','ฟันแท้ F','DMFT','ฟันน้ำนม d','ฟันน้ำนม m','ฟันน้ำนม f','dmft','รักษาแล้ว (ซี่)','ยังไม่รักษา (ซี่)','สถานะ'];
  const rows=[headers,...report.rows.map(x=>[x.citizenId,x.firstName,x.lastName,x.grade,x.school,x.examDate,x.permanentPresent,x.primaryPresent,x.totalPresent,x.cariesDetected,x.cariesTeeth.join(', '),x.filledTeeth,x.extractedTeeth,x.cariesFree,x.gumStatus,x.urgency,x.permanentD,x.permanentM,x.permanentF,x.permanentDMFT,x.primaryD,x.primaryM,x.primaryF,x.primaryDMFT,x.treatedCaries,x.untreatedCaries,x.status])];
  return csv_(rows);
}

function dashboard(){
  return {students:data_('Students').length,examinations:data_('Examinations').length,treatments:data_('TreatmentHistory').length};
}
function exportCsv(){
  const rows=[['เลขบัตรประชาชน','ชื่อ','นามสกุล','ชั้น','โรงเรียน'],...data_('Students').map(r=>[normalizeCitizenId_(r[0]),r[1],r[2],r[3],r[4]])];
  return csv_(rows);
}
