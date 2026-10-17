/**
 * PJP SURAJ V11 - Database Resolver
 * IMPORTANT: This function is intentionally self-contained so the Web App
 * never depends on SpreadsheetApp.getActiveSpreadsheet().
 */
function getDataSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const key = 'PJP_DATA_SPREADSHEET_ID';
  let id = props.getProperty(key);

  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (err) {
      // Stored ID is invalid/deleted; create a fresh database below.
      props.deleteProperty(key);
    }
  }

  const ss = SpreadsheetApp.create('PJP_SURAJ_Data');
  props.setProperty(key, ss.getId());
  return ss;
}

/**
 * PJP VISIT CALENDAR - Google Apps Script
 * Mobile-first Web App for SURAJ
 *
 * Spreadsheet tabs:
 * 1) Visits   - Date, Day, Place, Remarks, Status, Holiday Reason
 * 2) Holidays - Date, Reason
 * 3) Settings - Name
 */

const APP_NAME = 'PJP Visit Calendar';
const DEFAULT_NAME = 'SURAJ';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupApp() {
  const ss = getDataSpreadsheet_();
  ensureCoreSheets_(ss);
  return 'Setup complete: ' + ss.getName();
}

function getName() {
  const cache=CacheService.getScriptCache();
  const cached=cache.get('PJP_NAME');
  if(cached) return cached;
  const ss=getDataSpreadsheet_();
  ensureCoreSheets_(ss);
  const sh=ss.getSheetByName('Settings');
  const vals=sh.getDataRange().getValues();
  for(let i=1;i<vals.length;i++){
    if(String(vals[i][0]).toLowerCase()==='name'){
      const name=String(vals[i][1]||DEFAULT_NAME);
      cache.put('PJP_NAME',name,21600);
      return name;
    }
  }
  cache.put('PJP_NAME',DEFAULT_NAME,21600);
  return DEFAULT_NAME;
}

function saveName(name) {
  name = String(name || DEFAULT_NAME).trim() || DEFAULT_NAME;
  const ss = getDataSpreadsheet_();
  const sh = ss.getSheetByName('Settings') || ensureSheet_(ss, 'Settings', ['Key','Value']);
  const values = sh.getDataRange().getValues();
  for (let i=1;i<values.length;i++) {
    if (String(values[i][0]).toLowerCase() === 'name') {
      sh.getRange(i+1,2).setValue(name);
      CacheService.getScriptCache().put('PJP_NAME',name,21600);
      return name;
    }
  }
  sh.appendRow(['Name', name]);
  CacheService.getScriptCache().put('PJP_NAME',name,21600);
  return name;
}

function getProfile() {
  const ss=getDataSpreadsheet_();
  ensureCoreSheets_(ss);
  const sh=ss.getSheetByName('Settings');
  const out={name:DEFAULT_NAME,employeeId:''};
  if(sh.getLastRow()>1){
    const vals=sh.getRange(2,1,sh.getLastRow()-1,2).getValues();
    vals.forEach(r=>{
      const key=String(r[0]||'').trim().toLowerCase();
      if(key==='name') out.name=String(r[1]||DEFAULT_NAME).trim()||DEFAULT_NAME;
      if(key==='employee id' || key==='employeeid' || key==='employee_id') out.employeeId=String(r[1]||'').trim();
    });
  }
  return out;
}

function saveProfile(payload) {
  const name=String(payload && payload.name || DEFAULT_NAME).trim() || DEFAULT_NAME;
  const employeeId=String(payload && payload.employeeId || '').trim();
  const ss=getDataSpreadsheet_();
  ensureCoreSheets_(ss);
  const sh=ss.getSheetByName('Settings');
  const vals=sh.getDataRange().getValues();
  const upsert=(key,value)=>{
    for(let i=1;i<vals.length;i++){
      if(String(vals[i][0]||'').trim().toLowerCase()===key.toLowerCase()){
        sh.getRange(i+1,2).setValue(value); return;
      }
    }
    sh.appendRow([key,value]);
  };
  upsert('Name',name);
  upsert('Employee ID',employeeId);
  CacheService.getScriptCache().put('PJP_NAME',name,21600);
  return {name,employeeId};
}

function getMonthData(year, month) {
  year = Number(year);
  month = Number(month);
  const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
  const daysInMonth = new Date(year, month, 0).getDate();

  const ss = getDataSpreadsheet_();
  ensureCoreSheets_(ss);

  const visitSh = ss.getSheetByName('Visits');
  const holSh = ss.getSheetByName('Holidays');
  const counterSh = ss.getSheetByName('Counters');
  const settingsSh = ss.getSheetByName('Settings');

  const visits = {};
  if (visitSh.getLastRow() > 1) {
    const vals = visitSh.getRange(2,1,visitSh.getLastRow()-1,6).getValues();
    for (let i=0;i<vals.length;i++) {
      const r=vals[i];
      if (!r[0]) continue;
      const d=new Date(r[0]);
      if (d.getFullYear()===year && d.getMonth()+1===month) {
        const key=Utilities.formatDate(d,tz,'yyyy-MM-dd');
        visits[key]={place:String(r[2]||''),remarks:String(r[3]||'')};
      }
    }
  }

  const holidays = {};
  if (holSh.getLastRow() > 1) {
    const vals=holSh.getRange(2,1,holSh.getLastRow()-1,2).getValues();
    for (let i=0;i<vals.length;i++) {
      const r=vals[i];
      if (!r[0]) continue;
      const d=new Date(r[0]);
      if (d.getFullYear()===year && d.getMonth()+1===month) {
        holidays[Utilities.formatDate(d,tz,'yyyy-MM-dd')]=String(r[1]||'Holiday');
      }
    }
  }

  const counters = {};
  if (counterSh.getLastRow() > 1) {
    const vals=counterSh.getRange(2,1,counterSh.getLastRow()-1,4).getValues();
    for (let i=0;i<vals.length;i++) {
      const r=vals[i];
      if (Number(r[0])===year && Number(r[1])===month) {
        const p=String(r[2]||'').trim();
        if(p) counters[p.toLowerCase()]=Number(r[3]||0);
      }
    }
  }

  let name=DEFAULT_NAME;
  let employeeId='';
  if (settingsSh.getLastRow()>1) {
    const vals=settingsSh.getRange(2,1,settingsSh.getLastRow()-1,2).getValues();
    for (let i=0;i<vals.length;i++) {
      const settingKey=String(vals[i][0]||'').toLowerCase();
      if(settingKey==='name'){
        name=String(vals[i][1]||DEFAULT_NAME);
      } else if(settingKey==='employee id' || settingKey==='employeeid' || settingKey==='employee_id'){
        employeeId=String(vals[i][1]||'');
      }
    }
  }

  const rows=[];
  const placeCounts={};
  let working=0, holidayDays=0, visitDays=0;

  for(let day=1;day<=daysInMonth;day++){
    const dt=new Date(year,month-1,day);
    const key=Utilities.formatDate(dt,tz,'yyyy-MM-dd');
    const weekday=Utilities.formatDate(dt,tz,'EEEE');
    const sunday=dt.getDay()===0;
    const manualHoliday=holidays[key];
    const isHoliday=sunday || !!manualHoliday;
    const reason=sunday?'Sunday':(manualHoliday||'');
    const v=visits[key]||{};
    const place=isHoliday?'':(v.place||'');
    rows.push({date:key,day:weekday,dayNumber:day,place,remarks:v.remarks||'',isHoliday,holidayReason:reason});

    if(isHoliday) holidayDays++;
    else {
      working++;
      if(place){
        const normalized=place.toLowerCase().replace(/\s+/g,' ').trim();
        let display=Object.keys(placeCounts).find(k=>k.toLowerCase()===normalized);
        if(!display) display=place;
        placeCounts[display]=(placeCounts[display]||0)+1;
        visitDays++;
      }
    }
  }

  const places=Object.keys(placeCounts)
    .map(place=>({place,days:placeCounts[place],counter:Number(counters[place.toLowerCase()]||0)}))
    .sort((a,b)=>b.days-a.days || a.place.localeCompare(b.place));

  const summary={
    name,employeeId,year,month,totalDays:daysInMonth,workingDays:working,
    holidayDays,visitDays,places
  };
  return {name,employeeId,year,month,days:rows,summary};
}



function saveVisitsBulk(payload) {
  const assignments = Array.isArray(payload.assignments) ? payload.assignments : [];
  const remarks = String(payload.remarks || '').trim();
  if (!assignments.length) throw new Error('Please select at least one place and date.');

  const ss=getDataSpreadsheet_();
  ensureCoreSheets_(ss);
  const visitSh=ss.getSheetByName('Visits');
  const holSh=ss.getSheetByName('Holidays');
  const tz=Session.getScriptTimeZone()||'Asia/Kolkata';

  const rowMap={};
  if(visitSh.getLastRow()>1){
    const vals=visitSh.getRange(2,1,visitSh.getLastRow()-1,1).getValues();
    for(let i=0;i<vals.length;i++) if(vals[i][0]){
      const d=new Date(vals[i][0]);
      rowMap[Utilities.formatDate(d,tz,'yyyy-MM-dd')]=i+2;
    }
  }
  const holidayMap={};
  if(holSh.getLastRow()>1){
    const vals=holSh.getRange(2,1,holSh.getLastRow()-1,2).getValues();
    for(let i=0;i<vals.length;i++) if(vals[i][0]){
      const d=new Date(vals[i][0]);
      holidayMap[Utilities.formatDate(d,tz,'yyyy-MM-dd')]=true;
    }
  }

  // Build one combined date map first. A date may be assigned to only one
  // place in a single Save All operation. The UI normally moves a date when
  // the user assigns it to another place, but the server validates it too.
  const combined={};
  assignments.forEach(a=>{
    const place=String(a.place||'').trim();
    if(!place) return;
    (Array.isArray(a.dates)?a.dates:[]).forEach(key=>{
      const k=String(key);
      if(combined[k] && combined[k].place!==place) throw new Error('Date '+k+' is assigned to more than one place in this batch. Please keep it with one place.');
      combined[k]={place};
    });
  });

  const appendRows=[];
  const occupied=[];
  Object.keys(combined).sort().forEach(key=>{
    const date=new Date(key+'T00:00:00');
    if(date.getDay()===0 || holidayMap[key]) return;
    if(rowMap[key]){ occupied.push(key); return; }
    const place=combined[key].place;
    appendRows.push([date,Utilities.formatDate(date,tz,'EEEE'),place,remarks,'Working','']);
  });
  if(occupied.length){
    throw new Error('These dates are already saved: '+occupied.join(', ')+'. Use Fill Visit to reassign an existing date.');
  }
  if(appendRows.length) visitSh.getRange(visitSh.getLastRow()+1,1,appendRows.length,6).setValues(appendRows);

  const firstKey=Object.keys(combined).sort()[0];
  const firstDate=new Date(firstKey+'T00:00:00');
  return getMonthData(firstDate.getFullYear(),firstDate.getMonth()+1);
}

function clearAllSavedVisits(year, month) {
  year = Number(year);
  month = Number(month);
  if (!year || !month || month < 1 || month > 12) throw new Error('Invalid month.');

  const ss = getDataSpreadsheet_();
  ensureCoreSheets_(ss);
  const sh = ss.getSheetByName('Visits');
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return {removed: 0, data: getMonthData(year, month)};

  const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
  const vals = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  const rowsToDelete = [];

  vals.forEach((r, i) => {
    const raw = r[0];
    if (!raw) return;
    let key = '';
    if (raw instanceof Date && !isNaN(raw.getTime())) {
      key = Utilities.formatDate(raw, tz, 'yyyy-MM-dd');
    } else {
      const text = String(raw).trim();
      if (/^\d{4}-\d{2}-\d{2}/.test(text)) {
        key = text.slice(0, 10);
      } else {
        const parsed = new Date(text);
        if (!isNaN(parsed.getTime())) key = Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
      }
    }
    if (key && key.slice(0,4) === String(year) && Number(key.slice(5,7)) === month) {
      rowsToDelete.push(i + 2);
    }
  });

  // Delete actual rows from bottom to top so no blank/stale visit rows remain.
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sh.deleteRow(rowsToDelete[i]);
  }

  SpreadsheetApp.flush();
  return {removed: rowsToDelete.length, data: getMonthData(year, month)};
}

function saveVisit(payload) {
  const date = new Date(payload.date + 'T00:00:00');
  const year=date.getFullYear(), month=date.getMonth()+1;
  const ss=getDataSpreadsheet_();
  ensureCoreSheets_(ss);
  const visitSh=ss.getSheetByName('Visits');
  const holSh=ss.getSheetByName('Holidays');
  const tz=Session.getScriptTimeZone()||'Asia/Kolkata';
  const key=Utilities.formatDate(date,tz,'yyyy-MM-dd');

  if(date.getDay()===0 || isHoliday_(holSh,key)){
    throw new Error('This date is marked as a holiday. Remove the holiday first if you want to enter a visit.');
  }

  const place=String(payload.place||'').trim();
  if(!place) throw new Error('Place is required.');

  const row=findDateRow_(visitSh,key,tz);
  const values=[date,Utilities.formatDate(date,tz,'EEEE'),place,String(payload.remarks||'').trim(),'Working',''];
  if(row) visitSh.getRange(row,1,1,6).setValues([values]);
  else visitSh.appendRow(values);

  return getMonthData(year,month);
}

function markHoliday(payload) {
  const date=new Date(payload.date+'T00:00:00');
  const year=date.getFullYear(),month=date.getMonth()+1;
  const key=payload.date;
  const reason=String(payload.reason||'Holiday').trim()||'Holiday';
  const ss=getDataSpreadsheet_();
  ensureCoreSheets_(ss);
  const holSh=ss.getSheetByName('Holidays');
  const visitSh=ss.getSheetByName('Visits');
  const tz=Session.getScriptTimeZone()||'Asia/Kolkata';

  if(date.getDay()!==0){
    const row=findDateRow_(holSh,key,tz);
    if(row) holSh.getRange(row,1,1,2).setValues([[date,reason]]);
    else holSh.appendRow([date,reason]);
  }
  const vrow=findDateRow_(visitSh,key,tz);
  if(vrow) visitSh.deleteRow(vrow);
  return getMonthData(year,month);
}

function unmarkHoliday(payload) {
  const date=new Date(payload.date+'T00:00:00');
  const ss=getDataSpreadsheet_();
  ensureCoreSheets_(ss);
  const holSh=ss.getSheetByName('Holidays');
  const tz=Session.getScriptTimeZone()||'Asia/Kolkata';
  const row=findDateRow_(holSh,payload.date,tz);
  if(row) holSh.deleteRow(row);
  return getMonthData(date.getFullYear(),date.getMonth()+1);
}


function getCounters(year, month) {
  const ss=getDataSpreadsheet_();
  ensureCoreSheets_(ss);
  const sh=ss.getSheetByName('Counters');
  const out={};
  if(sh.getLastRow()<2) return out;
  const vals=sh.getRange(2,1,sh.getLastRow()-1,4).getValues();
  for(let i=0;i<vals.length;i++){
    const r=vals[i];
    if(Number(r[0])===Number(year)&&Number(r[1])===Number(month)){
      const p=String(r[2]||'').trim();
      if(p) out[p.toLowerCase()]=Number(r[3]||0);
    }
  }
  return out;
}

function saveCounter(payload){
  const year=Number(payload.year),month=Number(payload.month);
  const place=String(payload.place||'').trim();
  const counter=Math.max(0,Number(payload.counter||0));
  if(!place) throw new Error('Place is required.');
  const ss=getDataSpreadsheet_();
  ensureCoreSheets_(ss);
  const sh=ss.getSheetByName('Counters');
  if(sh.getLastRow()>1){
    const vals=sh.getRange(2,1,sh.getLastRow()-1,4).getValues();
    for(let i=0;i<vals.length;i++){
      if(Number(vals[i][0])===year&&Number(vals[i][1])===month&&String(vals[i][2]).trim().toLowerCase()===place.toLowerCase()){
        sh.getRange(i+2,4).setValue(counter);
        return getMonthData(year,month);
      }
    }
  }
  sh.appendRow([year,month,place,counter]);
  return getMonthData(year,month);
}

function getSummary(year,month){
  return getMonthData(year,month).summary;
}

function exportMonthXlsx(year, month) {
  const data=getMonthData(year,month);
  const summary=data.summary || getSummary(year,month);
  const name=data.name || DEFAULT_NAME;
  const tz=Session.getScriptTimeZone()||'Asia/Kolkata';
  const monthName=Utilities.formatDate(new Date(year,month-1,1),tz,'MMMM');

  const temp=SpreadsheetApp.create('PJP_'+name+'_'+monthName+'_'+year);
  const sh=temp.getSheets()[0];
  sh.setName('PJP Report');

  const NAVY='#17365D', BLUE='#2F75B5', LIGHT='#DCE6F1', VERY_LIGHT='#F7F9FC';
  const RED='#D93025', WHITE='#FFFFFF', BORDER='#7F8C9A', GREY='#EDEDED';
  const TEXT='#1F2937', GREEN='#E2F0D9';

  // Reference-style layout:
  // A Date | B Day | C Place | D Remarks | E Combined Working | F spacer | G Place | H Days | I Counter
  // Google Apps Script setColumnWidth() uses PIXELS (not character counts).
  // Keep generous widths so the exported XLSX opens cleanly and remains readable.
  const widths=[110,95,185,255,125,22,300,72,105];
  widths.forEach((w,i)=>sh.setColumnWidth(i+1,w));
  // Freeze rows only. Do NOT freeze columns here because the report uses
  // merged cells spanning A:I; Google Sheets rejects merges that cross a
  // frozen/non-frozen column boundary.
  sh.setFrozenRows(4);
  sh.setTabColor(NAVY);

  sh.getRange('A1:I1').merge().setValue('PJP — MONTHLY VISIT & DISTRICT HEADQUARTER REPORT');
  sh.getRange('A1:I1').setBackground(NAVY).setFontColor(WHITE).setFontSize(16)
    .setFontWeight('bold').setFontFamily('Arial').setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1,36);

  const employeeId=data.employeeId || '';
  sh.getRange('A2:B2').merge().setValue('NAME: '+name);
  sh.getRange('C2:D2').merge().setValue('EMP ID: '+(employeeId||'—'));
  sh.getRange('E2:I2').merge().setValue('PERIOD: '+monthName.toUpperCase()+' '+year);
  sh.getRange('A2:I2').setBackground(LIGHT).setFontColor(TEXT).setFontWeight('bold')
    .setFontFamily('Arial').setVerticalAlignment('middle');
  sh.getRange('A2:D2').setHorizontalAlignment('left');
  sh.getRange('E2:I2').setHorizontalAlignment('right');
  sh.setRowHeight(2,25);

  // Daily header — same visual logic as the user's reference.
  // Combined Working is used only as a holiday marker: normal days stay blank; holidays show 'Holiday'.
  const headerRow=4, startRow=5;
  sh.getRange(headerRow,1,1,5).setValues([[
    'Date','DAY','Place','Remarks','Combined Working'
  ]]);
  sh.getRange(headerRow,1,1,5).setBackground(BLUE).setFontColor(WHITE)
    .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true).setBorder(true,true,true,true,true,true,BORDER,SpreadsheetApp.BorderStyle.SOLID);

  // Right-side summary header
  sh.getRange(headerRow,7,1,3).setValues([[
    'No of days spend in District Headquarter','Days','No of Counter'
  ]]);
  sh.getRange(headerRow,7,1,3).setBackground(BLUE).setFontColor(WHITE)
    .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true).setBorder(true,true,true,true,true,true,BORDER,SpreadsheetApp.BorderStyle.SOLID);

  const dailyRows=data.days.map(r=>[
    new Date(r.date+'T00:00:00'),
    r.day,
    r.isHoliday?'Holiday':(r.place||''),
    r.isHoliday?(r.holidayReason||'Holiday'):(r.remarks||''),
    // Combined Working: normal working days stay blank; holiday days show 'Holiday'.
    (r.isHoliday ? 'Holiday' : '')
  ]);
  if(dailyRows.length){
    sh.getRange(startRow,1,dailyRows.length,5).setValues(dailyRows);
    sh.getRange(startRow,1,dailyRows.length,5)
      .setFontFamily('Arial').setVerticalAlignment('middle').setWrap(true)
      .setBorder(true,true,true,true,true,true,BORDER,SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(startRow,1,dailyRows.length,1).setNumberFormat('dd-mmm-yyyy').setHorizontalAlignment('center');
    sh.getRange(startRow,2,dailyRows.length,1).setHorizontalAlignment('center');
    sh.getRange(startRow,5,dailyRows.length,1).setHorizontalAlignment('center');
    for(let i=0;i<dailyRows.length;i++){
      const rr=startRow+i;
      sh.setRowHeight(rr,30);
      if(data.days[i].isHoliday){
        sh.getRange(rr,1,1,5).setBackground(RED).setFontColor(WHITE).setFontWeight('bold');
      }else if(i%2===0){
        sh.getRange(rr,1,1,5).setBackground(VERY_LIGHT);
      }
    }
  }

  // Place-wise summary
  const summaryRows=summary.places.map(x=>[x.place,x.days,Number(x.counter||0)]);
  if(summaryRows.length){
    sh.getRange(startRow,7,summaryRows.length,3).setValues(summaryRows);
    sh.getRange(startRow,7,summaryRows.length,3)
      .setFontFamily('Arial').setVerticalAlignment('middle').setWrap(true)
      .setBorder(true,true,true,true,true,true,BORDER,SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(startRow,8,summaryRows.length,2).setHorizontalAlignment('center');
    for(let i=0;i<summaryRows.length;i++){
      sh.setRowHeight(startRow+i,30);
      if(i%2===0) sh.getRange(startRow+i,7,1,3).setBackground(GREY);
    }
  }

  const summaryTotalRow=startRow+summaryRows.length;
  sh.getRange(summaryTotalRow,7,1,3).setValues([[
    'TOTAL',
    summary.places.reduce((s,x)=>s+Number(x.days||0),0),
    summary.places.reduce((s,x)=>s+Number(x.counter||0),0)
  ]]);
  sh.getRange(summaryTotalRow,7,1,3).setBackground(LIGHT).setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true,true,true,true,true,true,BORDER,SpreadsheetApp.BorderStyle.SOLID);
  // Keep summary totals visually distinct and readable.
  sh.getRange(summaryTotalRow,7,1,3).setFontSize(10);
  sh.getRange(startRow,7,Math.max(1,summaryRows.length),3).setWrap(true);

  // Monthly totals beneath the daily table
  const totalsRow=startRow+dailyRows.length+2;
  sh.getRange(totalsRow,1,4,2).setValues([
    ['Total Calendar Days',summary.totalDays],
    ['Working Days',summary.workingDays],
    ['Holiday Days',summary.holidayDays],
    ['Total Visit Days',summary.visitDays]
  ]);
  sh.getRange(totalsRow,1,4,2).setBorder(true,true,true,true,true,true,BORDER,SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(totalsRow,1,4,1).setBackground(LIGHT).setFontWeight('bold');
  sh.getRange(totalsRow,2,4,1).setBackground(VERY_LIGHT).setFontWeight('bold').setHorizontalAlignment('center');

  // Grand total / signature-style footer
  const footerRow=Math.max(totalsRow+6,summaryTotalRow+3);
  sh.getRange(footerRow,1,1,9).merge()
    .setValue('PJP Report • '+name+' • '+monthName+' '+year);
  sh.getRange(footerRow,1,1,9).setFontColor('#6B7280').setFontSize(9)
    .setHorizontalAlignment('center');

  // Final workbook layout / readability settings.
  sh.getRange(1,1,footerRow,9)
    .setFontFamily('Arial')
    .setVerticalAlignment('middle');
  sh.setHiddenGridlines(true);

  // Explicit row heights: no compressed/tiny cells.
  for(let r=1;r<=footerRow;r++) sh.setRowHeight(r,30);
  sh.setRowHeight(1,40);
  sh.setRowHeight(2,28);
  sh.setRowHeight(3,12);
  sh.setRowHeight(4,48);
  for(let r=startRow;r<startRow+dailyRows.length;r++) sh.setRowHeight(r,34);

  // Consistent alignment and wrapping.
  sh.getRange(1,1,footerRow,9).setWrap(true);
  sh.getRange(startRow,3,dailyRows.length,2).setHorizontalAlignment('left');
  sh.getRange(startRow,1,dailyRows.length,2).setHorizontalAlignment('center');
  sh.getRange(startRow,5,dailyRows.length,1).setHorizontalAlignment('center');

  // Give the summary section a stable visual hierarchy.
  if(summaryRows.length){
    sh.getRange(startRow,7,summaryRows.length,1).setHorizontalAlignment('left');
    sh.getRange(startRow,8,summaryRows.length,2).setHorizontalAlignment('center');
  }

  // Improve print/export behavior.
  // Keep row freeze only; merged A:I/A:D/E:I ranges are compatible with it.
  sh.setFrozenRows(4);
  try{
    sh.setPrintGridlines(false);
    sh.setPageOrientation(SpreadsheetApp.PageOrientation.LANDSCAPE);
    sh.setPaperSize(SpreadsheetApp.PaperSize.A4);
    sh.setFitToPage(true);
    sh.setPageMargins(0.25,0.25,0.35,0.35);
  }catch(e){}

  SpreadsheetApp.flush();

  const token=ScriptApp.getOAuthToken();
  const url='https://docs.google.com/spreadsheets/d/'+temp.getId()+'/export?format=xlsx';
  const blob=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+token}}).getBlob()
    .setName('PJP_'+name+'_'+monthName+'_'+year+'.xlsx');

  const exportFile=DriveApp.createFile(blob);
  try{exportFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);}catch(e){}
  const fileId=exportFile.getId();
  DriveApp.getFileById(temp.getId()).setTrashed(true);

  return {
    filename:exportFile.getName(),
    fileId:fileId,
    downloadUrl:'https://drive.google.com/uc?export=download&id='+fileId,
    fileUrl:'https://drive.google.com/file/d/'+fileId+'/view'
  };
}

function ensureCoreSheets_(ss){
  ensureSheet_(ss,'Visits',['Date','Day','Place','Remarks','Status','Holiday Reason']);
  ensureSheet_(ss,'Holidays',['Date','Reason']);
  ensureSheet_(ss,'Counters',['Year','Month','Place','Counter']);
  const settings=ensureSheet_(ss,'Settings',['Key','Value']);
  if(settings.getLastRow()===1) settings.appendRow(['Name',DEFAULT_NAME]);
}

function ensureSheet_(ss,name,headers) {
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  if(sh.getLastRow()===0) sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  return sh;
}

function findDateRow_(sh,key,tz) {
  if(!sh || sh.getLastRow()<2) return 0;
  const vals=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(let i=0;i<vals.length;i++){
    if(vals[i][0] && Utilities.formatDate(new Date(vals[i][0]),tz,'yyyy-MM-dd')===key) return i+2;
  }
  return 0;
}

function isHoliday_(sh,key) {
  if(!sh || sh.getLastRow()<2) return false;
  const tz=Session.getScriptTimeZone()||'Asia/Kolkata';
  const vals=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  return vals.some(r=>r[0] && Utilities.formatDate(new Date(r[0]),tz,'yyyy-MM-dd')===key);
}

function resetDatabaseSpreadsheet() {
  PropertiesService.getScriptProperties().deleteProperty('PJP_DATA_SPREADSHEET_ID');
  return 'Database link cleared. The next app load will create a new PJP_SURAJ_Data spreadsheet.';
}

function testDatabase() {
  const ss = getDataSpreadsheet_();
  ensureCoreSheets_(ss);
  return 'Database OK: ' + ss.getName() + ' | ID: ' + ss.getId();
}
