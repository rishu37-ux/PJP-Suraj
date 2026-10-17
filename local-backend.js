/**
 * PJP SURAJ - Local (phone) backend
 *
 * Same functions as apps-script/Code.gs, but data is stored in the phone's
 * browser (localStorage) instead of a Google Sheet. A small shim exposes them
 * as google.script.run so index.html works unchanged in both versions.
 */
(function(){
  const DB_KEY='PJP_SURAJ_DB_V1';
  const DEFAULT_NAME='SURAJ';
  const DAY_NAMES=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const MONTH_NAMES=['January','February','March','April','May','June','July','August','September','October','November','December'];

  function emptyDb_(){ return {profile:{name:DEFAULT_NAME,employeeId:''},visits:{},holidays:{},counters:{}}; }
  function loadDb_(){
    try{
      const raw=localStorage.getItem(DB_KEY);
      if(!raw) return emptyDb_();
      const db=Object.assign(emptyDb_(),JSON.parse(raw));
      db.profile=Object.assign({name:DEFAULT_NAME,employeeId:''},db.profile||{});
      return db;
    }catch(e){ return emptyDb_(); }
  }
  function saveDb_(db){
    try{ localStorage.setItem(DB_KEY,JSON.stringify(db)); }
    catch(e){ throw new Error('Could not save on this phone (storage blocked or full).'); }
  }
  function pad_(n){ return String(n).padStart(2,'0'); }
  function key_(d){ return d.getFullYear()+'-'+pad_(d.getMonth()+1)+'-'+pad_(d.getDate()); }
  function parseKey_(k){ return new Date(String(k).slice(0,10)+'T00:00:00'); }
  function counterKey_(y,m){ return y+'-'+m; }

  function getProfile(){ return Object.assign({},loadDb_().profile); }
  function saveProfile(payload){
    const db=loadDb_();
    db.profile={
      name:String(payload && payload.name || DEFAULT_NAME).trim() || DEFAULT_NAME,
      employeeId:String(payload && payload.employeeId || '').trim()
    };
    saveDb_(db);
    return Object.assign({},db.profile);
  }

  function getMonthData(year,month){
    year=Number(year); month=Number(month);
    const db=loadDb_();
    const daysInMonth=new Date(year,month,0).getDate();
    const counters=db.counters[counterKey_(year,month)]||{};
    const rows=[], placeCounts={};
    let working=0, holidayDays=0, visitDays=0;

    for(let day=1;day<=daysInMonth;day++){
      const dt=new Date(year,month-1,day);
      const key=key_(dt);
      const sunday=dt.getDay()===0;
      const manualHoliday=db.holidays[key];
      const isHoliday=sunday || !!manualHoliday;
      const reason=sunday?'Sunday':(manualHoliday||'');
      const v=db.visits[key]||{};
      const place=isHoliday?'':(v.place||'');
      rows.push({date:key,day:DAY_NAMES[dt.getDay()],dayNumber:day,place,remarks:v.remarks||'',isHoliday,holidayReason:reason});

      if(isHoliday) holidayDays++;
      else{
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
    const name=db.profile.name||DEFAULT_NAME, employeeId=db.profile.employeeId||'';
    const summary={name,employeeId,year,month,totalDays:daysInMonth,workingDays:working,holidayDays,visitDays,places};
    return {name,employeeId,year,month,days:rows,summary};
  }

  function saveVisitsBulk(payload){
    const assignments=Array.isArray(payload.assignments)?payload.assignments:[];
    const remarks=String(payload.remarks||'').trim();
    if(!assignments.length) throw new Error('Please select at least one place and date.');
    const db=loadDb_();

    const combined={};
    assignments.forEach(a=>{
      const place=String(a.place||'').trim();
      if(!place) return;
      (Array.isArray(a.dates)?a.dates:[]).forEach(k=>{
        k=String(k);
        if(combined[k] && combined[k]!==place) throw new Error('Date '+k+' is assigned to more than one place in this batch. Please keep it with one place.');
        combined[k]=place;
      });
    });

    const keys=Object.keys(combined).sort();
    const occupied=keys.filter(k=>db.visits[k]);
    if(occupied.length) throw new Error('These dates are already saved: '+occupied.join(', ')+'. Use Fill Visit to reassign an existing date.');
    keys.forEach(k=>{
      if(parseKey_(k).getDay()===0 || db.holidays[k]) return;
      db.visits[k]={place:combined[k],remarks};
    });
    saveDb_(db);
    const first=parseKey_(keys[0]);
    return getMonthData(first.getFullYear(),first.getMonth()+1);
  }

  function clearAllSavedVisits(year,month){
    year=Number(year); month=Number(month);
    if(!year || !month || month<1 || month>12) throw new Error('Invalid month.');
    const db=loadDb_();
    const prefix=year+'-'+pad_(month)+'-';
    let removed=0;
    Object.keys(db.visits).forEach(k=>{ if(k.startsWith(prefix)){ delete db.visits[k]; removed++; } });
    saveDb_(db);
    return {removed,data:getMonthData(year,month)};
  }

  function saveVisit(payload){
    const date=parseKey_(payload.date), key=key_(date);
    const db=loadDb_();
    if(date.getDay()===0 || db.holidays[key]) throw new Error('This date is marked as a holiday. Remove the holiday first if you want to enter a visit.');
    const place=String(payload.place||'').trim();
    if(!place) throw new Error('Place is required.');
    db.visits[key]={place,remarks:String(payload.remarks||'').trim()};
    saveDb_(db);
    return getMonthData(date.getFullYear(),date.getMonth()+1);
  }

  function markHoliday(payload){
    const date=parseKey_(payload.date), key=key_(date);
    const db=loadDb_();
    if(date.getDay()!==0) db.holidays[key]=String(payload.reason||'Holiday').trim()||'Holiday';
    delete db.visits[key];
    saveDb_(db);
    return getMonthData(date.getFullYear(),date.getMonth()+1);
  }

  function unmarkHoliday(payload){
    const date=parseKey_(payload.date);
    const db=loadDb_();
    delete db.holidays[key_(date)];
    saveDb_(db);
    return getMonthData(date.getFullYear(),date.getMonth()+1);
  }

  function saveCounter(payload){
    const year=Number(payload.year), month=Number(payload.month);
    const place=String(payload.place||'').trim();
    if(!place) throw new Error('Place is required.');
    const db=loadDb_();
    const ck=counterKey_(year,month);
    db.counters[ck]=db.counters[ck]||{};
    db.counters[ck][place.toLowerCase()]=Math.max(0,Number(payload.counter||0));
    saveDb_(db);
    return getMonthData(year,month);
  }

  /* ---------- Excel export (same layout as the Apps Script report) ---------- */
  function exportMonthXlsx(year,month){
    if(typeof XLSX==='undefined') throw new Error('Excel library not loaded. Open the app once with internet.');
    year=Number(year); month=Number(month);
    const data=getMonthData(year,month), summary=data.summary;
    const name=data.name||DEFAULT_NAME, employeeId=data.employeeId||'';
    const monthName=MONTH_NAMES[month-1];

    const NAVY='17365D', BLUE='2F75B5', LIGHT='DCE6F1', VERY_LIGHT='F7F9FC';
    const RED='D93025', WHITE='FFFFFF', BORDER='7F8C9A', GREY='EDEDED', TEXT='1F2937';
    const side={style:'thin',color:{rgb:BORDER}};
    const border={top:side,bottom:side,left:side,right:side};
    const base=(extra)=>Object.assign({font:{name:'Arial',sz:10,color:{rgb:TEXT}},alignment:{vertical:'center',wrapText:true}},extra||{});

    const ws={}, merges=[], rowsH=[];
    const set=(r,c,v,style,fmt)=>{
      const ref=XLSX.utils.encode_cell({r,c});
      const cell={v,t:typeof v==='number'?'n':(v instanceof Date?'d':'s'),s:style};
      if(fmt) cell.z=fmt;
      ws[ref]=cell;
    };
    const fill=rgb=>({patternType:'solid',fgColor:{rgb}});

    // Row 1: title (A1:I1)
    const titleStyle=base({fill:fill(NAVY),font:{name:'Arial',sz:16,bold:true,color:{rgb:WHITE}},alignment:{horizontal:'center',vertical:'center'}});
    for(let c=0;c<9;c++) set(0,c,c===0?'PJP — MONTHLY VISIT & DISTRICT HEADQUARTER REPORT':'',titleStyle);
    merges.push({s:{r:0,c:0},e:{r:0,c:8}});

    // Row 2: name / emp id / period
    const infoL=base({fill:fill(LIGHT),font:{name:'Arial',sz:10,bold:true,color:{rgb:TEXT}},alignment:{horizontal:'left',vertical:'center'}});
    const infoR=base({fill:fill(LIGHT),font:{name:'Arial',sz:10,bold:true,color:{rgb:TEXT}},alignment:{horizontal:'right',vertical:'center'}});
    for(let c=0;c<9;c++) set(1,c,'',c<4?infoL:infoR);
    set(1,0,'NAME: '+name,infoL);
    set(1,2,'EMP ID: '+(employeeId||'—'),infoL);
    set(1,4,'PERIOD: '+monthName.toUpperCase()+' '+year,infoR);
    merges.push({s:{r:1,c:0},e:{r:1,c:1}},{s:{r:1,c:2},e:{r:1,c:3}},{s:{r:1,c:4},e:{r:1,c:8}});

    // Row 4: headers
    const head=base({fill:fill(BLUE),font:{name:'Arial',sz:10,bold:true,color:{rgb:WHITE}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border});
    ['Date','DAY','Place','Remarks','Combined Working'].forEach((h,c)=>set(3,c,h,head));
    ['No of days spend in District Headquarter','Days','No of Counter'].forEach((h,c)=>set(3,6+c,h,head));

    // Daily rows
    const startRow=4;
    data.days.forEach((r,i)=>{
      const rr=startRow+i;
      const bg=r.isHoliday?RED:(i%2===0?VERY_LIGHT:WHITE);
      const font={name:'Arial',sz:10,bold:r.isHoliday,color:{rgb:r.isHoliday?WHITE:TEXT}};
      const mk=h=>base({fill:fill(bg),font,border,alignment:{horizontal:h,vertical:'center',wrapText:true}});
      set(rr,0,r.date.slice(8)+'-'+monthName.slice(0,3)+'-'+year,mk('center'));
      set(rr,1,r.day,mk('center'));
      set(rr,2,r.isHoliday?'Holiday':(r.place||''),mk('left'));
      set(rr,3,r.isHoliday?(r.holidayReason||'Holiday'):(r.remarks||''),mk('left'));
      set(rr,4,r.isHoliday?'Holiday':'',mk('center'));
    });

    // Place-wise summary
    summary.places.forEach((x,i)=>{
      const rr=startRow+i, bg=i%2===0?GREY:WHITE;
      set(rr,6,x.place,base({fill:fill(bg),border,alignment:{horizontal:'left',vertical:'center',wrapText:true}}));
      set(rr,7,x.days,base({fill:fill(bg),border,alignment:{horizontal:'center',vertical:'center'}}));
      set(rr,8,Number(x.counter||0),base({fill:fill(bg),border,alignment:{horizontal:'center',vertical:'center'}}));
    });
    const sumTotalRow=startRow+summary.places.length;
    const totStyle=base({fill:fill(LIGHT),font:{name:'Arial',sz:10,bold:true,color:{rgb:TEXT}},border,alignment:{horizontal:'center',vertical:'center'}});
    set(sumTotalRow,6,'TOTAL',totStyle);
    set(sumTotalRow,7,summary.places.reduce((s,x)=>s+Number(x.days||0),0),totStyle);
    set(sumTotalRow,8,summary.places.reduce((s,x)=>s+Number(x.counter||0),0),totStyle);

    // Monthly totals
    const totalsRow=startRow+data.days.length+2;
    [['Total Calendar Days',summary.totalDays],['Working Days',summary.workingDays],['Holiday Days',summary.holidayDays],['Total Visit Days',summary.visitDays]]
      .forEach(([label,val],i)=>{
        set(totalsRow+i,0,label,base({fill:fill(LIGHT),font:{name:'Arial',sz:10,bold:true,color:{rgb:TEXT}},border}));
        set(totalsRow+i,1,val,base({fill:fill(VERY_LIGHT),font:{name:'Arial',sz:10,bold:true,color:{rgb:TEXT}},border,alignment:{horizontal:'center',vertical:'center'}}));
      });

    // Footer
    const footerRow=Math.max(totalsRow+5,sumTotalRow+2);
    const footStyle=base({font:{name:'Arial',sz:9,color:{rgb:'6B7280'}},alignment:{horizontal:'center',vertical:'center'}});
    for(let c=0;c<9;c++) set(footerRow,c,c===0?'PJP Report • '+name+' • '+monthName+' '+year:'',footStyle);
    merges.push({s:{r:footerRow,c:0},e:{r:footerRow,c:8}});

    for(let r=0;r<=footerRow;r++) rowsH[r]={hpx:30};
    rowsH[0]={hpx:40}; rowsH[1]={hpx:28}; rowsH[2]={hpx:12}; rowsH[3]={hpx:48};
    for(let r=startRow;r<startRow+data.days.length;r++) rowsH[r]={hpx:34};

    ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:footerRow,c:8}});
    ws['!merges']=merges;
    ws['!cols']=[110,95,185,255,125,22,300,72,105].map(wpx=>({wpx}));
    ws['!rows']=rowsH;
    ws['!pageSetup']={orientation:'landscape',paperSize:9,fitToWidth:1,fitToHeight:0};
    ws['!margins']={left:0.25,right:0.25,top:0.35,bottom:0.35,header:0.2,footer:0.2};

    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,'PJP Report');
    const filename='PJP_'+name.replace(/[^\w\- ]+/g,'').trim().replace(/\s+/g,'_')+'_'+monthName+'_'+year+'.xlsx';
    XLSX.writeFile(wb,filename,{cellStyles:true});
    return {filename,local:true};
  }

  /* ---------- Backup / restore (phone data is per-browser) ---------- */
  function exportBackup(){
    const blob=new Blob([JSON.stringify(loadDb_(),null,2)],{type:'application/json'});
    const a=document.createElement('a');
    a.href=URL.createObjectURL(blob);
    a.download='PJP_Backup_'+key_(new Date())+'.json';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(a.href),2000);
  }
  function importBackup(text){
    const parsed=JSON.parse(text);
    if(!parsed || typeof parsed!=='object' || !parsed.visits) throw new Error('This is not a PJP backup file.');
    saveDb_(Object.assign(emptyDb_(),parsed));
  }

  const api={getProfile,saveProfile,getMonthData,saveVisitsBulk,clearAllSavedVisits,saveVisit,markHoliday,unmarkHoliday,saveCounter,exportMonthXlsx,
    getSummary:(y,m)=>getMonthData(y,m).summary};

  // google.script.run compatible shim
  function runner(ok,fail){
    return new Proxy({},{get(_,prop){
      if(prop==='withSuccessHandler') return fn=>runner(fn,fail);
      if(prop==='withFailureHandler') return fn=>runner(ok,fn);
      if(!api[prop]) return undefined;
      return (...args)=>setTimeout(()=>{
        let result;
        try{ result=api[prop](...args); }
        catch(e){ if(fail) fail(e); else alert(e.message); return; }
        if(ok) ok(result);
      },0);
    }});
  }
  window.PJP_LOCAL={exportBackup,importBackup};
  if(!(window.google && window.google.script)) window.google={script:{run:runner(null,null)}};
})();
