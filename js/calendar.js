async function loadCalendar(){
  const badge=el('calBadgeSelect').value||STATE.user.badge;
  const month=el('calMonth').value||new Date().toISOString().substring(0,7);
  if(!badge)return;
  const [yr,mn]=month.split('-').map(Number);
  const dateFrom=`01/${String(mn).padStart(2,'0')}/${yr}`;
  const lastDay=new Date(yr,mn,0).getDate();
  const dateTo=`${lastDay}/${String(mn).padStart(2,'0')}/${yr}`;
  el('calBadgeLabel').textContent='Badge: '+badge+' | '+month;
  spin(true);
  try{
    if(STATE.isDemo){renderCalDemo(yr,mn);spin(false);return}
    const monthStr = yr + '-' + String(mn).padStart(2,'0');
    const [d, sum] = await Promise.all([
      zkAPI('/api/employee/'+badge+'/report?from='+dateFrom+'&to='+dateTo+'&source=db'),
      zkAPI('/api/employee/'+badge+'/monthly-summary?month='+monthStr)
    ]);
    STATE.calData=d;
    el('calWorkDays').textContent=sum.working_days||0;
    el('calPresent').textContent=sum.present_days||0;
    el('calAbsent').textContent=sum.absent_days||0;
    el('calLate').textContent=sum.late_count||0;
    el('calEarly').textContent=sum.early_departure_count||0;
    el('calHD').textContent=sum.holiday_duty_days||0;
    el('calShiftInfo').textContent=d.shift?`Shift: ${d.shift.start} → ${d.shift.end} (grace: ${d.shift.grace}m)`:'Shift info unavailable';
    renderCalendar(d.days||[],yr,mn);
  }catch(e){toast('❌ Calendar: '+e.message)}
  spin(false);
}
function renderCalDemo(yr,mn){
  const days=[];
  const lastDay=new Date(yr,mn,0).getDate();
  for(let d=1;d<=lastDay;d++){
    const dt=new Date(yr,mn-1,d);
    const wd=dt.getDay();
    const isWork=wd>=0&&wd<=4;
    const isPresent=isWork&&Math.random()>0.15;
    days.push({date:`${String(d).padStart(2,'0')}/${String(mn).padStart(2,'0')}/${yr}`,date_iso:`${yr}-${String(mn).padStart(2,'0')}-${String(d).padStart(2,'0')}`,day:dt.toLocaleDateString('en-GB',{weekday:'long'}),working_day:isWork,present:isPresent,punches:isPresent?[{time:'07:'+String(20+Math.floor(Math.random()*30)).padStart(2,'0')+':00',machine:'Demo'}]:[],late:isPresent&&Math.random()>0.8,late_mins:15,early_departure:false,early_mins:0});
  }
  el('calWorkDays').textContent=days.filter(d=>d.working_day).length;
  el('calPresent').textContent=days.filter(d=>d.present).length;
  el('calAbsent').textContent=days.filter(d=>d.working_day&&!d.present).length;
  el('calLate').textContent=days.filter(d=>d.late).length;
  el('calEarly').textContent=0;
  el('calHD').textContent=0;
  el('calShiftInfo').textContent='Demo shift: 07:30 → 15:00 (grace: 15m)';
  renderCalendar(days,yr,mn);
}
function renderCalendar(days,yr,mn){
  const firstDay=new Date(yr,mn-1,1).getDay();
  const today=new Date().toISOString().substring(0,10);
  const grid=el('calGrid');
  // Clear except headers
  while(grid.children.length>7)grid.lastChild.remove();
  // Blank cells
  for(let i=0;i<firstDay;i++){const b=document.createElement('div');b.className='cal-day empty';grid.appendChild(b)}
  days.forEach(day=>{
    const d=document.createElement('div');
    let cls='cal-day';
    if(!day.working_day)cls+=' off';
    else if(day.present)cls+=' present';
    else cls+=' absent';
    if(day.date_iso===today)cls+=' today-date';
    if(day.holiday_duty)cls+=' holiday';
    d.className=cls;
    if(day.punches && day.punches.length) {
      d.title = 'Punches:\n' + day.punches.map(p=>p.time.substring(0,5) + ' - ' + (p.machine||'Device')).join('\n');
      if(day.holiday_duty) d.title += '\n[Holiday Duty]';
    } else if(day.holiday_duty) {
      d.title = '[Holiday Duty]';
    }
    d.innerHTML=`<div class="cdn">${day.date.substring(0,2)}</div>${day.punches&&day.punches.length?`<div class="cdpunch">${day.punches[0].time.substring(0,5)}${day.punches.length>1?' +'+(day.punches.length-1)+'':''}</div>`:''}${day.late?`<div class="cdlate">LATE</div>`:''}`;
    d.onclick=()=>openCalDay(day);
    grid.appendChild(d);
  });
}
function openCalDay(day){
  el('calDayTitle').textContent=day.date+' — '+day.day;
  const statusTag=!day.working_day?'<span class="tag tag-off">Off Day</span>':day.present?'<span class="tag tag-present">Present</span>':'<span class="tag tag-absent">Absent</span>';
  el('calDayBody').innerHTML=`
    <div style="margin-bottom:12px">${statusTag}${day.late?'<span class="tag tag-late" style="margin-left:6px">Late +'+day.late_mins+'m</span>':''}${day.holiday_duty?'<span class="tag tag-holiday" style="margin-left:6px">Holiday Duty</span>':''}</div>
    <div style="font-size:12px;color:var(--text2);margin-bottom:8px;font-family:var(--mono)">${day.punches&&day.punches.length?day.punches.length+' punch record(s)':'No punches recorded'}</div>
    ${(day.punches||[]).map(p=>`<div style="display:flex;justify-content:space-between;padding:6px 8px;background:var(--card2);border-radius:4px;margin-bottom:4px;font-size:12px;font-family:var(--mono)"><span>${p.time}</span><span class="text-muted">${p.machine||'—'}</span></div>`).join('')}
  `;
  openModal('calDayModal');
}
async function initCalendarSelect(){
  el('calMonth').value=new Date().toISOString().substring(0,7);
  const sel=el('calBadgeSelect');
  sel.innerHTML='';
  const emps=STATE.isDemo?DEMO_EMPLOYEES:STATE.empList;
  if(STATE.user.badge&&STATE.user.role!=='admin'){
    // Add "(Me)" entry first only if the user is not in the full employee list
    const alreadyInList=emps.some(e=>(e.code||e.badge)===STATE.user.badge);
    if(!alreadyInList){
      const opt=document.createElement('option');opt.value=STATE.user.badge;opt.textContent=STATE.user.name+' (Me)';sel.appendChild(opt);
    }
  }
  emps.forEach(e=>{
    const opt=document.createElement('option');opt.value=e.code||e.badge;opt.textContent=e.name+' ('+(e.code||e.badge)+')';
    if((e.code||e.badge)===STATE.user.badge)opt.textContent+=' ← Me';
    sel.appendChild(opt);
  });
  loadCalendar();
}
async function exportCalendar(){
  if(STATE.isDemo){toast('🎮 Demo');return}
  const badge=el('calBadgeSelect').value||STATE.user.badge;
  const month=el('calMonth').value||new Date().toISOString().substring(0,7);
  const [yr,mn]=month.split('-').map(Number);
  const dateFrom=`01/${String(mn).padStart(2,'0')}/${yr}`;
  const lastDay=new Date(yr,mn,0).getDate();
  const dateTo=`${lastDay}/${String(mn).padStart(2,'0')}/${yr}`;
  window.open(CFG.zkUrl+'/api/employee/'+badge+'/report/export?from='+dateFrom+'&to='+dateTo,'_blank');
}

