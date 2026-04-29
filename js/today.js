// ===========================================================================
//  TODAY
// ===========================================================================
let _todayRaw={absent:[],present:[],off:[]};
async function loadToday(){
  STATE.todayLoaded=true;
  el('todayDate').textContent=new Date().toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
  spin(true);
  try{
    let absent=[],present=[],off=[];
    if(STATE.isDemo){
      absent=DEMO_EMPLOYEES.filter(e=>DEMO_ABSENT.includes(e.code));
      present=DEMO_EMPLOYEES.filter(e=>DEMO_PRESENT.includes(e.code));
    }else{
      const d=await zkAPI('/api/today');
      STATE.todayData=d;
      absent=d.absent||[];present=d.present||[];off=d.off_today||[];
    }
    _todayRaw={absent,present,off};
    el('tPresent').textContent=present.length;
    el('tAbsent').textContent=absent.length;
    el('tOff').textContent=off.length;
    el('tAbsentCount2').textContent=absent.length+' employees';
    el('tPresentCount2').textContent=present.length+' employees';
    // Populate dept filters
    const depts=[...new Set([...absent,...present,...off].map(e=>e.dept||'').filter(Boolean))].sort();
    ['absentDeptFilter','presentDeptFilter'].forEach(id=>{
      el(id).innerHTML='<option value="">All Depts</option>'+depts.map(d=>`<option>${esc(d)}</option>`).join('');
    });
    renderTodayTable('absent',absent);
    renderTodayTable('present',present);
  }catch(e){toast('❌ Failed to load today: '+e.message)}
  spin(false);
}
function renderTodayTable(which,rows){
  const tbody=el(which==='absent'?'absentBody':'presentBody');
  tbody.innerHTML=rows.length?rows.map(e=>`<tr onclick="selectEmpQuick('${esc(e.code||e.badge)}','${esc(e.name)}','${esc(e.dept)}')" style="cursor:pointer">
    <td class="td-mono">${esc(e.code||e.badge||'—')}</td><td>${esc(e.name)}</td><td><span class="tag tag-off">${esc(e.dept)}</span></td>
  </tr>`).join(''):'<tr><td colspan="3" style="text-align:center;padding:20px;color:var(--text3)">None</td></tr>';
}
function filterToday(which,q){
  const dept=el(which==='absent'?'absentDeptFilter':'presentDeptFilter').value;
  let rows=_todayRaw[which]||[];
  if(q)rows=rows.filter(e=>(e.name||'').toLowerCase().includes(q.toLowerCase())||(e.code||'').includes(q));
  if(dept)rows=rows.filter(e=>e.dept===dept);
  renderTodayTable(which,rows);
}
async function exportToday(){
  if(STATE.isDemo){toast('🎮 Demo — export unavailable');return}
  window.open(CFG.zkUrl+'/api/today/export','_blank');
}

