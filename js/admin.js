// ===========================================================================
//  ADMIN PANEL
// ===========================================================================
async function loadDBStats(){
  if(STATE.isDemo){el('dbStatsBody').innerHTML='<div class="text-muted" style="font-size:12px">Demo mode — no real DB stats.</div>';return}
  try{
    const d=await zkAPI('/api/db/status');
    el('dbStatsBody').innerHTML=`
      <div class="dc-row"><span class="label">DB Size</span><span class="val">${d.size_mb} MB</span></div>
      <div class="dc-row"><span class="label">Employees</span><span class="val">${d.employees}</span></div>
      <div class="dc-row"><span class="label">Punch Records</span><span class="val">${(d.punch_records||0).toLocaleString()}</span></div>
      <div class="dc-row"><span class="label">First Punch</span><span class="val text-mono" style="font-size:10px">${fmtDate(d.first_punch)}</span></div>
      <div class="dc-row"><span class="label">Last Punch</span><span class="val text-mono" style="font-size:10px">${fmtDate(d.last_punch)}</span></div>
      <div class="dc-row"><span class="label">Unknown Users</span><span class="val ${d.unknown_users>0?'text-yellow':''}">${d.unknown_users}</span></div>
    `;
    el('sbPunchCount').textContent=(d.punch_records||0).toLocaleString();
  }catch(e){el('dbStatsBody').innerHTML='<div class="text-red">'+esc(e.message)+'</div>'}
}
async function runSync(mode){
  if(STATE.isDemo){toast('🎮 Demo');return}
  el('syncMsg').textContent='⏳ Syncing...';el('syncMsg').style.color='var(--yellow)';
  try{
    const path=mode==='full'?'/api/cache/refresh':'/api/cache/refresh';
    const d=await zkAPI(path,{method:'POST'});
    el('syncMsg').textContent='✅ Sync complete';el('syncMsg').style.color='var(--green)';
    toast('✅ Sync done');loadDBStats();
  }catch(e){el('syncMsg').textContent='❌ '+e.message;el('syncMsg').style.color='var(--red)'}
}
async function createBackup(){
  if(STATE.isDemo){toast('🎮 Demo');return}
  try{
    const d=await zkAPI('/api/db/backup',{method:'POST'});
    el('backupMsg').textContent='✅ Backup created';el('backupMsg').style.color='var(--green)';
    toast('✅ Backup created');
  }catch(e){el('backupMsg').textContent='❌ '+e.message;el('backupMsg').style.color='var(--red)'}
}
async function loadUnmapped(){
  if(STATE.isDemo){el('unmappedBody').innerHTML='<tr><td colspan="4" class="text-muted" style="padding:12px;text-align:center">No unmapped users (demo)</td></tr>';return}
  try{
    const d=await zkAPI('/api/db/unknown-users');
    el('unmappedBody').innerHTML=(d||[]).map(u=>`<tr>
      <td class="td-mono">${esc(u.uid)}</td>
      <td class="text-muted">${esc(u.device_ip)}</td>
      <td class="text-muted td-mono" style="font-size:10px">${fmtDateTime(u.seen_at)}</td>
      <td><button class="btn btn-secondary btn-sm" onclick="resolveUnmapped('${esc(u.device_ip)}','${esc(u.uid)}')">🔗 Map Badge</button></td>
    </tr>`).join('')||'<tr><td colspan="4" class="text-muted" style="padding:12px;text-align:center">No unmapped users ✅</td></tr>';
  }catch(e){}
}
async function resolveUnmapped(ip,uid){
  const badge=prompt('Map device UID "'+uid+'" (device: '+ip+') to employee badge number:\n\nEnter the Badgenumber from the employee list (e.g. 1712)',uid);
  if(badge===null)return;
  const b=badge.trim();
  if(!b){toast('❌ Badge cannot be empty');return}
  try{await zkAPI('/api/db/resolve-unknown',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({device_ip:ip, uid:uid, badge:b})});toast('✅ Mapped UID '+uid+' → Badge '+b);loadUnmapped()}catch(e){toast('❌ '+e.message)}
}
async function loadShifts(){
  if(STATE.isDemo){el('shiftsEditor').innerHTML='<div class="text-muted">Demo — connect ZK server for shift editing</div>';return}
  try{
    const d=await zkAPI('/api/shifts');
    const depts=Object.keys(d||{});
    el('shiftsEditor').innerHTML=`<div style="display:grid;grid-template-columns:1fr 100px 100px 80px;gap:8px;align-items:center;padding:8px 0;border-bottom:2px solid var(--border);font-size:10px;font-weight:700;text-transform:uppercase;color:var(--text2)"><div>Department</div><div>Start</div><div>End</div><div>Grace (min)</div></div>`+
    depts.map(dept=>`<div class="shift-row">
      <div class="shift-dept">${esc(dept)}</div>
      <input class="form-control" type="time" id="sh_start_${esc(dept)}" value="${d[dept].start||'07:30'}">
      <input class="form-control" type="time" id="sh_end_${esc(dept)}" value="${d[dept].end||'15:00'}">
      <input class="form-control" type="number" id="sh_grace_${esc(dept)}" value="${d[dept].grace||15}" min="0" max="120">
    </div>`).join('');
    el('shiftsEditor')._depts=depts;
  }catch(e){el('shiftsEditor').innerHTML='<div class="text-red">'+esc(e.message)+'</div>'}
}
async function saveShifts(){
  const depts=el('shiftsEditor')._depts||[];
  const shifts=depts.map(dept=>({dept,start:el('sh_start_'+dept)?.value||'07:30',end:el('sh_end_'+dept)?.value||'15:00',grace_mins:parseInt(el('sh_grace_'+dept)?.value||15)}));
  try{await zkAPI('/api/shifts',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({shifts})});toast('✅ Shifts saved')}catch(e){toast('❌ '+e.message)}
}
async function loadWorkdays(){
  if(STATE.isDemo){el('workdaysEditor').innerHTML='<div class="text-muted">Demo mode</div>';return}
  try{
    const d=await zkAPI('/api/workdays/dept');
    const days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    const depts=Object.keys(d||{});
    el('workdaysEditor').innerHTML=`<div style="display:grid;grid-template-columns:1fr repeat(7,50px);gap:6px;align-items:center;padding:8px 0;border-bottom:2px solid var(--border);font-size:10px;font-weight:700;color:var(--text2)"><div>Department</div>${days.map(d=>`<div style="text-align:center">${d}</div>`).join('')}</div>`+
    depts.map(dept=>{
      const wd=d[dept]||[0,1,2,3,6];
      return `<div style="display:grid;grid-template-columns:1fr repeat(7,50px);gap:6px;align-items:center;padding:6px 0;border-bottom:1px solid var(--border)">
        <div style="font-size:12px;font-weight:600">${esc(dept)}</div>
        ${[0,1,2,3,4,5,6].map(dnum=>`<div style="text-align:center"><input type="checkbox" id="wd_${esc(dept)}_${dnum}" ${wd.includes(dnum)?'checked':''} style="accent-color:var(--accent);transform:scale(1.2)"></div>`).join('')}
      </div>`;
    }).join('');
    el('workdaysEditor')._depts=depts;
  }catch(e){el('workdaysEditor').innerHTML='<div class="text-red">'+esc(e.message)+'</div>'}
}
async function saveWorkdays(){
  const depts=el('workdaysEditor')._depts||[];
  const workdays={};
  depts.forEach(dept=>{workdays[dept]=[0,1,2,3,4,5,6].filter(d=>el('wd_'+dept+'_'+d)?.checked)});
  try{await zkAPI('/api/workdays/dept',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(workdays)});toast('✅ Workdays saved')}catch(e){toast('❌ '+e.message)}
}
async function loadHolidays(){
  if(STATE.isDemo){el('holidaysList').innerHTML='<div class="text-muted">Demo mode</div>';return}
  try{
    const d=await zkAPI('/api/holidays');
    el('holidaysList').innerHTML=(d||[]).map(h=>`<div class="holiday-item">
      <div class="holiday-dates">${esc(h.date)} → ${esc(h.date_end||h.date)}</div>
      <div class="holiday-label">${esc(h.label)}</div>
      <div class="holiday-scope">${esc(h.scope==='all'?'All employees':'Dept: '+h.dept)}</div>
      <button class="btn btn-danger btn-sm" onclick="deleteHoliday(${h.id})">🗑</button>
    </div>`).join('')||'<div class="text-muted">No holidays configured</div>';
  }catch(e){el('holidaysList').innerHTML='<div class="text-red">'+esc(e.message)+'</div>'}
}
function openAddHoliday(){el('hlFrom').value='';el('hlTo').value='';el('hlLabel').value='';el('hlScope').value='all';el('hlDept').value='';el('hlErr').textContent='';openModal('holidayModal')}
async function submitHoliday(){
  const from=el('hlFrom').value,to=el('hlTo').value,label=el('hlLabel').value.trim(),scope=el('hlScope').value,dept=el('hlDept').value.trim();
  if(!from||!label){el('hlErr').textContent='Date and label required';return}
  try{
    await zkAPI('/api/holidays',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({date:from,date_end:to||from,label,scope,dept})});
    toast('✅ Holiday added');closeModal('holidayModal');loadHolidays();
  }catch(e){el('hlErr').textContent=e.message}
}
async function deleteHoliday(id){
  if(!confirm('Delete this holiday?'))return;
  try{await zkAPI('/api/holidays/'+id,{method:'DELETE'});toast('✅ Deleted');loadHolidays()}catch(e){toast('❌ '+e.message)}
}
async function loadRoles(){
  const canManage=STATE.isAdmin;
  if(STATE.isDemo){
    el('rolesBody').innerHTML=['Admin','Technician','Employee'].map(r=>`<tr><td>${r}</td><td>●</td><td>Built-in</td><td>—</td></tr>`).join('');return;
  }
  try{
    if(CFG.gasUrl){
      const d=await gasAPI({action:'getRoles'});
      if(Array.isArray(d)){
        STATE.roles=d.slice(1).map(r=>({name:r[0],color:r[1],desc:r[2]})).filter(r=>r.name);
        el('rolesBody').innerHTML=STATE.roles.map(r=>`<tr>
          <td>${esc(r.name)}</td>
          <td><span style="color:${r.color||'#6b7280'}">●</span></td>
          <td>${esc(r.desc||'—')}</td>
          <td>${canManage&&!['Admin','Technician','User'].includes(r.name)?`<button class="btn btn-danger btn-sm" onclick="deleteRole('${esc(r.name)}')">🗑</button>`:'—'}</td>
        </tr>`).join('');
      }
    }
  }catch(e){}
}
function openAddRole(){el('rlName').value='';el('rlDesc').value='';el('rlColor').value='#6366f1';el('rlErr').textContent='';openModal('roleModal')}
async function submitRole(){
  const name=el('rlName').value.trim();
  if(!name){el('rlErr').textContent='Name required';return}
  try{
    if(STATE.isDemo){toast('🎮 Demo');closeModal('roleModal');return}
    await gasAPI({action:'addRole',name,color:el('rlColor').value,description:el('rlDesc').value});
    toast('✅ Role created');closeModal('roleModal');loadRoles();
  }catch(e){el('rlErr').textContent=e.message}
}
async function deleteRole(name){
  if(!confirm('Delete role '+name+'?'))return;
  try{await gasAPI({action:'deleteRole',name});toast('✅ Deleted');loadRoles()}catch(e){toast('❌ '+e.message)}
}
async function loadSessions(){
  if(STATE.isDemo){el('sessionsBody').innerHTML='<tr><td colspan="6" class="text-muted" style="padding:12px;text-align:center">Demo session</td></tr>';return}
  try{
    const d=await zkAPI('/api/sessions');
    el('sessionsBody').innerHTML=(d||[]).map(s=>`<tr>
      <td class="td-mono">${esc(s.username)}</td>
      <td><span class="tag tag-${s.role==='admin'?'admin':'employee'}">${esc(s.role)}</span></td>
      <td class="td-mono text-muted" style="font-size:10px">${esc(s.ip||'—')}</td>
      <td class="td-mono text-muted" style="font-size:10px">${fmtDateTime(s.login_time)}</td>
      <td class="td-mono text-muted" style="font-size:10px">${fmtDateTime(s.last_active)}</td>
      <td><button class="btn btn-danger btn-sm" onclick="killSession('${esc(s.sid)}','${esc(s.username)}')">Logout</button></td>
    </tr>`).join('')||'<tr><td colspan="6" class="text-muted" style="padding:12px;text-align:center">No active sessions</td></tr>';
  }catch(e){el('sessionsBody').innerHTML='<tr><td colspan="6" class="text-red">'+esc(e.message)+'</td></tr>'}
}
async function killSession(sid,username){
  if(!confirm('Force logout '+username+'?'))return;
  try{await zkAPI('/api/sessions/user/'+username,{method:'POST'});toast('✅ Session terminated');loadSessions()}catch(e){toast('❌ '+e.message)}
}

