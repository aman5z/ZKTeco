// ===========================================================================
//  REMOTE PUNCH
// ===========================================================================
async function loadPunchStatus(){
  if(STATE.isDemo){
    el('myPunchList').innerHTML='<div class="punch-entry"><span>07:32:14</span><span class="tag tag-present">confirmed</span></div>';
    el('selfPunchBtn').style.display=STATE.user.badge&&STATE.user.badge!=='DEMO'?'':'none';
    el('adminPunchSection').style.display=STATE.isAdmin?'':'none';
    return;
  }
  try{
    const d=await zkAPI('/api/punch/status');
    el('myPunchList').innerHTML=(d.punches||[]).map(p=>`<div class="punch-entry">
      <span class="text-mono">${(p.time||'').substring(11,19)||'—'}</span>
      <span class="tag tag-${p.status==='confirmed'?'present':p.status==='pending'?'pending':p.status==='device_ok'?'approved':'off'}">${p.status}</span>
      <span class="text-muted" style="font-size:10px">${p.source||'—'}</span>
    </div>`).join('')||'<div class="text-muted" style="font-size:11px;text-align:center">No punches today</div>';
  }catch(e){el('selfPunchBtn').style.display='none'}
  el('adminPunchSection').style.display=STATE.isAdmin?'':'none';
  if(STATE.isAdmin)loadApprovals();
}
async function saveOfflinePunch(url, method, headers, body) {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('ERP_OfflineDB', 1);
    req.onupgradeneeded = e => {
      e.target.result.createObjectStore('offline_punches', { keyPath: 'id', autoIncrement: true });
    };
    req.onsuccess = e => {
      const db = e.target.result;
      const tx = db.transaction('offline_punches', 'readwrite');
      tx.objectStore('offline_punches').add({ url, method, headers, body, ts: Date.now() });
      tx.oncomplete = () => resolve();
    };
    req.onerror = reject;
  });
}

async function doSelfPunch(){
  if(STATE.isDemo){toast('🎮 Demo');return}
  const url='/api/punch/self', method='POST', headers={'Content-Type':'application/json'}, body='{}';
  try{
    const d=await zkAPI(url,{method,headers,body});
    toast('✅ '+(d.message||'Punched in'));loadPunchStatus();
  }catch(e){
    if(e.message==='Failed to fetch' || e.message.includes('fetch') || !navigator.onLine){
      await saveOfflinePunch(CFG.zkUrl+url, method, headers, body);
      if('serviceWorker' in navigator && 'SyncManager' in window){
        try{
          const reg=await navigator.serviceWorker.ready;
          await reg.sync.register('sync-punches');
          toast('📴 Offline: Punch saved. Will sync when online.');
        }catch(swE){toast('📴 Offline: Punch saved locally.');}
      }else{toast('📴 Offline: Punch saved locally.');}
    }else{toast('❌ '+e.message)}
  }
}
async function submitPunchRequest(){
  const dt=el('reqPunchTime').value,reason=el('reqPunchReason').value.trim();
  if(!dt){toast('Select date/time');return}
  if(!reason){toast('Enter reason');return}
  if(STATE.isDemo){toast('🎮 Demo');return}
  try{
    const punchTime=dt.replace('T',' ')+':00';
    const d=await zkAPI('/api/punch/request',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({punch_time:punchTime,reason})});
    el('punchReqMsg').textContent='✅ '+(d.message||'Request submitted');
    el('punchReqMsg').style.color='var(--green)';
    el('reqPunchTime').value='';el('reqPunchReason').value='';
  }catch(e){el('punchReqMsg').textContent='❌ '+e.message;el('punchReqMsg').style.color='var(--red)'}
}
async function loadApprovals(){
  if(STATE.isDemo){el('approvalsList').innerHTML='<div class="text-muted" style="font-size:12px">No pending requests (demo).</div>';return}
  try{
    const rows=await zkAPI('/api/punch/admin/pending');
    el('approvalBadge').textContent=rows.length;
    el('approvalBadge').style.display=rows.length?'':'none';
    el('approvalsList').innerHTML=rows.length?rows.map(r=>`<div class="approval-row">
      <div class="emp-name">${esc(r.employee_name||r.badge)} <span class="text-muted text-mono" style="font-size:10px">#${r.badge}</span></div>
      <div class="punch-time">${esc(r.punch_time)}</div>
      <div class="reason">${esc(r.reason||'')}</div>
      <div class="approval-btns">
        <button class="btn btn-green btn-sm" onclick="doApprove(${r.id},'approve')">✓ Approve</button>
        <button class="btn btn-danger btn-sm" onclick="doApprove(${r.id},'reject')">✗ Reject</button>
      </div>
    </div>`).join(''):'<div class="text-muted" style="font-size:12px">No pending requests.</div>';
  }catch(e){el('approvalsList').innerHTML='<div class="text-red" style="font-size:12px">'+esc(e.message)+'</div>'}
}
async function doApprove(id,action){
  try{
    const d=await zkAPI('/api/punch/admin/approve/'+id,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action})});
    toast('✅ '+(d.message||action));loadApprovals();if(action==='approve')loadDeviceQueue();
  }catch(e){toast('❌ '+e.message)}
}
async function loadDeviceQueue(){
  if(STATE.isDemo){el('deviceQueueList').innerHTML='<div class="text-muted" style="font-size:12px">All clear (demo).</div>';return}
  try{
    const rows=await zkAPI('/api/punch/admin/device-queue');
    el('deviceQueueList').innerHTML=rows.length?rows.map(r=>`<div style="display:flex;align-items:center;gap:8px;padding:6px;background:var(--card2);border-radius:4px;margin-bottom:4px;font-size:11px">
      <span class="tag ${r.device_status==='success'?'tag-approved':r.device_status==='failed'?'tag-rejected':'tag-pending'}">${r.device_status}</span>
      <span class="fw7">${esc(r.employee_name||r.badge)}</span>
      <span class="text-mono text-muted">${r.punch_time?.substring(0,16)||'—'}</span>
      ${r.device_error?`<span class="text-red" style="margin-left:auto;font-size:10px">⚠ ${esc(r.device_error.substring(0,40))}</span>`:''}
    </div>`).join(''):'<div class="text-muted" style="font-size:12px">✓ All punches pushed to device.</div>';
  }catch(e){el('deviceQueueList').innerHTML='<div class="text-red" style="font-size:12px">'+esc(e.message)+'</div>'}
}
async function retryDevicePush(){
  if(STATE.isDemo){toast('🎮 Demo');return}
  try{const d=await zkAPI('/api/punch/admin/retry-device',{method:'POST'});toast(d.message||'Retrying...');setTimeout(loadDeviceQueue,2500)}catch(e){toast('❌ '+e.message)}
}
async function submitAdminPunch(){
  const badge=el('adminPunchBadge').value.trim(),dt=el('adminPunchTime').value,reason=el('adminPunchReason').value.trim();
  if(!badge){toast('Enter badge');return}if(!dt){toast('Select date/time');return}
  if(STATE.isDemo){toast('🎮 Demo');return}
  const punchTime=dt.replace('T',' ')+':00';
  try{
    const d=await zkAPI('/api/punch/admin/direct',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({badge,punch_time:punchTime,reason:reason||'Admin manual entry'})});
    el('adminPunchMsg').textContent='✅ '+(d.message||'Punch added');el('adminPunchMsg').style.color='var(--green)';loadDeviceQueue();
  }catch(e){el('adminPunchMsg').textContent='❌ '+e.message;el('adminPunchMsg').style.color='var(--red)'}
}
function lookupEmpName(){
  const badge=el('adminPunchBadge').value.trim();
  const emp=STATE.empList.find(e=>(e.code||e.badge)===badge);
  el('adminEmpNameHint').textContent=emp?(emp.name+' · '+emp.dept):'';
}
function _loadPunchBadgeSuggestions(){
  const dl=el('punchBadgeSuggestions');if(!dl)return;
  dl.innerHTML=STATE.empList.map(e=>`<option value="${esc(e.code||e.badge)}" label="${esc(e.name)}"></option>`).join('');
}
