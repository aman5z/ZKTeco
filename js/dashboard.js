async function loadDashboard(){
  el('dashDate').textContent=new Date().toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
  setCacheDot('loading');
  try{
    if(STATE.isDemo){
      renderDashboardData({present:DEMO_PRESENT,absent:DEMO_ABSENT,off:[],total:DEMO_EMPLOYEES.length,devices:DEMO_DEVICES});
    }else{
      const d=await zkAPI('/api/today');
      STATE.todayData=d;
      STATE.todayLoaded=true;
      // Fetch live device status for dashboard card
      if(!d.devices||!d.devices.length){try{d.devices=await zkAPI('/api/devices/status');}catch(e){}}
      renderDashboardData({
        present:d.present||[],absent:d.absent||[],off:d.off_today||[],
        presentCount:d.present_count,absentCount:d.absent_count,offCount:d.off_today_count,
        total:d.total||(d.present_count||0)+(d.absent_count||0)+(d.off_today_count||0),
        devices:d.devices||d.device_status||d.device_statuses||[]
      });
      setCacheDot('fresh',d.last_updated);
    }
  }catch(e){setCacheDot('error');console.warn('Dashboard load:',e)}
  // Also load GAS dashboard (non-blocking)
  if(CFG.gasUrl){
    el('dashTickets').innerHTML='<div class="empty-state" style="padding:20px;text-align:center;font-size:12px;color:var(--text2)"><div class="icon" style="font-size:20px;animation:spin 1s linear infinite">⏳</div><p>Loading Helpdesk...</p></div>';
    gasAPI({action:'getDashboard'}).then(d=>{
      if(d&&d.tickets)renderDashTickets(d.tickets,d.recentAudit);
      else renderDashTickets({total:0,open:0,closed:0,resolved:0,overdue:0},[]);
    }).catch(()=>{
      renderDashTickets({total:0,open:0,closed:0,resolved:0,overdue:0},[]);
    });
  }else if(STATE.isDemo){renderDashTickets({total:3,open:2,closed:1,resolved:1,overdue:1},[]);
  }else{
    el('dashTickets').innerHTML='<div class="empty-state" style="padding:20px;text-align:center;font-size:12px;color:var(--text2)">Configure GAS URL in Settings to enable Helpdesk.</div>';
  }
}

function renderDashboardData({present,absent,off,presentCount,absentCount,offCount,total,devices}){
  const pCnt = presentCount!==undefined ? presentCount : (Array.isArray(present)?present.length:present);
  const aCnt = absentCount!==undefined  ? absentCount  : (Array.isArray(absent)?absent.length:absent);
  const oCnt = offCount!==undefined     ? offCount     : (Array.isArray(off)?off.length:off);
  el('dsPresentCount').textContent=pCnt;
  el('dsAbsentCount').textContent=aCnt;
  el('dsOffCount').textContent=oCnt;
  el('dsTotalCount').textContent=total||'—';
  const sbBadge=el('sbAbsentBadge');
  sbBadge.textContent=aCnt;sbBadge.style.display=aCnt>0?'':'none';
  el('sbEmpCount').textContent=total||'—';
  // Devices
  const dOnline=(devices||[]).filter(d=>d.online).length;
  el('deviceOnlineCount').textContent=dOnline+'/'+(devices||[]).length+' online';
  el('zkStatusDot').className='sb-dot'+(dOnline>0?'':' red');
  el('zkStatusText').textContent='ZK: '+dOnline+'/'+(devices||[]).length;
  el('dashDevices').innerHTML=(devices||[]).map(d=>`
    <div style="display:flex;align-items:center;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--border);font-size:12px">
      <div><span class="tag ${d.online?'tag-online':'tag-offline'}">${d.online?'●':'○'}</span> <strong>${esc(d.name||d.ip)}</strong> <span class="text-mono text-muted" style="font-size:10px">${d.ip}</span></div>
      <div class="text-mono" style="font-size:11px;color:var(--accent)">${d.punches_today||0} punches</div>
    </div>`).join('')||'<div class="text-muted" style="font-size:12px">No device data</div>';
  // Department Attendance %
  if(present && absent){
    let dStats = {};
    const allEmps = [...(Array.isArray(present)?present:[]), ...(Array.isArray(absent)?absent:[]), ...(Array.isArray(off)?off:[])];
    allEmps.forEach(e => {
        let d = e.dept || 'Unknown';
        if(!dStats[d]) dStats[d] = { p:0, t:0 };
        dStats[d].t++;
    });
    (Array.isArray(present)?present:[]).forEach(e => {
        let d = e.dept || 'Unknown';
        if(dStats[d]) dStats[d].p++;
    });
    const zeroDepts = [];
    const html = Object.entries(dStats).sort().map(([d,v]) => {
        let pct = v.t > 0 ? Math.round((v.p/v.t)*100) : 0;
        const noShow = v.t > 0 && v.p === 0;
        if(noShow) zeroDepts.push(d);
        return `<div style="margin-bottom:10px${noShow?';background:rgba(239,68,68,.08);border-radius:6px;padding:4px 6px;border-left:3px solid var(--red)':''}">
          <div style="display:flex;justify-content:space-between;font-size:11px;margin-bottom:4px">
            <strong>${esc(d)}${noShow?' <span style="color:var(--red);font-size:10px">⚠ 0 punches</span>':''}</strong><span class="text-mono${noShow?' text-red':''}">${pct}% (${v.p}/${v.t})</span>
          </div>
          <div style="height:6px;background:var(--border);border-radius:3px;overflow:hidden">
            <div style="height:100%;background:${noShow?'var(--red)':'var(--accent)'};width:${pct||100}%;border-radius:3px;opacity:${noShow?'.35':'1'}"></div>
          </div>
        </div>`;
    }).join('');
    if(zeroDepts.length){
      const warn=el('dashDeptWarn');
      if(warn){warn.textContent='⚠ No punches today: '+zeroDepts.join(', ');warn.style.display='';}
    }else{
      const warn=el('dashDeptWarn');
      if(warn) warn.style.display='none';
    }
    if(el('dashDeptAtt')) el('dashDeptAtt').innerHTML = html || '<div class="text-muted" style="font-size:12px;text-align:center;padding:10px;">No department data.</div>';
  }

  // Audit
  const auditRows=(STATE.allAudit||[]).slice(0,8);
  if(auditRows.length){
    el('dashAudit').innerHTML=auditRows.map(a=>`<tr>
      <td class="td-mono" style="font-size:10px">${fmtDateTime(a.ts||a.time)}</td>
      <td>${esc(a.username||a.actor||'—')}</td>
      <td><span class="atag atag-${(a.action||'').replace(/[^A-Z_]/g,'')}">${esc(a.action||'—')}</span></td>
      <td class="text-muted" style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(a.detail||'')}</td>
    </tr>`).join('');
  }
}

function renderDashTickets(ts,audit){
  el('dashTickets').innerHTML=`
    <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:10px">
      <div class="stat-card" style="padding:12px" data-color="orange"><div class="sc-label">Open</div><div class="sc-value" style="font-size:24px;color:var(--orange)">${ts.open||0}</div></div>
      <div class="stat-card" style="padding:12px" data-color="red"><div class="sc-label">Overdue</div><div class="sc-value" style="font-size:24px;color:var(--red)">${ts.overdue||0}</div></div>
      <div class="stat-card" style="padding:12px" data-color="green"><div class="sc-label">Resolved</div><div class="sc-value" style="font-size:24px;color:var(--green)">${ts.resolved||0}</div></div>
      <div class="stat-card" style="padding:12px"><div class="sc-label">Total</div><div class="sc-value" style="font-size:24px">${ts.total||0}</div></div>
    </div>
    <div style="margin-top:10px"><button class="btn btn-secondary btn-sm" onclick="showPage('tickets')">View Helpdesk →</button></div>
  `;
  if(ts.open>0){el('sbTicketBadge').textContent=ts.open;el('sbTicketBadge').style.display='';}
  if(audit&&audit.length){
    STATE.allAudit=[...audit,...STATE.allAudit].slice(0,100);
    const auditRows=audit.slice(0,8);
    el('dashAudit').innerHTML=auditRows.map(a=>`<tr>
      <td class="td-mono" style="font-size:10px">${fmtDateTime(a.time)}</td>
      <td>${esc(a.actor||'—')}</td>
      <td><span class="atag atag-${(a.action||'').replace(/[^A-Z_]/g,'')}">${esc(a.action||'—')}</span></td>
      <td class="text-muted">${esc(a.detail||'')}</td>
    </tr>`).join('');
  }
}

function setCacheDot(state,ts){
  const d=el('cacheDot');
  d.className='cache-dot '+state;
  // ts can be "HH:MM:SS", ISO string, or Date — handle all
  let label='';
  if(ts){
    if(typeof ts==='string'&&/^\d{2}:\d{2}(:\d{2})?$/.test(ts)){
      label='Updated '+ts;
    }else{
      const dt=new Date(ts);
      label='Updated '+(isNaN(dt)?ts:dt.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit',second:'2-digit'}));
    }
  }
  el('cacheTime').textContent=label;
}

async function forceRefresh(){
  if(typeof _gasCache !== 'undefined') _gasCache.clear();
  if(STATE.isDemo){toast('🎮 Demo mode — no real data to refresh');return}
  try{
    await zkAPI('/api/cache/refresh',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
    toast('↻ Refresh triggered');
    setTimeout(loadDashboard,1500);
  }catch(e){toast('❌ Refresh failed: '+e.message)}
}

