async function loadHistory(){
  const from=el('histFrom').value,to=el('histTo').value;
  if(!from||!to){toast('Select date range');return}
  // Warn if range > 14 days (MDB queries can be slow)
  const dayDiff=(new Date(to)-new Date(from))/(86400000);
  if(dayDiff>14){
    if(!confirm('Range is '+Math.round(dayDiff)+' days. Large ranges may take a while or timeout. Continue?'))return;
  }
  spin(true);
  // Show progress message for long queries
  el('histResults').innerHTML='<div class="empty-state"><div class="icon">⏳</div><p>Loading history from database...<br><small style="color:var(--text2)">Large ranges may take 30-60 seconds</small></p></div>';
  try{
    if(STATE.isDemo){
      el('histResults').innerHTML=`<div class="card"><div class="card-header"><h3>Demo History Report</h3></div><div class="card-body"><p class="text-muted" style="font-size:12px">Connect to ZKTeco server for real history data.</p></div></div>`;
      spin(false);return;
    }
    const fromFmt=from.split('-').reverse().join('/'),toFmt=to.split('-').reverse().join('/');
    // Use AbortController with 90s timeout
    const ctrl=new AbortController();
    const tid=setTimeout(()=>ctrl.abort(),90000);
    let d;
    try{
      const r=await fetch(CFG.zkUrl+'/api/history?from='+fromFmt+'&to='+toFmt,{credentials:'include',signal:ctrl.signal});
      clearTimeout(tid);
      if(!r.ok){
        const err=await r.json().catch(()=>({error:'Server error '+r.status}));
        throw new Error(err.error||'Server error '+r.status);
      }
      d=await r.json();
    }catch(fetchErr){
      if(fetchErr.name==='AbortError'){
        el('histResults').innerHTML='<div class="empty-state text-red"><div class="icon">⏱</div><p>Request timed out (90s). Try a smaller date range — 7 days max works best.</p></div>';
        spin(false);return;
      }
      throw fetchErr;
    }
    const days=d.days||[];
    el('histResults').innerHTML=days.map(day=>`
      <div class="card" style="margin-bottom:12px">
        <div class="card-header">
          <h3>📅 ${esc(day.date)}</h3>
          <div style="display:flex;gap:10px;font-size:12px;font-family:var(--mono)">
            <span class="text-green">✅ ${day.present_count} present</span>
            <span class="text-red">❌ ${day.absent_count} absent</span>
          </div>
        </div>
        <div class="card-body" style="padding:10px">
          <details><summary style="cursor:pointer;font-size:12px;color:var(--text2)">Show absent (${day.absent_count})</summary>
          <div style="margin-top:8px;display:flex;flex-wrap:wrap;gap:4px">${(day.absent||[]).map(e=>`<span class="tag tag-absent">${esc(e.name)}</span>`).join('')||'—'}</div></details>
        </div>
      </div>`).join('')||'<div class="empty-state"><div class="icon">📊</div><p>No data for selected range</p></div>';
  }catch(e){el('histResults').innerHTML=`<div class="empty-state text-red"><div class="icon">⚠</div><p>${esc(e.message)}</p></div>`}
  spin(false);
}
async function exportHistory(){
  if(STATE.isDemo){toast('🎮 Demo');return}
  const from=el('histFrom').value,to=el('histTo').value;
  if(!from||!to){toast('Select date range');return}
  const fromFmt=from.split('-').reverse().join('/'),toFmt=to.split('-').reverse().join('/');
  window.open(CFG.zkUrl+'/api/history/export?from='+fromFmt+'&to='+toFmt,'_blank');
}

