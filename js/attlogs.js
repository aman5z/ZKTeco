async function loadAttLogs(){
  const f=el('logsFrom')?el('logsFrom').value:'';
  const t=el('logsTo')?el('logsTo').value:'';
  const b=el('logsBadge')?el('logsBadge').value:'';
  if(!f || !t) return toast('⚠ Select start and end dates');
  el('logsBody').innerHTML='<tr><td colspan="5" style="text-align:center"><div class="spinner-border" style="width:20px;height:20px;border-width:2px;border-color:var(--accent) transparent var(--accent) transparent;border-radius:50%;animation:spin 1s linear infinite;"></div> Loading...</td></tr>';
  try{
    const r=await fetch(CFG.zkUrl+'/api/punches?from='+f+'&to='+t+'&badge='+encodeURIComponent(b), {credentials:'include'});
    if(r.ok){
      const d=await r.json();
      if(!d || !d.length){
        el('logsBody').innerHTML='<tr><td colspan="5" class="empty-state" style="padding:20px;text-align:center">No punch logs found</td></tr>';
        if(el('logsCount'))el('logsCount').textContent='0 records';
        return;
      }
      if(el('logsCount'))el('logsCount').textContent=d.length+' records';
      el('logsBody').innerHTML=d.map(r=>`<tr>
        <td class="td-mono">${r.timestamp||r.punch_time||''}</td>
        <td class="td-mono">${esc(r.badge||'')}</td>
        <td>${esc(r.name||'')}</td>
        <td>${esc(r.dept||'')}</td>
        <td class="td-mono text-muted">${esc(r.device_ip||'')}</td>
      </tr>`).join('');
    }else{
      el('logsBody').innerHTML='<tr><td colspan="5" class="text-red">Error loading logs</td></tr>';
    }
  }catch(e){
    el('logsBody').innerHTML='<tr><td colspan="5" class="text-red">Error: '+esc(e.message)+'</td></tr>';
  }
}

