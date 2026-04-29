async function loadTickets(){
  spin(true);
  try{
    if(STATE.isDemo){STATE.tickets=DEMO_TICKETS;}
    else{
      const d = await zkAPI('/api/tickets');
      STATE.tickets = (d.tickets||[]).map(r=>({
        id:r.id||'',title:r.title||'',priority:r.priority||'Low',status:r.status||'Open',
        requesterName:r.requester_badge||'',assignedTo:r.assigned_to||'Unassigned',
        dueTime:r.due_time||'',desc:r.desc||'',created:r.created_at||''
      }));
    }
    el('ticketCount').textContent=STATE.tickets.length+' tickets';
    filterTickets();
  }catch(e){toast('❌ Tickets: '+e.message);console.error('Tickets error:',e);}
  spin(false);
}
function filterTickets(){
  const q=(el('tickSearch').value||'').toLowerCase();
  const sf=el('tickStatusFilter').value,pf=el('tickPriFilter').value;
  let t=STATE.tickets;
  if(q)t=t.filter(x=>(x.title||'').toLowerCase().includes(q)||(x.requesterName||'').toLowerCase().includes(q));
  if(sf)t=t.filter(x=>x.status===sf);
  if(pf)t=t.filter(x=>x.priority===pf);
  const now=new Date();
  el('ticketBody').innerHTML=t.map(tk=>{
    const due=new Date(tk.dueTime),overdue=now>due&&tk.status==='Open';
    const sb=tk.status==='Closed'?'tag-closed':tk.status==='Resolved'?'tag-resolved':overdue?'tag-overdue':'tag-open';
    const pb=tk.priority==='High'?'tag-high':tk.priority==='Medium'?'tag-medium':'tag-low';
    return `<tr onclick="selectTicket('${esc(tk.id)}')">
      <td class="td-mono" style="font-size:10px">${esc(tk.id)}</td>
      <td>${esc(tk.title)}</td>
      <td>${esc(tk.requesterName||'—')}</td>
      <td><span class="tag ${pb}">${esc(tk.priority)}</span></td>
      <td><span class="tag ${sb}">${esc(overdue?'Overdue':tk.status)}</span></td>
      <td>${esc(tk.assignedTo||'—')}</td>
      <td class="td-mono text-muted" style="font-size:10px">${tk.dueTime?fmtDateTime(tk.dueTime):'—'}</td>
    </tr>`;
  }).join('')||'<tr><td colspan="7" class="empty-state" style="padding:20px;text-align:center">No tickets found</td></tr>';
}
function selectTicket(id){
  STATE.selectedTicket=id;
  const tk=STATE.tickets.find(t=>t.id==id);if(!tk)return;
  const now=new Date(),due=new Date(tk.dueTime),overdue=now>due&&tk.status==='Open';
  const sb=tk.status==='Closed'?'tag-closed':tk.status==='Resolved'?'tag-resolved':overdue?'tag-overdue':'tag-open';
  const pb=tk.priority==='High'?'tag-high':tk.priority==='Medium'?'tag-medium':'tag-low';
  el('ticketDetailPane').innerHTML=`
    <div class="dp-header">Ticket Details</div>
    <div class="dp-avatar">
      <div style="font-size:32px;margin-bottom:8px">🎫</div>
      <div class="dp-name text-mono" style="font-size:11px">${esc(tk.id)}</div>
      <div class="dp-sub" style="max-width:200px;word-wrap:break-word;white-space:normal">${esc(tk.title)}</div>
    </div>
    <div class="dp-row"><label>Priority</label><span class="tag ${pb}">${esc(tk.priority)}</span></div>
    <div class="dp-row"><label>Status</label><span class="tag ${sb}">${esc(overdue?'Overdue':tk.status)}</span></div>
    <div class="dp-row"><label>Requester</label><span>${esc(tk.requesterName||'—')}</span></div>
    <div class="dp-row"><label>Assigned</label><span>${esc(tk.assignedTo||'—')}</span></div>
    <div class="dp-row"><label>Due</label><span class="text-mono">${fmtDateTime(tk.dueTime)}</span></div>
    <div class="dp-row"><label>Description</label><span style="white-space:pre-wrap;font-size:10px;line-height:1.5">${esc(tk.desc||'—')}</span></div>
    ${STATE.isAdmin||can('tickets.manage')?`<div class="dp-actions">
      <button class="dp-btn" onclick="openUpdateTicket('${esc(tk.id)}','${esc(tk.status)}','${esc(tk.assignedTo)}')">✏ Update</button>
      <button class="dp-btn danger" onclick="closeTicket('${esc(tk.id)}')">🔒 Close Ticket</button>
    </div>`:''}
  `;
}
function openNewTicket(){
  ['tkName','tkTitle','tkDesc'].forEach(id=>el(id).value='');
  el('tkPri').value='Medium';el('tkErr').textContent='';
  openModal('ticketModal');
}
async function submitTicket(){
  const name=el('tkName').value.trim(),title=el('tkTitle').value.trim(),desc=el('tkDesc').value.trim(),pri=el('tkPri').value;
  if(!name||!title){el('tkErr').textContent='Name and title required';return}
  spin(true);
  try{
    if(STATE.isDemo){
      const tid='DEMO-'+Date.now();
      STATE.tickets.unshift({id:tid,title,priority:pri,status:'Open',requesterName:name,assignedTo:'Unassigned',dueTime:new Date(Date.now()+(pri==='High'?4:pri==='Medium'?12:24)*3600000).toISOString(),desc});
      toast('✅ Ticket created (demo)');closeModal('ticketModal');filterTickets();spin(false);return;
    }
    await zkAPI('/api/tickets', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({title, desc, priority: pri})
    });
    toast('✅ Ticket submitted');closeModal('ticketModal');loadTickets();
  }catch(e){el('tkErr').textContent=e.message}
  spin(false);
}
function openUpdateTicket(id){
  const tk=STATE.tickets.find(x=>x.id==id);if(!tk)return;
  el('utId').value=tk.id;el('utTitle').value=tk.title;el('utDesc').value=tk.desc||'';
  el('utStatus').value=tk.status;el('utPriority').value=tk.priority;
  el('utAssigned').value=tk.assignedTo==='Unassigned'?'':tk.assignedTo;
  el('utErr').textContent='';openModal('updateTicketModal');
}
async function submitUpdateTicket(){
  const id=el('utId').value,status=el('utStatus').value,assignedTo=el('utAssigned').value.trim()||'Unassigned',title=el('utTitle').value.trim(),desc=el('utDesc').value.trim(),priority=el('utPriority').value;
  spin(true);
  try{
    if(STATE.isDemo){const t=STATE.tickets.find(x=>x.id==id);if(t){t.status=status;t.assignedTo=assignedTo;t.title=title;t.desc=desc;t.priority=priority;}toast('✅ Updated (demo)');closeModal('updateTicketModal');filterTickets();selectTicket(id);spin(false);return}
    await zkAPI('/api/tickets/'+id, {
      method: 'PUT',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({title, desc, status, priority, assigned_to: assignedTo})
    });
    toast('✅ Ticket updated');closeModal('updateTicketModal');loadTickets();
  }catch(e){el('utErr').textContent=e.message}
  spin(false);
}
async function closeTicket(id){
  if(!confirm('Close ticket '+id+'?'))return;
  spin(true);
  try{
    if(STATE.isDemo){const t=STATE.tickets.find(x=>x.id==id);if(t)t.status='Closed';toast('🔒 Closed (demo)');filterTickets();selectTicket(id);spin(false);return}
    await zkAPI('/api/tickets/'+id, {
      method: 'PUT',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({status: 'Closed', assigned_to: STATE.user.username})
    });
    toast('🔒 Ticket closed');loadTickets();
  }catch(e){toast('❌ '+e.message)}
  spin(false);
}

