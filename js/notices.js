// ===========================================================================
//  NOTICE BOARD LOGIC
// ===========================================================================
async function loadAnnouncements(){
  try{
    const d=await zkAPI('/api/announcements');
    if(d&&d.length>0){
      el('announcementBanner').style.display='block';
      let html = d.map(a => `[${a.created_at.split(' ')[0]}] ${esc(a.message)}`).join(' &nbsp;|&nbsp; ');
      el('announcementText').innerHTML = html;
    }else{
      el('announcementBanner').style.display='none';
    }
  }catch(e){}
}
async function loadAdminAnnouncements(){
  try{
    const d=await zkAPI('/api/announcements');
    if(!d||d.error)throw new Error(d.error||'Load fail');
    el('adminAnnBody').innerHTML = d.map(a=>`<tr>
      <td>${esc(a.message)}</td><td class="text-mono">${esc(a.created_at)}</td><td class="text-mono">${esc(a.created_by)}</td>
      <td><button class="btn btn-danger btn-sm" onclick="delAnnouncement(${a.id})">Remove</button></td>
    </tr>`).join('')||'<tr><td colspan="4" class="text-muted text-center">No active announcements</td></tr>';
  }catch(e){toast('Error loading notices: '+e.message)}
}
async function postAnnouncement(){
  const msg=el('newAnnText').value.trim();
  if(!msg)return;
  try{
    const d=await zkAPI('/api/announcements',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({message:msg})});
    if(d.error)throw new Error(d.error);
    el('newAnnText').value='';
    toast('✅ Announcement posted');
    loadAdminAnnouncements();
  }catch(e){toast('❌ '+e.message)}
}
async function delAnnouncement(id){
  if(!confirm('Remove this announcement?'))return;
  try{
    const d=await zkAPI('/api/announcements/'+id,{method:'DELETE'});
    if(d.error)throw new Error(d.error);
    toast('✅ Removed');
    loadAdminAnnouncements();
  }catch(e){toast('❌ '+e.message)}
}

// Initialize Announcements on start
setTimeout(loadAnnouncements, 1000);
