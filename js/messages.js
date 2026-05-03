// ===========================================================================
//  MESSAGING & NOTES
// ===========================================================================
let ALL_MESSAGES = [];
let ACTIVE_CHAT_PEER = null;
let ONLINE_USERS = [];

// Start polling when entering messages page, stop when leaving

function getEmpName(b) {
   if(!STATE.empList) return b;
   const e = STATE.empList.find(x => (x.code||x.badge||'') === b);
   return e ? e.name : b;
}

async function loadMessages(){
  try{
    const [msgs, onlines] = await Promise.all([
       zkAPI('/api/messages'),
       zkAPI('/api/users/online').catch(()=>[])
    ]);
    if(msgs.error) throw new Error(msgs.error);
    ALL_MESSAGES = Array.isArray(msgs) ? msgs : [];
    ONLINE_USERS = Array.isArray(onlines) ? onlines : [];

    // Current user identity — server stores messages by username
    const me = STATE.user.username || STATE.user.badge;

    // Group by peer (the other party in the conversation)
    const peers = {};
    ALL_MESSAGES.forEach(m => {
       const peer = m.sender === me ? m.receiver : m.sender;
       if(!peers[peer]) peers[peer] = [];
       peers[peer].push(m);
    });

    let unreadTotal = 0;

    const chatListHTML = Object.keys(peers).map(peer => {
       const pMsgs = peers[peer];
       pMsgs.sort((a,b)=>new Date(a.timestamp)-new Date(b.timestamp));
       const last = pMsgs[pMsgs.length-1];
       const unreadCounts = pMsgs.filter(m => m.receiver === me && !m.is_read).length;
       unreadTotal += unreadCounts;
       const isOnline = ONLINE_USERS.includes(peer);

       return `<div class="chat-list-item ${peer===ACTIVE_CHAT_PEER?'active':''}" onclick="openChat('${esc(peer)}')" style="padding:10px;border-bottom:1px solid var(--border);cursor:pointer;display:flex;justify-content:space-between;align-items:center;border-radius:4px;margin-bottom:4px;background:${peer===ACTIVE_CHAT_PEER?'var(--card2)':'transparent'}">
         <div style="flex:1;overflow:hidden">
           <div style="font-weight:bold;font-size:13px;display:flex;gap:6px;align-items:center">
             ${isOnline?'<span style="width:8px;height:8px;background:var(--green);border-radius:50%;display:inline-block" title="Online"></span>':'<span style="width:8px;height:8px;background:var(--border);border-radius:50%;display:inline-block" title="Offline"></span>'}
             ${esc(getEmpName(peer))}
           </div>
           <div style="font-size:11px;color:var(--text2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(last.message)}</div>
         </div>
         ${unreadCounts>0?`<div style="background:var(--accent);color:#fff;font-size:10px;padding:2px 6px;border-radius:10px">${unreadCounts}</div>`:''}
       </div>`;
    }).join('') || '<div class="text-muted text-center" style="padding:20px;font-size:12px">No conversations yet.<br><br><button class="btn btn-secondary btn-sm" onclick="openNewChatModal()">+ New Message</button></div>';

    el('chatList').innerHTML = chatListHTML;

    if(el('sbMsgBadge')){
      el('sbMsgBadge').textContent=unreadTotal;
      el('sbMsgBadge').style.display=unreadTotal>0?'inline-block':'none';
    }

    if(ACTIVE_CHAT_PEER) openChat(ACTIVE_CHAT_PEER);

  }catch(e){ console.error('[Messages]', e); toast('❌ Messages: '+e.message); }
}

function openChat(peer) {
   ACTIVE_CHAT_PEER = peer;
   document.querySelectorAll('.chat-list-item').forEach(e=>e.style.background='transparent');
   const myItem = [...document.querySelectorAll('.chat-list-item')].find(e=>e.onclick&&e.getAttribute('onclick')&&e.getAttribute('onclick').includes("'"+peer+"'"));
   if(myItem) myItem.style.background = 'var(--card2)';

   el('chatPeerName').textContent = getEmpName(peer);
   el('chatPeerStatus').style.display = 'inline-block';
   const isOnline = ONLINE_USERS.includes(peer);
   el('chatPeerStatus').className = 'tag ' + (isOnline?'tag-present':'tag-off');
   el('chatPeerStatus').textContent = isOnline?'Online':'Offline';
   el('chatInputArea').style.display = 'flex';   // was always hidden due to duplicate display:none

   // Use username as identity — matches what server stores
   const me = STATE.user.username || STATE.user.badge;
   const pMsgs = ALL_MESSAGES.filter(m =>
     (m.sender === peer || m.receiver === peer) &&
     (m.sender === me   || m.receiver === me)
   );
   pMsgs.sort((a,b)=>new Date(a.timestamp)-new Date(b.timestamp));

   el('messagesBody').innerHTML = pMsgs.length ? pMsgs.map(m => {
      const isSent = m.sender === me;
      return `<div style="display:flex;flex-direction:column;align-items:${isSent?'flex-end':'flex-start'};margin-bottom:8px">
         <div style="max-width:75%;padding:8px 12px;border-radius:8px;font-size:13px;position:relative;background:${isSent?'var(--accent)':'var(--card2)'};color:${isSent?'#fff':'var(--text)'};">
           ${esc(m.message)}
           ${isSent ? `<span style="position:absolute;top:-8px;right:-8px;cursor:pointer;font-size:10px;background:var(--red);color:#fff;border-radius:50%;width:16px;height:16px;display:flex;align-items:center;justify-content:center;opacity:0.5" onclick="deleteMessage(${m.id})" onmouseover="this.style.opacity=1" onmouseout="this.style.opacity=0.5" title="Delete">✕</span>` : ''}
         </div>
         <div style="font-size:10px;color:var(--text3);margin-top:4px">${esc(m.timestamp||'')} ${isSent?(m.is_read?'· ✓ Read':''):''}</div>
      </div>`;
   }).join('') : '<div class="empty-state" style="margin:auto;padding:40px;text-align:center;color:var(--text3)">No messages yet.<br>Say hello 👋</div>';

   setTimeout(() => { el('messagesBody').scrollTop = el('messagesBody').scrollHeight; }, 50);

   const unreads = pMsgs.filter(m => m.receiver === me && !m.is_read);
   unreads.forEach(m => markMsgRead(m.id));
}

function openNewChatModal() {
   const dl = el('empBadgeSuggest3');
   if(dl) dl.innerHTML = STATE.empList.map(e=>`<option value="${e.code||e.badge}">${e.name} (${e.dept})</option>`).join('');
   el('newChatRecipient').value='';
   el('newChatMessage').value='';
   openModal('newChatModal');
}
async function sendNewChatMessage() {
   const peer = el('newChatRecipient').value.trim();
   const msg = el('newChatMessage').value.trim();
   if(!peer || !msg) return;
   try {
     const r = await zkAPI('/api/messages',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({receiver:peer,message:msg})});
     if(r.error) throw new Error(r.error);
     closeModal('newChatModal');
     toast('✅ Message sent!');
     loadMessages();
   } catch(e) { toast('❌ '+e.message); }
}

async function sendChatMessage(){
  const msg=el('msgText').value.trim();
  if(!ACTIVE_CHAT_PEER){toast('Select a conversation first');return;}
  if(!msg)return;
  try{
    const r=await zkAPI('/api/messages',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({receiver:ACTIVE_CHAT_PEER,message:msg})});
    if(r.error)throw new Error(r.error);
    el('msgText').value='';
    loadMessages();
  }catch(e){toast('❌ '+e.message)}
}

async function deleteMessage(id){
  if(!confirm("Delete this message?")) return;
  try{
    const r=await zkAPI('/api/messages/'+id,{method:'DELETE'});
    if(r.error)throw new Error(r.error);
    loadMessages();
  }catch(e){toast('❌ '+e.message)}
}

async function markMsgRead(id){
  try{
    await zkAPI('/api/messages/'+id+'/read',{method:'POST'});
  }catch(e){}
}
// ── POLL MESSAGES LOOP ──
// Stored so doLogout() can clear it alongside other timers
STATE._msgPollTimer = setInterval(()=>{if(STATE.user && STATE.currentPage!=='messages') loadMessages()}, 60000);
