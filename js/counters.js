async function loadCounters(){
  spin(true);
  try{
    if(STATE.isDemo){STATE.counters=DEMO_COUNTERS;}
    else if(CFG.gasUrl){
      const d=await gasAPI({action:'getCounters'});
      if(Array.isArray(d))STATE.counters=d.slice(1).map(r=>({id:r[0],name:r[1]||'Counter',value:Number(r[2])||1,key:r[3]||'',prefix:r[4]||'A'})).filter(c=>c.id);
    }else{STATE.counters=[];}
    renderCounters();
  }catch(e){toast('❌ Counters: '+e.message);console.error('Counters error:',e);}
  spin(false);
}
function renderCounters(){
  const canManage=STATE.isAdmin||can('tokens.manage');
  el('counterGrid').innerHTML=STATE.counters.map(c=>`
    <div class="counter-card">
      <div class="counter-name">${esc(c.name)}</div>
      <div><span class="counter-prefix">${esc(c.prefix||'')}</span><span class="counter-value">${c.value}</span></div>
      <div class="counter-actions">
        ${canManage?`<button class="btn btn-secondary btn-sm" onclick="counterAction('prev','${c.id}')">◀ Prev</button>`:''}
        <button class="btn btn-primary btn-sm" onclick="counterAction('next','${c.id}')">▶ Next</button>
        <button class="btn btn-secondary btn-sm" onclick="counterAction('repeat','${c.id}')">↻ Repeat</button>
        ${canManage?`<button class="btn btn-danger btn-sm" onclick="counterAction('reset','${c.id}')">⟳ Reset</button>`:''}
      </div>
    </div>
  `).join('')||'<div class="empty-state"><div class="icon">🎟</div><p>No counters configured</p></div>';
}
async function counterAction(action,id){
  spin(true);
  try{
    if(STATE.isDemo){
      const c=STATE.counters.find(x=>x.id==id);if(!c){spin(false);return}
      if(action==='next')c.value++;
      else if(action==='prev')c.value=Math.max(1,c.value-1);
      else if(action==='reset')c.value=1;
      renderCounters();spin(false);return;
    }
    if(CFG.gasUrl){
      const actionMap={next:'nextToken',prev:'previousToken',repeat:'repeatToken',reset:'resetCounter'};
      await gasAPI({action:actionMap[action],counterId:id});
      await loadCounters();
    }
  }catch(e){toast('❌ '+e.message)}
  spin(false);
}
function openAddCounter(){el('newCounterName').value='';el('cErr').textContent='';openModal('addCounterModal')}
async function submitAddCounter(){
  const name=el('newCounterName').value.trim();
  if(!name){el('cErr').textContent='Name required';return}
  spin(true);
  try{
    if(STATE.isDemo){STATE.counters.push({id:'demo-'+Date.now(),name,value:1,prefix:String.fromCharCode(65+STATE.counters.length%26),key:''});closeModal('addCounterModal');toast('✅ Counter added (demo)');renderCounters();spin(false);return}
    if(CFG.gasUrl){await gasAPI({action:'addCounter',name});closeModal('addCounterModal');toast('✅ Counter added');loadCounters();}
  }catch(e){el('cErr').textContent=e.message}
  spin(false);
}

