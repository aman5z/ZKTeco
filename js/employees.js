async function loadEmployees(){
  try{
    if(STATE.isDemo){STATE.empList=DEMO_EMPLOYEES;}
    else{const d=await zkAPI('/api/employees');STATE.empList=d.employees||[];}
    el('empCountLabel').textContent=STATE.empList.length+' employees';
    el('sbEmpCount').textContent=STATE.empList.length;
    const depts=[...new Set(STATE.empList.map(e=>e.dept||'').filter(Boolean))].sort();
    el('empDeptFilter').innerHTML='<option value="">All Departments</option>'+depts.map(d=>`<option>${esc(d)}</option>`).join('');
    filterEmployees();
    initCalendarSelect();
  }catch(e){console.warn('Employees:',e)}
}
function filterEmployees(){
  const q=(el('empSearch').value||'').toLowerCase();
  const dept=el('empDeptFilter').value;
  const status=el('empStatusFilter').value;
  let emps=STATE.empList;
  if(q)emps=emps.filter(e=>(e.name||'').toLowerCase().includes(q)||(e.code||e.badge||'').includes(q));
  if(dept)emps=emps.filter(e=>e.dept===dept);
  if(status==='active')emps=emps.filter(e=>e.active!==0&&e.active!==false);
  if(status==='inactive')emps=emps.filter(e=>e.active===0||e.active===false);
  renderEmployeeList(emps);
}
function renderEmployeeList(emps){
  // Group by dept
  const groups={};
  emps.forEach(e=>{const d=e.dept||'—';if(!groups[d])groups[d]=[];groups[d].push(e)});
  const list=el('empList');
  list.innerHTML='';
  if(!emps.length){list.innerHTML='<div class="empty-state"><div class="icon">👥</div><p>No employees found</p></div>';return}
  Object.entries(groups).sort(([a],[b])=>a.localeCompare(b)).forEach(([dept,rows])=>{
    const sec=document.createElement('div');sec.className='dept-section';
    sec.innerHTML=`<div class="dept-header" onclick="this.parentElement.classList.toggle('collapsed')">
      <strong>${esc(dept)}</strong><span class="dept-count">${rows.length} employees</span><span class="dept-arrow">▾</span>
    </div><div class="dept-rows">
      <div class="table-wrap"><table><thead><tr><th>Code</th><th>Name</th><th>Department</th><th>Status</th></tr></thead>
      <tbody>${rows.map(e=>`<tr onclick="selectEmployee('${esc(e.code||e.badge)}')">
        <td class="td-mono">${esc(e.code||e.badge||'—')}</td>
        <td>${esc(e.name)}</td>
        <td><span class="tag tag-off">${esc(e.dept)}</span></td>
        <td><span class="tag ${e.active!==0&&e.active!==false?'tag-present':'tag-absent'}">${e.active!==0&&e.active!==false?'Active':'Inactive'}</span></td>
      </tr>`).join('')}</tbody></table></div>
    </div>`;
    list.appendChild(sec);
  });
}
function selectEmployee(code){
  const emp=STATE.empList.find(e=>(e.code||e.badge)===code);
  if(!emp)return;
  STATE.selectedEmp=emp;
  const domain=CFG.domain;
  el('empDetailPane').innerHTML=`
    <div class="dp-header">Employee Properties</div>
    <div class="dp-avatar"><div class="dp-av">${(emp.name||'?').split(' ').map(w=>w[0]).join('').substring(0,2).toUpperCase()}</div>
      <div class="dp-name">${esc(emp.name)}</div>
      <div class="dp-sub">${esc(emp.code||emp.badge)+'@'+domain}</div>
    </div>
    <div class="dp-row"><label>Badge / Code</label><span class="text-mono">${esc(emp.code||emp.badge)}</span></div>
    <div class="dp-row"><label>Department</label><span>${esc(emp.dept||'—')}</span></div>
    <div class="dp-row"><label>Status</label><span class="tag ${emp.active!==0?'tag-present':'tag-absent'}">${emp.active!==0?'Active':'Inactive'}</span></div>
    <div class="dp-actions">
      <button class="dp-btn" onclick="viewEmpCalendar('${esc(emp.code||emp.badge)}')">🗓 View Calendar</button>
      <button class="dp-btn" onclick="openAddEmpUserFor('${esc(emp.code||emp.badge)}','${esc(emp.name)}','${esc(emp.dept)}')" data-admin>🔐 Create Login</button>
      ${STATE.isAdmin?`<button class="dp-btn danger" onclick="toggleEmpActive('${esc(emp.code||emp.badge)}',${emp.active?0:1})">⊘ ${emp.active?'Deactivate':'Activate'}</button>`:''}
      ${STATE.isAdmin?`<button class="dp-btn danger" onclick="deleteEmployee('${esc(emp.code||emp.badge)}','${esc(emp.name)}')" title="Permanently remove this employee">🗑 Delete</button>`:''}
    </div>
  `;
}
function viewEmpCalendar(badge){
  showPage('calendar');
  el('calBadgeSelect').value=badge;
  loadCalendar();
}
function selectEmpQuick(code,name,dept){selectEmployee(code)}
async function toggleEmpActive(badge,active){
  if(STATE.isDemo){toast('🎮 Demo');return}
  try{
    await zkAPI('/api/employee/'+badge+'/active',{method:'POST'});
    toast('✅ Employee status updated');loadEmployees();
  }catch(e){toast('❌ '+e.message)}
}
async function deleteEmployee(badge,name){
  if(STATE.isDemo){toast('🎮 Demo');return}
  if(!confirm('Permanently delete '+name+' ('+badge+') from the employee list?\n\nPunch history will be retained, but the employee will no longer appear in reports or the dashboard.'))return;
  try{
    const d=await zkAPI('/api/employee/'+badge,{method:'DELETE'});
    if(d.error)throw new Error(d.error);
    toast('🗑 '+(d.message||'Employee deleted'));
    STATE.selectedEmp=null;
    el('empDetailPane').innerHTML='';
    loadEmployees();
  }catch(e){toast('❌ '+e.message)}
}
function openAddEmpUser(){openAddEmpUserFor('','','');}
function openAddEmpUserFor(badge,name,dept){
  el('uBadge').value=badge;el('uName').value=name;el('uDept').value=dept;
  el('uRole').value='employee';el('uEnabled').checked=true;el('uMustChange').checked=false;
  el('uPass').value='';el('uPassC').value='';el('uErr').textContent='';
  el('userModalTitle').textContent='👤 Create Login for '+(name||'Employee');
  buildPermGrid(null);
  openModal('userModal');
}

