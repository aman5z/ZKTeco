async function loadDrive(folderId){
  STATE.driveFolderId=folderId||null;
  spin(true);
  try{
    if(STATE.isDemo){
      el('driveGrid').innerHTML='<div class="empty-state"><div class="icon">🖴</div><p>Google Drive requires GAS connection.</p></div>';
      el('driveStorageInfo').textContent='Google Drive (demo)';
      spin(false);return;
    }
    if(!CFG.gasUrl){el('driveGrid').innerHTML='<div class="empty-state"><div class="icon">⚠</div><p>Configure GAS URL in Settings</p></div>';spin(false);return}
    const params={action:'listDriveFiles'};
    if(folderId)params.folderId=folderId;
    const d=await gasAPI(params);
    if(d.error){el('driveGrid').innerHTML='<div class="empty-state"><div class="icon">⚠</div><p>'+esc(d.error)+'</p></div>';spin(false);return}
    // Storage info
    const si=await gasAPI({action:'getStorageInfo'}).catch(()=>null);
    if(si&&si.usedGB)el('driveStorageInfo').textContent=si.usedGB+' GB / '+si.limitTB+' TB used ('+si.percent+'%)';
    // Files
    const folders=(d.folders||[]).map(f=>`<div class="drive-item" ondblclick="loadDrive('${f.id}')"><div class="icon">📁</div><div class="name">${esc(f.name)}</div></div>`);
    const files=(d.files||[]).map(f=>{
      const icon=f.mimeType.includes('image')?'🖼':f.mimeType.includes('pdf')?'📄':f.mimeType.includes('sheet')?'📊':f.mimeType.includes('document')?'📝':'📎';
      const sz=f.size==null?'—':f.size>1048576?(f.size/1048576).toFixed(1)+' MB':f.size>1024?(f.size/1024).toFixed(0)+' KB':f.size+' B';
      return `<div class="drive-item" ondblclick="window.open('${f.url}','_blank')"><div class="icon">${icon}</div><div class="name">${esc(f.name)}</div><div class="size">${sz}</div></div>`;
    });
    el('driveGrid').innerHTML=folders.concat(files).join('')||'<div class="empty-state"><div class="icon">📁</div><p>Empty folder</p></div>';
    el('drivePath').innerHTML='<span class="drive-path-seg" onclick="loadDrive()">🖴 Drive</span>'+(folderId?`<span class="drive-path-sep">/</span><span class="drive-path-seg">${esc(d.name||'Folder')}</span>`:'');
  }catch(e){el('driveGrid').innerHTML='<div class="empty-state text-red"><div class="icon">⚠</div><p>'+esc(e.message)+'</p></div>'}
  spin(false);
}
function openUploadFile(){el('uploadFile').value='';el('uploadErr').textContent='';openModal('uploadModal')}
async function submitUpload(){
  const file=el('uploadFile').files[0];
  if(!file){el('uploadErr').textContent='Select a file';return}
  spin(true);
  try{
    const b64=await new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result.split(',')[1]);r.onerror=rej;r.readAsDataURL(file)});
    const params={action:'uploadDriveFile',fileName:file.name,fileData:b64,mimeType:file.type||'application/octet-stream'};
    if(STATE.driveFolderId)params.folderId=STATE.driveFolderId;
    await gasAPI(params);
    toast('✅ File uploaded');closeModal('uploadModal');loadDrive(STATE.driveFolderId);
  }catch(e){el('uploadErr').textContent=e.message}
  spin(false);
}

