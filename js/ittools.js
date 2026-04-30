function setLayout(layout) {
  document.body.setAttribute('data-layout', layout);
  localStorage.setItem('erp_layout', layout);
  document.querySelectorAll('#layoutList .theme-btn').forEach(b => {
    if(b.getAttribute('data-layout') === layout) b.classList.add('active');
    else b.classList.remove('active');
  });
}
// Init Layout
const savedLayout = localStorage.getItem('erp_layout') || 'sidebar';
document.body.setAttribute('data-layout', savedLayout);
if(el('layoutList')){
    document.querySelectorAll('#layoutList .theme-btn').forEach(b => {
        if(b.getAttribute('data-layout') === savedLayout) b.classList.add('active');
        else b.classList.remove('active');
    });
}

// IT TOOLS — all functions prefixed it* to avoid dashboard conflicts
function itFilterTools(){const q=(document.getElementById('itSearchBar')||{}).value||'';const qL=q.toLowerCase().trim();const boxes=document.querySelectorAll('#ittools-wrap .tool-box');let vis=0;boxes.forEach(b=>{const show=!qL||(b.getAttribute('data-tool')||'').toLowerCase().includes(qL)||(b.querySelector('h3')?.textContent.toLowerCase()||'').includes(qL);b.style.display=show?'':'none';if(show)vis++;});const nr=document.getElementById('itNoResults');if(nr)nr.style.display=vis===0?'block':'none';const nr2=document.getElementById('noResults');if(nr2)nr2.style.display=vis===0?'block':'none';}
let itFavs=JSON.parse(localStorage.getItem('itFavs')||'[]');
function itToggleFav(btn,name){const idx=itFavs.indexOf(name);if(idx>=0){itFavs.splice(idx,1);btn.textContent='☆';btn.classList.remove('active');}else{itFavs.push(name);btn.textContent='⭐';btn.classList.add('active');}localStorage.setItem('itFavs',JSON.stringify(itFavs));}
function itCopyPre(id){navigator.clipboard.writeText(document.getElementById(id)?.textContent||'').catch(()=>{});toast('✅ Copied!');}
function itExportAll(){if(!ebAllEmails||!ebAllEmails.length){toast('⚠️ No Email Batch data to export');return;}const out=`IT TOOLS EXPORT\n${new Date().toLocaleString()}\n${'═'.repeat(50)}\n\n[ Email Batch Manager ]\nTotal unique emails: ${ebAllEmails.length}\nBatches: ${ebBatches.length}\n\n${ebBatches.map((b,i)=>`--- Batch ${i+1} (${b.length} emails) ---\n${b.join(', ')}`).join('\n\n')}\n`;const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([out],{type:'text/plain'}));a.download='IT-Export.txt';a.click();toast('✅ Exported!');}
const EB_BATCH_SIZE=500,EB_EMAIL_REGEX=/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;
let ebAllEmails=[],ebBatches=[];
function ebEv(e,cls){e.preventDefault();document.getElementById('ebUploadZone').className='it-drop'+(cls?' eb-drag':'');}
function ebDrop(e){e.preventDefault();document.getElementById('ebUploadZone').className='it-drop';const f=e.dataTransfer.files[0];if(f&&f.name.endsWith('.xlsx'))ebHandleFile(f);else toast('❌ Drop an .xlsx file');}
function ebHandleFile(file){if(!file)return;document.getElementById('ebFileName').textContent='📄 '+file.name;const reader=new FileReader();reader.onload=function(e){try{if(typeof XLSX==='undefined'){toast('❌ XLSX not loaded');return;}const data=new Uint8Array(e.target.result),wb=XLSX.read(data,{type:'array'}),rawEmails=[];wb.SheetNames.forEach(name=>{XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:''}).forEach(row=>row.forEach(cell=>{const found=String(cell).match(EB_EMAIL_REGEX);if(found)found.forEach(em=>rawEmails.push(em.toLowerCase()));}));});const totalRaw=rawEmails.length,unique=[...new Set(rawEmails)].sort(),dupes=totalRaw-unique.length,numB=Math.ceil(unique.length/EB_BATCH_SIZE);ebBatches=Array.from({length:numB},(_,i)=>unique.slice(i*EB_BATCH_SIZE,(i+1)*EB_BATCH_SIZE));ebAllEmails=unique;document.getElementById('eb-s-total').textContent=totalRaw.toLocaleString();document.getElementById('eb-s-unique').textContent=unique.length.toLocaleString();document.getElementById('eb-s-dupes').textContent=dupes.toLocaleString();document.getElementById('eb-s-batches').textContent=numB.toLocaleString();['ebBtnExport','ebBtnCopyAll','ebBtnReset'].forEach(id=>document.getElementById(id).disabled=false);document.getElementById('ebEmptyState').style.display='none';ebRenderBatches(ebBatches);}catch(err){toast('❌ Parse error: '+err.message);}};reader.readAsArrayBuffer(file);}
function ebRenderBatches(batches){const list=document.getElementById('ebBatchList');list.innerHTML='';document.getElementById('ebBatchSection').style.display='block';document.getElementById('ebBatchCountLabel').textContent=batches.length+' batches';batches.forEach((batch,i)=>{const comma=batch.join(', ');const card=document.createElement('div');card.style.cssText='background:var(--surface);border:1px solid var(--border);border-radius:8px;overflow:hidden;margin-bottom:6px;';card.innerHTML=`<div style="display:flex;align-items:center;gap:10px;padding:10px 12px;cursor:pointer" onclick="ebTogBatch(${i})"><span style="background:var(--accent);color:#fff;border-radius:4px;padding:2px 8px;font-size:11px">Batch ${i+1}</span><span style="flex:1;font-size:11px;color:var(--text3)">${batch.length} emails · Acc ${Math.ceil((i+1)/4)} · Send ${(i%4)+1}/4</span><button class="it-btn gr" style="padding:4px 10px;font-size:11px;margin:0" id="ebCp${i}" onclick="event.stopPropagation();ebCpBatch(${i})">⎘ Copy</button></div><div id="ebBody${i}" style="display:none;border-top:1px solid var(--border);padding:10px 12px;font-size:11px;font-family:monospace;color:var(--text3);max-height:110px;overflow-y:auto;word-break:break-all">${comma}</div>`;card.querySelector('#ebCp'+i)._raw=comma;list.appendChild(card);});}
function ebTogBatch(i){const b=document.getElementById('ebBody'+i);if(b)b.style.display=b.style.display==='block'?'none':'block';}
function ebCpBatch(i){const btn=document.getElementById('ebCp'+i);navigator.clipboard.writeText(btn._raw).then(()=>{btn.textContent='✓ Copied!';setTimeout(()=>btn.textContent='⎘ Copy',2000);});}
function ebCopyAll(){navigator.clipboard.writeText(ebAllEmails.join(', ')).then(()=>toast('✅ All emails copied!'));}
function ebExportXlsx(){if(typeof XLSX==='undefined'){toast('❌ XLSX not loaded');return;}const wb=XLSX.utils.book_new(),wsData=[];for(let r=0;r<EB_BATCH_SIZE;r++){const row=[];for(let c=0;c<ebBatches.length;c++)row.push(ebBatches[c]&&ebBatches[c][r]?ebBatches[c][r]:'');wsData.push(row);}const ws=XLSX.utils.aoa_to_sheet(wsData);ws['!cols']=Array(ebBatches.length).fill({wch:38});XLSX.utils.book_append_sheet(wb,ws,'Emails');XLSX.writeFile(wb,'emails_output.xlsx');toast('✅ Exported!');}
function ebReset(){ebAllEmails=[];ebBatches=[];document.getElementById('ebUploadZone').className='it-drop';document.getElementById('ebFileName').textContent='';document.getElementById('eb-file-input').value='';['eb-s-total','eb-s-unique','eb-s-dupes','eb-s-batches'].forEach(id=>document.getElementById(id).textContent='—');['ebBtnExport','ebBtnCopyAll','ebBtnReset'].forEach(id=>document.getElementById(id).disabled=true);document.getElementById('ebBatchSection').style.display='none';document.getElementById('ebEmptyState').style.display='block';document.getElementById('ebBatchList').innerHTML='';}

/* ═══════════════════════════════════════════════════════════
   IT TOOLS — Functions (from Active Directory Console)
═══════════════════════════════════════════════════════════ */

function ad_toast(msg, dur=2800){
  let t=document.getElementById("ad_toast_el");
  if(!t){t=document.createElement("div");t.id="ad_toast_el";t.className="ad-toast-global";document.body.appendChild(t);}
  t.textContent=msg;t.classList.add("show");
  clearTimeout(t._t);t._t=setTimeout(()=>t.classList.remove("show"),dur);
}

function ad_copyOutput(id){const t=document.getElementById(id)?.textContent||"";if(!t.trim())return ad_toast("Nothing to copy!");navigator.clipboard.writeText(t).then(()=>ad_toast("Copied!")).catch(()=>ad_toast("Copy failed"));}

function ad_toggleCollapse(cid,aid){
  const c=document.getElementById(cid),a=document.getElementById(aid);
  if(!c)return;const v=c.style.display==="none";c.style.display=v?"block":"none";
  if(a)a.style.transform=v?"rotate(0deg)":"rotate(-90deg)";
}

function ad_filterTools(){
  const q=(document.getElementById("toolsSearchBar")||{}).value||"";
  const ql=q.toLowerCase().trim();
  const boxes=document.querySelectorAll("#ittools-wrap .tool-box, .tools-wrap .tool-box");
  let vis=0;
  boxes.forEach(b=>{
    const show=!ql||(b.getAttribute("data-tool")||"").toLowerCase().includes(ql)||(b.querySelector("h3")?.textContent.toLowerCase()||"").includes(ql);
    b.style.display=show?"":"none";if(show)vis++;
  });
  const nr=document.getElementById("noResults");if(nr)nr.style.display=vis===0?"block":"none";
  const nr2=document.getElementById("itNoResults");if(nr2)nr2.style.display=vis===0?"block":"none";
}

let ad_favs=JSON.parse(localStorage.getItem("ad_favs")||"[]");
let pingData=[];
let monitorInterval=null;
let qrObj=null;
function ad_toggleFav(btn,name){
  const idx=ad_favs.indexOf(name);
  if(idx>=0){ad_favs.splice(idx,1);btn.textContent="☆";btn.classList.remove("active");}
  else{ad_favs.push(name);btn.textContent="⭐";btn.classList.add("active");}
  localStorage.setItem("ad_favs",JSON.stringify(ad_favs));
  ad_renderFavs();
}
function ad_renderFavs(){
  const sec=document.getElementById("favSection");if(!sec)return;
  const pills=document.getElementById("favPills");if(!pills)return;
  pills.innerHTML="";
  if(!ad_favs.length){sec.style.display="none";return;}
  sec.style.display="block";
  ad_favs.forEach(name=>{
    const pill=document.createElement("span");pill.className="fav-pill";
    const rmBtn=document.createElement("span");
    rmBtn.style.cssText="cursor:pointer;margin-left:4px";rmBtn.textContent="✕";
    rmBtn.dataset.favName=name;
    rmBtn.addEventListener("click",function(ev){ev.stopPropagation();ad_removeFav(this.dataset.favName);});
    pill.textContent=name+" ";pill.appendChild(rmBtn);
    pill.addEventListener("click",()=>{
      const boxes=document.querySelectorAll("#ittools-wrap .tool-box");
      boxes.forEach(b=>{const h=b.querySelector("h3");if(h&&h.textContent.includes(name))b.scrollIntoView({behavior:"smooth",block:"center"});});
    });
    pills.appendChild(pill);
  });
}
function ad_removeFav(name){
  const idx=ad_favs.indexOf(name);if(idx>=0)ad_favs.splice(idx,1);
  localStorage.setItem("ad_favs",JSON.stringify(ad_favs));
  // Reset star buttons
  document.querySelectorAll("#ittools-wrap .star-btn").forEach(btn=>{
    const toolName=btn.getAttribute("onclick")?.match(/'([^']+)'/)?.[1]||"";
    if(toolName===name){btn.textContent="☆";btn.classList.remove("active");}
  });
  ad_renderFavs();
}

function ad_saveHistory(id){const val=document.getElementById(id)?.value||"";if(val.trim())localStorage.setItem("hist_"+id,val);}

function ad_calculateSubnet(){
  let ip=document.getElementById("subnetIp").value.trim(),cidr=parseInt(document.getElementById("subnetCidr").value.trim());
  if(!ip||isNaN(cidr)||cidr<0||cidr>32){ad_toast("Invalid IP or CIDR");return;}
  let oct=ip.split(".").map(Number);
  if(oct.length!==4||oct.some(o=>o<0||o>255)){ad_toast("Invalid IP");return;}
  let ipInt=((oct[0]<<24)>>>0)+(oct[1]<<16)+(oct[2]<<8)+oct[3];
  let maskInt=cidr===0?0:(~0<<(32-cidr))>>>0;
  let mask=[(maskInt>>>24)&255,(maskInt>>>16)&255,(maskInt>>>8)&255,maskInt&255].join(".");
  let netInt=ipInt&maskInt,net=[(netInt>>>24)&255,(netInt>>>16)&255,(netInt>>>8)&255,netInt&255].join(".");
  let bcInt=netInt|(~maskInt>>>0),bc=[(bcInt>>>24)&255,(bcInt>>>16)&255,(bcInt>>>8)&255,bcInt&255].join(".");
  let hosts=cidr>=31?0:Math.pow(2,32-cidr)-2;
  let fh=cidr>=31?net:[(netInt>>>24)&255,(netInt>>>16)&255,(netInt>>>8)&255,(netInt&255)+1].join(".");
  let lh=cidr>=31?bc:[(bcInt>>>24)&255,(bcInt>>>16)&255,(bcInt>>>8)&255,(bcInt&255)-1].join(".");
  document.getElementById("subnetOutput").textContent=`IP Address:        ${ip}\nCIDR:              /${cidr}\nSubnet Mask:       ${mask}\nNetwork Address:   ${net}\nBroadcast Address: ${bc}\nFirst Usable Host: ${fh}\nLast Usable Host:  ${lh}\nUsable Hosts:      ${hosts.toLocaleString()}`;
}
function ad_encodeB64(){document.getElementById("b64output").textContent=btoa(document.getElementById("b64input").value);}
function ad_decodeB64(){try{document.getElementById("b64output").textContent=atob(document.getElementById("b64input").value);}catch{document.getElementById("b64output").textContent="Invalid Base64";}}
function ad_genQR(){
  const val=document.getElementById("qrcodeInput").value.trim(),fg=document.getElementById("qrFg").value,bg=document.getElementById("qrBg").value;
  const el=document.getElementById("qrcode");el.innerHTML="";qrObj=null;
  if(!val)return;
  qrObj=new QRCode(el,{text:val,width:180,height:180,colorDark:fg,colorLight:bg});
}
function ad_getQRImg(){return document.querySelector("#qrcode img")||document.querySelector("#qrcode canvas");}
function ad_copyQRCode(){const img=ad_getQRImg();if(!img){ad_toast("Generate a QR first!");return;}fetch(img.src||img.toDataURL()).then(r=>r.blob()).then(b=>navigator.clipboard.write([new ClipboardItem({"image/png":b})])).then(()=>ad_toast("Copied!")).catch(()=>ad_toast("Failed"));}
function ad_downloadQRCode(){const img=ad_getQRImg();if(!img){ad_toast("Generate a QR first!");return;}const a=document.createElement("a");a.href=img.src||img.toDataURL();a.download="qrcode.png";a.click();}
async function ad_shareQRCode(){const img=ad_getQRImg();if(!img){ad_toast("No QR!");return;}const blob=await fetch(img.src||img.toDataURL()).then(r=>r.blob());const file=new File([blob],"QR.png",{type:"image/png"});if(navigator.canShare&&navigator.canShare({files:[file]}))await navigator.share({files:[file],title:"QR Code"});else ad_toast("Sharing not supported.");}
function ad_genBarcode(){const val=document.getElementById("barcodeInput").value.trim();const svg=document.getElementById("barcode");if(!val){svg.innerHTML="";svg.style.display="none";return;}svg.style.display="block";JsBarcode("#barcode",val,{format:"CODE128",displayValue:true,width:2,height:60,background:"transparent",lineColor:document.documentElement.getAttribute("data-theme")==="dark"?"#e2e8f0":"#1e1b4b"});}
function ad_getSvgCanvas(cb){const svg=document.getElementById("barcode");if(!svg.innerHTML.trim()){ad_toast("No barcode!");return;}const svgData=new XMLSerializer().serializeToString(svg);const canvas=document.createElement("canvas"),ctx=canvas.getContext("2d"),img=new Image();const url=URL.createObjectURL(new Blob([svgData],{type:"image/svg+xml;charset=utf-8"}));img.onload=()=>{canvas.width=img.width;canvas.height=img.height;ctx.fillStyle="#fff";ctx.fillRect(0,0,canvas.width,canvas.height);ctx.drawImage(img,0,0);URL.revokeObjectURL(url);cb(canvas);};img.src=url;}
function ad_copyBarcode(){ad_getSvgCanvas(c=>c.toBlob(b=>navigator.clipboard.write([new ClipboardItem({"image/png":b})]).then(()=>ad_toast("Copied!"))));}
function ad_downloadBarcode(){ad_getSvgCanvas(c=>{const a=document.createElement("a");a.href=c.toDataURL("image/png");a.download="barcode.png";a.click();});}
async function ad_shareBarcode(){ad_getSvgCanvas(async c=>{const blob=await new Promise(r=>c.toBlob(r,"image/png"));const file=new File([blob],"barcode.png",{type:"image/png"});if(navigator.canShare&&navigator.canShare({files:[file]}))await navigator.share({files:[file],title:"Barcode"});else ad_toast("Sharing not supported.");});}
function ad_pdfOut(msg){document.getElementById("pdfOutput").textContent=msg;}
function ad_pdfLink(blob,filename){const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=filename;a.textContent="⬇ Download "+filename;a.style.color="var(--sel)";const out=document.getElementById("pdfOutput");out.textContent="";out.appendChild(a);}
async function ad_mergePDFs(){const files=document.getElementById("mergeFiles").files;if(files.length<2){ad_toast("Select at least 2 PDFs");return;}ad_pdfOut("Merging...");try{const m=await PDFLib.PDFDocument.create();for(const f of files){const p=await PDFLib.PDFDocument.load(await f.arrayBuffer());(await m.copyPages(p,p.getPageIndices())).forEach(pg=>m.addPage(pg));}ad_pdfLink(new Blob([await m.save()],{type:"application/pdf"}),"merged.pdf");}catch(e){ad_pdfOut("Error: "+e);}}
async function ad_splitPDF(){const file=document.getElementById("splitFile").files[0],n=parseInt(document.getElementById("splitPage").value);if(!file||isNaN(n)){ad_toast("Select file and page number");return;}ad_pdfOut("Splitting...");try{const src=await PDFLib.PDFDocument.load(await file.arrayBuffer()),total=src.getPageCount();if(n<=0||n>=total){ad_toast("Page 1–"+(total-1));return;}async function ad_part(idxs){const d=await PDFLib.PDFDocument.create();(await d.copyPages(src,idxs)).forEach(p=>d.addPage(p));return new Blob([await d.save()],{type:"application/pdf"});}const b1=await ad_part([...Array(n).keys()]),b2=await ad_part([...Array(total-n).keys()].map(i=>i+n));const out=document.getElementById("pdfOutput");out.textContent="";[["part1.pdf",b1],["part2.pdf",b2]].forEach(([name,blob])=>{const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=name;a.textContent="⬇ "+name;a.style.marginRight="14px";a.style.color="var(--sel)";out.appendChild(a);});}catch(e){ad_pdfOut("Error: "+e);}}
async function ad_compressPDF(){const file=document.getElementById("compressFile").files[0];if(!file){ad_toast("Select a PDF");return;}ad_pdfOut("Compressing...");try{const bytes=await file.arrayBuffer();const pdf=await PDFLib.PDFDocument.load(bytes,{updateMetadata:false});const comp=await pdf.save({useObjectStreams:true});const before=(bytes.byteLength/1024).toFixed(1),after=(comp.byteLength/1024).toFixed(1),saved=(100-(comp.byteLength/bytes.byteLength*100)).toFixed(1);const blob=new Blob([comp],{type:"application/pdf"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="compressed.pdf";a.textContent="⬇ Download compressed.pdf";a.style.color="var(--sel)";const out=document.getElementById("pdfOutput");out.textContent="Before: "+before+" KB → After: "+after+" KB ("+saved+"% saved)\n";out.appendChild(a);}catch(e){ad_pdfOut("Error: "+e);}}
async function ad_watermarkPDF(){const file=document.getElementById("watermarkFile").files[0],text=document.getElementById("watermarkText").value.trim();if(!file||!text){ad_toast("Select PDF and enter text");return;}ad_pdfOut("Adding watermark...");try{const pdf=await PDFLib.PDFDocument.load(await file.arrayBuffer());const pages=pdf.getPages(),font=await pdf.embedFont(PDFLib.StandardFonts.HelveticaBold);pages.forEach(page=>{const{width,height}=page.getSize();page.drawText(text,{x:width/2-text.length*12,y:height/2,size:40,font,color:PDFLib.rgb(.8,.1,.1),opacity:.3,rotate:PDFLib.degrees(45)});});ad_pdfLink(new Blob([await pdf.save()],{type:"application/pdf"}),"watermarked.pdf");}catch(e){ad_pdfOut("Error: "+e);}}
async function ad_rotatePDF(){const file=document.getElementById("rotateFile").files[0],deg=parseInt(document.getElementById("rotateDeg").value);if(!file){ad_toast("Select a PDF");return;}ad_pdfOut("Rotating...");try{const pdf=await PDFLib.PDFDocument.load(await file.arrayBuffer());pdf.getPages().forEach(p=>p.setRotation(PDFLib.degrees((p.getRotation().angle+deg)%360)));ad_pdfLink(new Blob([await pdf.save()],{type:"application/pdf"}),"rotated.pdf");}catch(e){ad_pdfOut("Error: "+e);}}
function ad_runDiff(){
  const a=document.getElementById("diffOriginal").value.split("\n"),b=document.getElementById("diffModified").value.split("\n");
  const out=document.getElementById("diffOutput");out.innerHTML="";out.style.display="block";
  let added=0,removed=0;const lcs=ad_buildLCS(a,b);let ai=0,bi=0,li=0;const lines=[];
  while(ai<a.length||bi<b.length){
    if(ai<a.length&&bi<b.length&&li<lcs.length&&a[ai]===lcs[li]&&b[bi]===lcs[li]){lines.push({type:"same",text:a[ai]});ai++;bi++;li++;}
    else if(bi<b.length&&(li>=lcs.length||b[bi]!==lcs[li])){lines.push({type:"add",text:b[bi]});bi++;added++;}
    else{lines.push({type:"rem",text:a[ai]});ai++;removed++;}
  }
  lines.forEach(l=>{const div=document.createElement("div");div.className="diff-line "+(l.type==="add"?"diff-add":l.type==="rem"?"diff-rem":"diff-same");div.textContent=(l.type==="add"?"+ ":l.type==="rem"?"− ":" ")+l.text;out.appendChild(div);});
  document.getElementById("diffStats").textContent="+"+added+" added  −"+removed+" removed";
}
function ad_buildLCS(a,b){const m=a.length,n=b.length,dp=Array.from({length:m+1},()=>Array(n+1).fill(0));for(let i=1;i<=m;i++)for(let j=1;j<=n;j++)dp[i][j]=a[i-1]===b[j-1]?dp[i-1][j-1]+1:Math.max(dp[i-1][j],dp[i][j-1]);const res=[];let i=m,j=n;while(i>0&&j>0){if(a[i-1]===b[j-1]){res.unshift(a[i-1]);i--;j--;}else if(dp[i-1][j]>dp[i][j-1])i--;else j--;}return res;}
function ad_clearDiff(){document.getElementById("diffOriginal").value="";document.getElementById("diffModified").value="";document.getElementById("diffOutput").innerHTML="";document.getElementById("diffOutput").style.display="none";document.getElementById("diffStats").textContent="";}
function ad_clearPing(){pingData=[];ad_drawChart();document.getElementById("pingOutput").textContent="";document.getElementById("pingStats").textContent="";}
function ad_startMonitoring(){const url=document.getElementById("pingInput").value.trim(),ms=parseInt(document.getElementById("pingInterval").value)*1000;if(!url||isNaN(ms)||ms<1000){ad_toast("Enter valid URL & interval");return;}ad_stopMonitoring();ad_pingServer(url);monitorInterval=setInterval(()=>ad_pingServer(url),ms);}
function ad_stopMonitoring(){clearInterval(monitorInterval);monitorInterval=null;}
async function ad_pingServer(url){const out=document.getElementById("pingOutput");try{const t0=performance.now(),r=await fetch(url,{method:"HEAD",mode:"no-cors"}),ms=parseFloat((performance.now()-t0).toFixed(1));pingData.push({ms,ok:true});if(pingData.length>40)pingData.shift();out.textContent="["+new Date().toLocaleTimeString()+"] "+ms+"ms\n"+out.textContent;out.textContent=out.textContent.split("\n").slice(0,30).join("\n");}catch{pingData.push({ms:null,ok:false});if(pingData.length>40)pingData.shift();out.textContent="["+new Date().toLocaleTimeString()+"] FAILED\n"+out.textContent;}ad_drawChart();ad_updatePingStats();}
function ad_drawChart(){const canvas=document.getElementById("pingChart");const ctx=canvas.getContext("2d");canvas.width=canvas.offsetWidth||600;canvas.height=140;ctx.clearRect(0,0,canvas.width,canvas.height);if(pingData.length<2)return;const vals=pingData.map(d=>d.ms||0),maxVal=Math.max(...vals,1);const w=canvas.width,h=canvas.height,ad_pad=16;ctx.strokeStyle="rgba(128,128,128,.15)";ctx.lineWidth=1;[.25,.5,.75,1].forEach(f=>{ctx.beginPath();ctx.moveTo(ad_pad,h-ad_pad-f*(h-2*ad_pad));ctx.lineTo(w-ad_pad,h-ad_pad-f*(h-2*ad_pad));ctx.stroke();});const isDark=document.documentElement.getAttribute("data-theme")==="dark";ctx.strokeStyle=isDark?"#818cf8":"#4f46e5";ctx.lineWidth=2;ctx.lineJoin="round";ctx.beginPath();pingData.forEach((d,i)=>{const x=ad_pad+(i/(pingData.length-1))*(w-2*ad_pad);const y=d.ok?h-ad_pad-((d.ms/maxVal)*(h-2*ad_pad)):h-ad_pad;if(i===0)ctx.moveTo(x,y);else ctx.lineTo(x,y);ctx.save();ctx.beginPath();ctx.arc(x,y,3,0,Math.PI*2);ctx.fillStyle=d.ok?"#22c55e":"#ef4444";ctx.fill();ctx.restore();});ctx.stroke();ctx.fillStyle=isDark?"#94a3b8":"#6b7280";ctx.font="11px Arial";ctx.fillText(maxVal.toFixed(0)+"ms",2,ad_pad+4);}
function ad_updatePingStats(){const vals=pingData.filter(d=>d.ok).map(d=>d.ms);if(!vals.length)return;const avg=(vals.reduce((a,b)=>a+b,0)/vals.length).toFixed(1),min=Math.min(...vals).toFixed(1),max=Math.max(...vals).toFixed(1),fails=pingData.filter(d=>!d.ok).length;document.getElementById("pingStats").textContent="Avg: "+avg+"ms  Min: "+min+"ms  Max: "+max+"ms  Failures: "+fails+"/"+pingData.length;}
function ad_convertNumber(){let v=document.getElementById("numberInput").value.trim();if(!v)return;let num=/^0x/i.test(v)?parseInt(v,16):/^[01]+$/.test(v)?parseInt(v,2):parseInt(v,10);document.getElementById("numberResult").textContent="Decimal : "+num+"\nHex     : 0x"+num.toString(16).toUpperCase()+"\nBinary  : "+num.toString(2)+"\nOctal   : "+num.toString(8);}
function ad_genPass(){let chars="abcdefghijklmnopqrstuvwxyz";if(document.getElementById("passUpper").checked)chars+="ABCDEFGHIJKLMNOPQRSTUVWXYZ";if(document.getElementById("passNum").checked)chars+="0123456789";if(document.getElementById("passSym").checked)chars+="!@#$%^&*()_+-=[]{}";const len=parseInt(document.getElementById("passLen").value)||16;const p=Array.from({length:len},()=>chars[Math.floor(Math.random()*chars.length)]).join("");document.getElementById("passOut").textContent=p;let s=0;if(len>=12)s++;if(len>=16)s++;if(/[A-Z]/.test(p))s++;if(/[0-9]/.test(p))s++;if(/[^a-zA-Z0-9]/.test(p))s++;const labels=["","Weak 🔴","Fair 🟡","Good 🟢","Strong 💪","Very Strong 🔒"];document.getElementById("passStrength").textContent="Strength: "+labels[Math.min(s,5)];}
function ad_lookupMac(){const mac=document.getElementById("macInput").value.trim();if(!mac){ad_toast("Enter a MAC");return;}fetch("https://api.macvendors.com/"+mac).then(r=>r.text()).then(d=>document.getElementById("macOutput").textContent=d).catch(()=>document.getElementById("macOutput").textContent="Lookup failed");}
async function ad_getLocalIP(){return new Promise(r=>{const rtc=new RTCPeerConnection({iceServers:[]});rtc.createDataChannel("");rtc.createOffer().then(o=>rtc.setLocalDescription(o)).catch(()=>{});rtc.onicecandidate=evt=>{if(!evt||!evt.candidate)return;const m=evt.candidate.candidate.match(/(\d{1,3}(\.\d{1,3}){3})/);if(m)r(m[1]);};setTimeout(()=>r("Unavailable"),2000);});}
async function ad_showSystemInfo(){const out=document.getElementById("sysInfoOutput");out.textContent="Gathering...";let gpu="Unavailable";try{const gl=document.createElement("canvas").getContext("webgl");const di=gl&&gl.getExtension("WEBGL_debug_renderer_info");if(di)gpu=gl.getParameter(di.UNMASKED_RENDERER_WEBGL);}catch{}let battery="Unavailable";try{const b=await navigator.getBattery();battery=Math.round(b.level*100)+"% Charging: "+b.charging;}catch{}let publicIP="Unavailable";try{publicIP=(await(await fetch("https://api.ipify.org?format=json")).json()).ip;}catch{}let localIP="Unavailable";try{localIP=await ad_getLocalIP();}catch{}const conn=navigator.connection||navigator.mozConnection||navigator.webkitConnection;const connType=conn?(conn.effectiveType||conn.type||"Unknown"):"Unavailable";const connSpeed=conn&&conn.downlink?conn.downlink+" Mbps":"?";out.textContent="=== SYSTEM INFO ===\nPublic IP:   "+publicIP+"\nLocal IP:    "+localIP+"\nPlatform:    "+navigator.platform+"\nCPU Cores:   "+navigator.hardwareConcurrency+"\nRAM (approx):"+navigator.deviceMemory+" GB\nGPU:         "+gpu+"\nScreen:      "+window.screen.width+"x"+window.screen.height+" ("+window.devicePixelRatio+"x DPR)\nTimezone:    "+Intl.DateTimeFormat().resolvedOptions().timeZone+"\n\n=== CONNECTION ===\nType:  "+connType+"\nSpeed: "+connSpeed+"\n\n=== BROWSER ===\n"+navigator.userAgent+"\n\n=== BATTERY ===\n"+battery;}
function ad_getIpInfo(){const i=document.getElementById("networkInput").value.trim(),out=document.getElementById("networkOutput");if(!i){ad_toast("Enter IP or domain");return;}out.textContent="Fetching...";fetch(/\d+\.\d+\.\d+\.\d+/.test(i)?`https://ipapi.co/${i}/json/`:"https://ipapi.co/json/").then(r=>r.json()).then(d=>out.textContent=JSON.stringify(d,null,2)).catch(()=>out.textContent="Failed");}
function ad_urlToIp(){const i=document.getElementById("networkInput").value.trim(),out=document.getElementById("networkOutput");if(!i){ad_toast("Enter a URL");return;}out.textContent="Resolving...";fetch("https://dns.google/resolve?name="+i).then(r=>r.json()).then(d=>{out.textContent="IP for "+i+": "+(d.Answer?d.Answer.map(a=>a.data).join(", "):"No IP found");}).catch(()=>out.textContent="Failed");}
function ad_dnsLookup(){const i=document.getElementById("networkInput").value.trim(),out=document.getElementById("networkOutput");if(!i){ad_toast("Enter a domain");return;}out.textContent="Looking up...";fetch("https://dns.google/resolve?name="+i).then(r=>r.json()).then(d=>out.textContent=JSON.stringify(d,null,2)).catch(()=>out.textContent="Failed");}
async function ad_checkHosting(){const domain=document.getElementById("networkInput").value.trim(),out=document.getElementById("networkOutput");if(!domain){ad_toast("Enter a domain");return;}out.textContent="Checking...";try{const dns=await(await fetch("https://dns.google/resolve?name="+domain)).json();if(!dns.Answer){out.textContent="Could not resolve";return;}const ip=dns.Answer.find(a=>a.type===1)?.data;if(!ip){out.textContent="No IPv4 found";return;}const info=await(await fetch("https://ipwhois.app/json/"+ip)).json();out.textContent="Domain:  "+domain+"\nIP:      "+ip+"\nISP:     "+(info.isp||"?")+"\nOrg:     "+(info.org||"?")+"\nCountry: "+(info.country||"?")+"\nCity:    "+(info.city||"?");}catch(e){out.textContent="Error: "+e;}}
async function ad_shortenURL(){const i=document.getElementById("networkInput").value.trim(),out=document.getElementById("networkOutput");if(!i){ad_toast("Enter a URL");return;}out.textContent="Shortening...";try{out.textContent="Shortened:\n"+(await(await fetch("https://tinyurl.com/api-create.php?url="+encodeURIComponent(i))).text());}catch{out.textContent="Failed";}}
function ad_startUpload(makeQR=false){
  if(!toolsFile){ad_toast("Select a file first.");return;}
  toolsFileName=toolsFile.name;
  document.getElementById("qrSection").style.display="none";
  document.getElementById("uploadStatus").textContent="Uploading...";
  const reader=new FileReader();
  reader.onload=e=>{
    fetch(TOOLS_SCRIPT_URL,{method:"POST",body:new URLSearchParams({name:toolsFile.name,mimeType:toolsFile.type,file:e.target.result.split(",")[1]})})
    .then(r=>r.text()).then(link=>{document.getElementById("uploadStatus").textContent="Uploaded:\n"+link;if(makeQR){document.getElementById("qrSection").style.display="block";document.getElementById("qrOutput").innerHTML="";new QRCode(document.getElementById("qrOutput"),{text:link,width:200,height:200});}})
    .catch(err=>document.getElementById("uploadStatus").textContent="Failed: "+err);
  };
  reader.readAsDataURL(toolsFile);
}
function ad_copyQR(){const img=document.querySelector("#qrOutput img");if(!img){ad_toast("No QR!");return;}fetch(img.src).then(r=>r.blob()).then(b=>navigator.clipboard.write([new ClipboardItem({"image/png":b})])).then(()=>ad_toast("Copied!")).catch(()=>ad_toast("Failed"));}
function ad_downloadQR(){const img=document.querySelector("#qrOutput img");if(!img){ad_toast("No QR!");return;}const a=document.createElement("a");a.href=img.src;a.download=(toolsFileName.replace(/\.[^/.]+$/,"")||"qr")+"_QR.png";a.click();}
async function ad_shareQR(){const img=document.querySelector("#qrOutput img");if(!img){ad_toast("No QR!");return;}const blob=await fetch(img.src).then(r=>r.blob());const file=new File([blob],"QR.png",{type:blob.type});if(navigator.canShare&&navigator.canShare({files:[file]}))await navigator.share({files:[file],title:"QR Code"});else ad_toast("Sharing not supported.");}
function ad_exportAllOutputs(){
  const ids=[["Subnet Calculator","subnetOutput"],["Base64","b64output"],["Hex/Dec/Bin","numberResult"],["Password","passOut"],["MAC Lookup","macOutput"],["Network Tools","networkOutput"],["Ping Monitor","pingOutput"],["System Info","sysInfoOutput"],["PDF Log","pdfOutput"],["Upload Status","uploadStatus"]];
  const now=new Date().toLocaleString();let report="IT TOOLS EXPORT\nGenerated: "+now+"\n"+"═".repeat(50)+"\n\n";let hasContent=false;
  ids.forEach(([label,id])=>{const el=document.getElementById(id);const text=el?el.textContent.trim():"";if(text&&!["Press button to scan...","Results will appear here...","Waiting for file...","PDF output here..."].includes(text)){report+="[ "+label+" ]\n"+text+"\n\n"+"─".repeat(40)+"\n\n";hasContent=true;}});
  if(!hasContent){ad_toast("No outputs to export yet");return;}
  const a=document.createElement("a");a.href="data:text/plain;charset=utf-8,"+encodeURIComponent(report);a.download="IT-Tools-Export-"+Date.now()+".txt";a.click();ad_toast("✅ Exported!");
}

// DRIVE UPLOAD (Tools tab)
var TOOLS_SCRIPT_URL="https://script.google.com/macros/s/AKfycbwuumk20HhmH8o2SJsp8n_KgYPYu-3OuOk9JvmEzoSk66yyICYqB39jFCM10_uBUDMb/exec";
var toolsFile=null,toolsFileName="";
(()=>{
  const da=document.getElementById("dropArea"),fi=document.getElementById("fileInput");
  if(!da||!fi) return;
  da.addEventListener("click",()=>fi.click());
  da.addEventListener("dragover",e=>{e.preventDefault();da.style.borderColor="var(--accent)";});
  da.addEventListener("dragleave",()=>{da.style.borderColor="";});
  da.addEventListener("drop",e=>{e.preventDefault();da.style.borderColor="";toolsFile=e.dataTransfer.files[0];document.getElementById("uploadStatus").textContent="Selected: "+toolsFile.name;});
  fi.addEventListener("change",()=>{toolsFile=fi.files[0];document.getElementById("uploadStatus").textContent="Selected: "+toolsFile.name;});
})();

// Init favorites on page load
document.addEventListener("DOMContentLoaded",()=>{
  ad_renderFavs();
  // Restore star states
  document.querySelectorAll("#ittools-wrap .star-btn").forEach(btn=>{
    const m=btn.getAttribute("onclick")?.match(/'([^']+)'/);
    if(m&&ad_favs.includes(m[1])){btn.textContent="⭐";btn.classList.add("active");}
  });
});

