// ===========================================================================
//  INIT
// ===========================================================================
(function init(){
  // Apply saved theme instantly
  document.documentElement.setAttribute('data-theme',ls('theme')||'dark');
  applyBranding();
  // Set history date defaults
  const now=new Date(),yr=now.getFullYear(),mn=String(now.getMonth()+1).padStart(2,'0'),dy=String(now.getDate()).padStart(2,'0');
  el('histTo').value=yr+'-'+mn+'-'+dy;
  const wk=new Date(now);wk.setDate(wk.getDate()-7);
  el('histFrom').value=wk.getFullYear()+'-'+String(wk.getMonth()+1).padStart(2,'0')+'-'+String(wk.getDate()).padStart(2,'0');
  el('calMonth').value=yr+'-'+mn;
  el('adminPunchTime').value=yr+'-'+mn+'-'+dy+'T07:30';
  // Status bar clock
  setInterval(()=>{el('sbClock').textContent=new Date().toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit',second:'2-digit'})},1000);

  // Server Ping
  setInterval(async () => {
    try {
      const r = await fetch(CFG.zkUrl + '/api/ping');
      if(r.ok) {
        STATE.serverStatus = 'online';
        if(el('serverStatusDot')) { el('serverStatusDot').className = 'sb-dot green'; el('serverStatusDot').title = 'Server Online'; }
      } else throw new Error();
    } catch(e) {
      STATE.serverStatus = 'offline';
      if(el('serverStatusDot')) { el('serverStatusDot').className = 'sb-dot red'; el('serverStatusDot').title = 'Server Offline'; }
    }
  }, 5000);

  // PWA Prompt
  let deferredPrompt;
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    if(el('pwaInstallBtn')) {
      el('pwaInstallBtn').style.display = 'inline-block';
      el('pwaInstallBtn').addEventListener('click', async () => {
        if(deferredPrompt) {
          deferredPrompt.prompt();
          const { outcome } = await deferredPrompt.userChoice;
          if (outcome === 'accepted') { deferredPrompt = null; el('pwaInstallBtn').style.display = 'none'; }
        }
      });
    }
  });
  // Add Event listener to SQL input for ctrl+enter
  const sqlInput = el('sqlInput');
  if(sqlInput) {
    sqlInput.addEventListener('keydown', e => {
      if(e.ctrlKey && e.key === 'Enter') runSql();
    });
  }

})();
// ── INITIAL AUTH LOAD ──
window.addEventListener('DOMContentLoaded', async () => {
  async function fetchWeather() {
    try {
      const res = await fetch('https://api.open-meteo.com/v1/forecast?latitude=25.3573&longitude=55.3911&current_weather=true');
      const data = await res.json();
      const temp = Math.round(data.current_weather.temperature);
      const wcode = data.current_weather.weathercode;
      let icon = '☁️';
      if (wcode === 0) icon = '☀️';
      else if (wcode > 0 && wcode <= 3) icon = '⛅';
      else if (wcode >= 45 && wcode <= 48) icon = '🌫️';
      else if (wcode >= 51 && wcode <= 67) icon = '🌧️';
      else if (wcode >= 71 && wcode <= 82) icon = '🌨️';
      else if (wcode >= 95) icon = '🌩️';
      const wEl = document.getElementById('weatherWidget');
      if (wEl) wEl.textContent = `${icon} ${temp}°C Sharjah`;
    } catch(e) { console.log('Weather fetch failed', e); }
  }
  fetchWeather();
    applyBranding();
    try {
        const r = await fetch(CFG.zkUrl + '/api/auth/me', {credentials: 'include'});
        if (r.ok) {
            const d = await r.json();
            if (d.authenticated || d.ok) {
                await onLoginSuccess({
                    username: d.username, name: d.name || d.username, role: d.role || 'employee',
                    badge: d.badge || d.username, permissions: d.permissions || {},
                    theme: d.theme || 'dark', source: 'zk'
                });
                tryGASLogin();
                return;
            }
        }
    } catch(e) {}
    // No active session — load the app shell directly without authentication
    await onLoginSuccess({username:'guest',name:'Guest',role:'employee',badge:'',permissions:{},theme:ls('theme')||'dark',source:'zk'});
});
