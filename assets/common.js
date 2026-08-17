/* ============================================================
   Rumpess Timber And Boards Center — shared app helpers
   Used by index.html, employees.html, boss.html
   Requires APPS_SCRIPT_URL and APP_SECRET to be defined before this loads.
   ============================================================ */

function money(n){ return 'KES ' + Number(n||0).toLocaleString(undefined,{maximumFractionDigits:0}); }
function esc(s){ return (s||'').toString().replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function fmtDate(v){ const d=new Date(v); return isNaN(d) ? '' : d.toLocaleDateString(undefined,{day:'2-digit',month:'short',year:'numeric'})+' '+d.toLocaleTimeString(undefined,{hour:'2-digit',minute:'2-digit'}); }

function showToast(msg){
  const t=document.getElementById('toast');
  if(!t) return;
  t.textContent=msg; t.classList.add('show'); setTimeout(()=>t.classList.remove('show'),2400);
}

async function apiGet(){
  const res = await fetch(APPS_SCRIPT_URL + '?action=getData&secret=' + encodeURIComponent(APP_SECRET));
  return res.json();
}
async function apiPost(action, payload){
  const res = await fetch(APPS_SCRIPT_URL, {
    method:'POST',
    headers:{'Content-Type':'text/plain;charset=utf-8'},
    body: JSON.stringify(Object.assign({action, secret: APP_SECRET}, payload))
  });
  return res.json();
}

function isConfigured(){
  return APPS_SCRIPT_URL && APPS_SCRIPT_URL.indexOf('PASTE_YOUR')!==0
      && APP_SECRET && APP_SECRET.indexOf('PASTE_YOUR')!==0;
}

function getSession(){
  try{ return JSON.parse(localStorage.getItem('tyd_session')); }catch(e){ return null; }
}
function clearSession(){ localStorage.removeItem('tyd_session'); }
function logout(){ clearSession(); window.location.href = '/'; }

// Redirects to login if no session. Pass 'boss' to restrict a page to boss-only.
function requireSession(restrictToRole){
  const session = getSession();
  if(!session || !session.role){ window.location.href = '/'; return null; }
  if(restrictToRole && session.role !== restrictToRole){
    window.location.href = session.role === 'boss' ? '/boss.html' : '/employees.html';
    return null;
  }
  return session;
}
