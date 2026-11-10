/* DEBBIOM v3.4.1 — auth.js (seguridad corregida) */

var _DB=null;
var SUPABASE_URL='https://ynopjrdndccjqyxcqnzl.supabase.co';
var SUPABASE_KEY='sb_publishable_DbJ-Qbg94jLDbaUSNIkfmw_o-vQ4EyK';

function getDB(){
if(_DB)return _DB;
if(!SUPABASE_URL||!SUPABASE_KEY)return null;
if(typeof supabase==='undefined'||!supabase.createClient)return null;
_DB=supabase.createClient(SUPABASE_URL,SUPABASE_KEY,{auth:{persistSession:true,storage:{
  getItem:function(k){try{return sessionStorage.getItem(k);}catch(e){return null;}},
  setItem:function(k,v){try{sessionStorage.setItem(k,v);}catch(e){}},
  removeItem:function(k){try{sessionStorage.removeItem(k);}catch(e){}}
}}});
return _DB;
}

var RATE_LIMIT_URL=SUPABASE_URL?SUPABASE_URL+'/functions/v1/rate-limit':null;
var CREATE_USER_URL=SUPABASE_URL?SUPABASE_URL+'/functions/v1/create-user':null;
var AUTH_DOMAIN='debbiom.app';
function _toEmail(u){return u.trim().toLowerCase()+'@'+AUTH_DOMAIN;}
function _fromEmail(e){return e?e.replace('@'+AUTH_DOMAIN,''):'';}

var _guard={a:0,l:0,mx:5,ms:60000,
locked:function(){if(Date.now()<this.l)return true;if(this.l>0&&Date.now()>=this.l){this.a=0;this.l=0;}return false;},
fail:function(){this.a++;if(this.a>=this.mx)this.l=Date.now()+this.ms;},
ok:function(){this.a=0;this.l=0;},
rem:function(){return Math.ceil(Math.max(0,this.l-Date.now())/1000);}
};

var AUTH={currentUser:null,loginError:'',showPassword:false,users:[],_initialized:false};

async function initAuth(){
var c=getDB();if(!c)return false;
try{
var sr=await c.auth.getSession();
if(sr.data&&sr.data.session&&sr.data.session.user){
var p=await _loadProfile(sr.data.session.user.id);
if(p)AUTH.currentUser={id:sr.data.session.user.id,username:p.username,role:p.role,fullName:p.full_name};
}
c.auth.onAuthStateChange(function(ev){if(ev==='SIGNED_OUT'){AUTH.currentUser=null;if(typeof render==='function')render();}});
AUTH._initialized=true;return true;
}catch(e){return false;}
}

async function attemptLogin(username,password){
if(!username||!password){AUTH.loginError='Ingresa usuario y contrase\u00f1a';return false;}
username=username.trim();
if(username.length>50||password.length>128){AUTH.loginError='Datos no v\u00e1lidos';return false;}
if(_guard.locked()){AUTH.loginError='Demasiados intentos. Espera '+_guard.rem()+'s';return false;}
var c=getDB();if(!c){AUTH.loginError='Sin conexi\u00f3n';return false;}
try{
var loginOk=false;var userId=null;
if(RATE_LIMIT_URL){
try{
var resp=await fetch(RATE_LIMIT_URL,{method:'POST',headers:{'Content-Type':'application/json','apikey':SUPABASE_KEY},body:JSON.stringify({username:username,password:password})});
var result=await resp.json();
if(!resp.ok||result.error){
_guard.fail();
if(typeof addLog==='function')addLog('loginFail','Intento fallido: '+username);
AUTH.loginError=result.error||'Usuario o contrasena incorrectos';
if(result.locked){AUTH.loginError='Cuenta bloqueada ('+(result.remaining||60)+'s)';if(typeof addLog==='function')addLog('loginBlock','Bloqueo: '+username);}
return false;
}
if(result.session){
var sr=await c.auth.setSession({access_token:result.session.access_token,refresh_token:result.session.refresh_token});
if(sr.error){AUTH.loginError='Error al restaurar sesion';return false;}
userId=result.user?result.user.id:(sr.data.user?sr.data.user.id:null);
loginOk=true;
}
}catch(_rlErr){/* rate-limit no disponible, usar login directo */}
}
if(!loginOk){
var r=await c.auth.signInWithPassword({email:_toEmail(username),password:password});
if(r.error){
_guard.fail();
if(typeof addLog==='function')addLog('loginFail','Intento fallido: '+username);
if(_guard.locked()){AUTH.loginError='Cuenta bloqueada (60s)';if(typeof addLog==='function')addLog('loginBlock','Bloqueo: '+username);}
else{AUTH.loginError='Usuario o contrase\u00f1a incorrectos ('+(5-_guard.a)+' intentos)';}
return false;
}
userId=r.data.user.id;
}
_guard.ok();
var p=await _loadProfile(userId);
if(!p){AUTH.loginError='Sin perfil. Contacta al administrador.';await c.auth.signOut();return false;}
AUTH.currentUser={id:userId,username:p.username,role:p.role,fullName:p.full_name};
AUTH.loginError='';
if(typeof initSync==='function'){
var so=await initSync();
if(so){
if(typeof syncLoadState==='function'){var ld=await syncLoadState();if(ld&&typeof render==='function')render();}
if(typeof listenStateChanges==='function')listenStateChanges();
}
}
if(typeof registerPresence==='function')registerPresence();
if(typeof addLog==='function')addLog('login','Sesi\u00f3n iniciada');
if(typeof saveState==='function')saveState();
if(typeof render==='function')render();
return true;
}catch(e){AUTH.loginError='Error de conexi\u00f3n. Intenta de nuevo.';return false;}
}

async function logout(){
if(typeof addLog==='function')addLog('logout','Sesi\u00f3n cerrada');
if(typeof saveState==='function')saveState();
if(typeof STATE!=='undefined')STATE._alertBannerShown=false;
if(typeof unregisterPresence==='function')unregisterPresence();
var c=getDB();if(c)await c.auth.signOut();
AUTH.currentUser=null;
try{sessionStorage.clear();}catch(e){}
if(typeof render==='function')render();
}

function isAdmin(){return AUTH.currentUser&&AUTH.currentUser.role==='admin';}

async function _loadProfile(uid){
var c=getDB();if(!c)return null;
try{var r=await c.from('profiles').select('username,role,full_name').eq('id',uid).single();return(r.error||!r.data)?null:r.data;}catch(e){return null;}
}

async function loadAllProfiles(){
var c=getDB();if(!c)return[];
try{var r=await c.from('profiles').select('id,username,role,full_name,created_at').order('created_at',{ascending:true}).limit(200);if(r.error||!r.data)return[];AUTH.users=r.data;return r.data;}catch(e){return[];}
}

async function createUser(username,password,fullName,role){
var c=getDB();if(!c||!isAdmin())return{error:'Sin permisos'};
username=username.trim();
if(!validateUsername(username))return{error:'Usuario inv\u00e1lido'};
if(!validatePassword(password))return{error:'Contrase\u00f1a: m\u00edn. 8, may\u00fascula, min\u00fascula, n\u00famero y car\u00e1cter especial'};
try{
  /* CRIT-03: Usar Edge Function con service_role para crear usuarios @debbiom.app */
  if(CREATE_USER_URL){
    var sess=await c.auth.getSession();
    var token=sess.data&&sess.data.session?sess.data.session.access_token:null;
    if(!token)return{error:'Sesion expirada. Vuelve a iniciar sesion.'};
    var resp=await fetch(CREATE_USER_URL,{method:'POST',headers:{'Content-Type':'application/json','Authorization':'Bearer '+token,'apikey':SUPABASE_KEY},body:JSON.stringify({username:username,password:password,fullName:fullName||username,role:role||'user'})});
    var result=await resp.json();
    if(!resp.ok||result.error)return{error:result.error||'Error al crear usuario'};
    if(typeof addLog==='function')addLog('userAdd',username);
    return{success:true,userId:result.userId||null};
  }
  /* Fallback: signUp directo (solo funciona con dominios permitidos) */
  var r=await c.auth.signUp({email:_toEmail(username),password:password,options:{data:{username:username,role:role||'user',full_name:fullName||username}}});
  if(r.error)return{error:'Error al crear usuario. Verifica los datos.'};
  var uid=r.data.user?r.data.user.id:null;
  if(uid){await c.from('profiles').insert({id:uid,username:sanitizeInput(username),role:role||'user',full_name:sanitizeInput(fullName)||sanitizeInput(username)});}
  if(typeof addLog==='function')addLog('userAdd',username);
  return{success:true,userId:uid};
}catch(e){return{error:'Error al crear usuario'};}
}

async function changeOwnPassword(np){
var c=getDB();if(!c||!AUTH.currentUser)return{error:'Sin sesi\u00f3n'};
if(!validatePassword(np))return{error:'Contrase\u00f1a: m\u00edn. 8, may\u00fascula, min\u00fascula, n\u00famero y car\u00e1cter especial'};
try{var r=await c.auth.updateUser({password:np});if(r.error)return{error:'Error al cambiar contrase\u00f1a'};if(typeof addLog==='function')addLog('userPw',AUTH.currentUser.username);return{success:true};}catch(e){return{error:'Error de conexi\u00f3n'};}
}

async function changeUserRole(uid,nr){
var c=getDB();if(!c||!isAdmin())return{error:'Sin permisos'};
/* MED-01: Validar que no se quede sin administradores */
try{
  var currentProfile=await c.from('profiles').select('role').eq('id',uid).single();
  if(currentProfile.data&&currentProfile.data.role==='admin'&&nr!=='admin'){
    var admins=await c.from('profiles').select('id').eq('role','admin');
    if(admins.data&&admins.data.length<=1)return{error:'No se puede cambiar: es el \u00faltimo administrador del sistema.'};
  }
  var r=await c.from('profiles').update({role:nr}).eq('id',uid);
  if(r.error)return{error:'Error al cambiar rol'};
  if(typeof addLog==='function')addLog('userRole','Rol: '+nr);
  return{success:true};
}catch(e){return{error:'Error de conexi\u00f3n'};}
}

async function deactivateUser(uid,un){
var c=getDB();if(!c||!isAdmin())return{error:'Sin permisos'};
if(AUTH.currentUser&&AUTH.currentUser.id===uid)return{error:'No puedes desactivarte'};
/* MED-01: Validar ultimo admin al desactivar */
try{
  var currentProfile=await c.from('profiles').select('role').eq('id',uid).single();
  if(currentProfile.data&&currentProfile.data.role==='admin'){
    var admins=await c.from('profiles').select('id').eq('role','admin');
    if(admins.data&&admins.data.length<=1)return{error:'No se puede desactivar: es el \u00faltimo administrador.'};
  }
  var r=await c.from('profiles').update({role:'disabled'}).eq('id',uid);
  if(r.error)return{error:'Error al desactivar usuario'};
  if(typeof addLog==='function')addLog('userDel',un+' desactivado');
  return{success:true};
}catch(e){return{error:'Error de conexi\u00f3n'};}
}

function sanitizeInput(s){if(typeof s!=='string')return'';return s.replace(/[<>"'&]/g,function(c){return({'<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;','&':'&amp;'})[c]||c;});}
function validateUsername(u){return u&&typeof u==='string'&&/^[a-zA-Z0-9._-]{2,30}$/.test(u.trim());}
function validatePassword(p){if(!p||typeof p!=='string')return false;return p.length>=8&&p.length<=128&&/[A-Z]/.test(p)&&/[a-z]/.test(p)&&/[0-9]/.test(p)&&/[!@#$%^&*()_+\-=\[\]{};':"\\|,.<>\/?]/.test(p);}

var SESSION_TIMEOUT=30*60*1000;var _lastAct=Date.now();
function _rAct(){_lastAct=Date.now();}
document.addEventListener('click',_rAct);
document.addEventListener('keydown',_rAct);
document.addEventListener('scroll',_rAct);
setInterval(function(){if(AUTH.currentUser&&(Date.now()-_lastAct>SESSION_TIMEOUT)){if(typeof addLog==='function')addLog('sessionTimeout','Inactividad 30 min');if(typeof logout==='function')logout();if(typeof showToast==='function')showToast('Sesi\u00f3n cerrada por inactividad','error');}},60000);
