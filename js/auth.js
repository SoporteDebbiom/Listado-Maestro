var _DB=null,SUPABASE_URL,SUPABASE_KEY;
(function(){var _c=[118,54,122,0,19,18,95,74,49,60,32,96,78,71,22,24,18,244,225,224,253,206,195,205,214,214,178,179,169,187,174,136,151,152,213,113,124,96,118,124,68,95,86,20,34,39,109,122,127,6,73,72,91,243,229,209,229,233,193,198,216,203,215,167,175,184,190,189,173,146,189,211,84,110,116,35,21,66,99,114,95,37,30,1,23,41,12,8,24,11,220,229,188,238,206,146,232,205,240,224,180];var s='';for(var i=0;i<_c.length;i++)s+=String.fromCharCode(_c[i]^((i*7+13)%256));try{var d=JSON.parse(s);SUPABASE_URL=d.a;SUPABASE_KEY=d.b;}catch(e){}})();

var _memStore={};
function getDB(){
if(_DB)return _DB;
if(!SUPABASE_URL||!SUPABASE_KEY)return null;
if(typeof supabase==='undefined'||!supabase.createClient)return null;
_DB=supabase.createClient(SUPABASE_URL,SUPABASE_KEY,{auth:{persistSession:true,storage:{
  getItem:function(k){return _memStore[k]||null;},
  setItem:function(k,v){_memStore[k]=v;},
  removeItem:function(k){delete _memStore[k];}
}}});
return _DB;
}

var _csrfToken=(function(){var a=new Uint8Array(24);if(typeof crypto!=='undefined'&&crypto.getRandomValues)crypto.getRandomValues(a);else for(var i=0;i<24;i++)a[i]=Math.floor(Math.random()*256);return Array.from(a,function(b){return b.toString(16).padStart(2,'0');}).join('');})();
function getCsrfToken(){return _csrfToken;}
function validateCsrf(t){return t===_csrfToken;}

var RATE_LIMIT_URL=SUPABASE_URL?SUPABASE_URL+'/functions/v1/rate-limit':null;
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
try{var r=await c.auth.updateUser({password:np});if(r.error)return{error:'Error al cambiar contraseña'};if(typeof addLog==='function')addLog('userPw',AUTH.currentUser.username);return{success:true};}catch(e){return{error:'Error de conexión'};}
}

async function changeUserRole(uid,nr){
var c=getDB();if(!c||!isAdmin())return{error:'Sin permisos'};
try{var r=await c.from('profiles').update({role:nr}).eq('id',uid);if(r.error)return{error:'Error al cambiar rol'};if(typeof addLog==='function')addLog('userRole','Rol: '+nr);return{success:true};}catch(e){return{error:'Error de conexión'};}
}

async function deactivateUser(uid,un){
var c=getDB();if(!c||!isAdmin())return{error:'Sin permisos'};
if(AUTH.currentUser&&AUTH.currentUser.id===uid)return{error:'No puedes desactivarte'};
try{var r=await c.from('profiles').update({role:'disabled'}).eq('id',uid);if(r.error)return{error:'Error al desactivar usuario'};if(typeof addLog==='function')addLog('userDel',un+' desactivado');return{success:true};}catch(e){return{error:'Error de conexión'};}
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
