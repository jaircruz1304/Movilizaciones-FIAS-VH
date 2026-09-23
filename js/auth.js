import { AUTH_CONFIG, ACCESS_CONFIG } from '../config/msal-config.js?v=2.3.0';

let msalApp=null;
let account=null;
const tokenCache=new Map();

function normalizeScopes(scopes){
  return [...new Set((scopes?.length?scopes:AUTH_CONFIG.scopes).map(String))].sort();
}
function cacheKey(scopes){return normalizeScopes(scopes).join(' ');}

export function getRedirectUri(){
  return AUTH_CONFIG.redirectUri || (window.location.origin + window.location.pathname);
}

export async function initAuth(){
  if(!window.msal) throw new Error('No se cargó la librería MSAL. Revise el acceso a alcdn.msauth.net.');
  msalApp=new msal.PublicClientApplication({
    auth:{
      clientId:AUTH_CONFIG.clientId,
      authority:`https://login.microsoftonline.com/${AUTH_CONFIG.tenantId}`,
      redirectUri:getRedirectUri()
    },
    cache:{
      cacheLocation:AUTH_CONFIG.cacheLocation || 'sessionStorage',
      storeAuthStateInCookie:false
    }
  });
  if(typeof msalApp.initialize==='function') await msalApp.initialize();
  const accounts=msalApp.getAllAccounts();
  if(accounts.length){
    account=accounts[0];
    try{await acquireToken(AUTH_CONFIG.scopes);}
    catch(err){console.warn('Sesión encontrada, token pendiente de interacción.',err);}
  }
  return account;
}

export function getAccount(){ return account || (msalApp && msalApp.getAllAccounts()[0]) || null; }
export function isAuthenticated(){ return !!getAccount(); }

export function getAuthenticatedEmail(){
  const a=getAccount();
  const c=a?.idTokenClaims||{};
  const candidates=[a?.username,c.preferred_username,c.email,c.upn,c.unique_name];
  const email=candidates.find(v=>typeof v==='string'&&v.includes('@'))||'';
  return email.trim().toLowerCase();
}

export function isGpsAdministrator(){
  const email=getAuthenticatedEmail();
  return !!email && (ACCESS_CONFIG.gpsAdministrators||[]).some(x=>String(x).trim().toLowerCase()===email);
}

export async function login(){
  if(!msalApp) await initAuth();
  const scopes=normalizeScopes(AUTH_CONFIG.scopes);
  const result=await msalApp.loginPopup({scopes});
  account=result.account;
  if(result.accessToken) tokenCache.set(cacheKey(scopes),result.accessToken);
  return account;
}

// Permite consentimiento incremental: lectura para usuarios normales y escritura solo al administrador.
export async function acquireToken(scopes=AUTH_CONFIG.scopes){
  if(!msalApp) await initAuth();
  const active=getAccount();
  if(!active) throw new Error('No existe una cuenta Microsoft autenticada.');
  const normalized=normalizeScopes(scopes);
  try{
    const result=await msalApp.acquireTokenSilent({scopes:normalized,account:active});
    account=result.account||active;
    if(result.accessToken) tokenCache.set(cacheKey(normalized),result.accessToken);
    return result.accessToken;
  }catch(err){
    if(window.msal && err instanceof msal.InteractionRequiredAuthError){
      const result=await msalApp.acquireTokenPopup({scopes:normalized,account:active});
      account=result.account||active;
      if(result.accessToken) tokenCache.set(cacheKey(normalized),result.accessToken);
      return result.accessToken;
    }
    throw err;
  }
}

export async function logout(){
  const active=getAccount();
  tokenCache.clear(); account=null;
  if(msalApp && active){
    await msalApp.logoutPopup({
      account:active,
      postLogoutRedirectUri:getRedirectUri(),
      mainWindowRedirectUri:getRedirectUri()
    });
  }
}

export function authDiagnostics(){
  return {
    tenantId:AUTH_CONFIG.tenantId,
    clientId:AUTH_CONFIG.clientId,
    scopes:[...AUTH_CONFIG.scopes],
    adminScopes:[...(AUTH_CONFIG.adminScopes||[])],
    redirectUri:getRedirectUri(),
    account:getAuthenticatedEmail(),
    isGpsAdministrator:isGpsAdministrator()
  };
}
