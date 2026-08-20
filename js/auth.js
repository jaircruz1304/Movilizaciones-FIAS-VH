import { AUTH_CONFIG } from '../config/msal-config.js?v=1.5.0';

let msalApp=null;
let account=null;
let token=null;

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
  if(accounts.length){ account=accounts[0]; try{await acquireToken();}catch(err){console.warn('Sesión encontrada, token pendiente de interacción.',err);} }
  return account;
}

export function getAccount(){ return account || (msalApp && msalApp.getAllAccounts()[0]) || null; }
export function isAuthenticated(){ return !!getAccount(); }

export async function login(){
  if(!msalApp) await initAuth();
  const result=await msalApp.loginPopup({scopes:AUTH_CONFIG.scopes});
  account=result.account;
  token=result.accessToken || null;
  return account;
}

export async function acquireToken(){
  if(!msalApp) await initAuth();
  const active=getAccount();
  if(!active) throw new Error('No existe una cuenta Microsoft autenticada.');
  try{
    const result=await msalApp.acquireTokenSilent({scopes:AUTH_CONFIG.scopes,account:active});
    account=result.account||active; token=result.accessToken; return token;
  }catch(err){
    if(window.msal && err instanceof msal.InteractionRequiredAuthError){
      const result=await msalApp.acquireTokenPopup({scopes:AUTH_CONFIG.scopes,account:active});
      account=result.account||active; token=result.accessToken; return token;
    }
    throw err;
  }
}

export async function logout(){
  const active=getAccount();
  token=null; account=null;
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
    redirectUri:getRedirectUri(),
    account:getAccount()?.username||''
  };
}
