import { environment, UserInfo } from '../../../../environments/environment'
import { BrowserCacheLocation, PublicClientApplication } from '@azure/msal-browser';

export const msalConfig = {
  auth: {
    clientId: environment.tropicalAzure.clientId,
    authority: environment.tropicalAzure.authority,
    // No redirectUri → MSAL uses window.location.origin automatically
  },
  cache: {
    cacheLocation: 'sessionStorage' as BrowserCacheLocation,
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  scopes: ['openid', 'profile', 'email'],
};

export const msalInstance = new PublicClientApplication(msalConfig);
