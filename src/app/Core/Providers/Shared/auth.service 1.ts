import { inject, Injectable } from '@angular/core';
import { msalInstance, loginRequest } from '../../Providers/Shared/msal-config';
import {
  AccountInfo,
  AuthenticationResult,
  PopupRequest,
  BrowserAuthError,
} from '@azure/msal-browser';
import { UserNetworkService } from './global.services';
import { ActivatedRoute, Router } from '@angular/router';
import { environment, UserInfo } from '../../../../environments/environment'

@Injectable({ providedIn: 'root' })
export class AuthServices {
  private account: AccountInfo | null = null;
  private initialized = false;
  private loginInProgress = false;
  public userInfo : any;

  private router = inject(Router);
  
  constructor(private networkService: UserNetworkService, private routers: Router, private route: ActivatedRoute) {
    this.ensureInitialized();
  }

  
  /** Initialize MSAL once and handle any old redirect leftovers */
  private async ensureInitialized(): Promise<void> {
    if (this.initialized) return;
    try {
      await msalInstance.initialize();
      this.initialized = true;
      console.log('✅ MSAL initialized');

      // Handle leftover redirect states (safety only)
      try {
        const response = await msalInstance.handleRedirectPromise();
        if (response && response.account) {
          msalInstance.setActiveAccount(response.account);
          console.log('✅ Previous redirect handled for:', response.account.username);
        }
      } catch (redirectErr) {
        console.warn('⚠️ Redirect cleanup error ignored:', redirectErr);
      }
    } catch (err) {
      console.error('❌ MSAL initialization failed:', err);
    }
  }

  /** Popup login with MFA, no redirects */
  async loginWithUsernameHint(username?: string): Promise<void> {
    if (this.loginInProgress) {
      console.warn('⚠️ Login already in progress — ignoring new attempt');
      return;
    }

    this.loginInProgress = true;
    await this.ensureInitialized();

    const request: PopupRequest = {
      ...loginRequest,
      loginHint: username,
      prompt: 'select_account', // force popup
    };


    console.log('Log Request : ', request);

    try {
      // Clear old redirect state (prevents unwanted redirects)
      await msalInstance.clearCache();

      // Perform popup login
      const result: AuthenticationResult = await msalInstance.loginPopup(request);
      this.account = result.account!;
      msalInstance.setActiveAccount(this.account);

      console.log('✅ Login successful:', this.account);
      this.logSessionIdAndUsername(this.account, username);

      // Optional backend call
      await this.callStatusApi();
    } catch (err) {
      if (err instanceof BrowserAuthError && err.errorCode === 'interaction_in_progress') {
        console.warn('⚠️ Login already running — wait for popup to finish');
      } else {
        console.error('❌ Login popup failed:', err);
      }
    } finally {
      this.loginInProgress = false;
    }
  }

  /** Log session ID + username to console */
  private logSessionIdAndUsername(account: AccountInfo, user : any): void {
    const claims: any = account.idTokenClaims || {};
    const sessionId = claims.sid || claims.oid || account.homeAccountId;
    console.log('🔑 Session ID:', sessionId);
    console.log('👤 Username:', account.username);
  }

  /** Optional: call your backend status API */
  private async callStatusApi(): Promise<void> {
    const account = this.account || msalInstance.getActiveAccount();
    if (!account) return;

    const claims: any = account.idTokenClaims || {};
    const sessionId = claims.sid || claims.oid || account.homeAccountId;
    const username = account.username;
    //const  username= 'axel-testuser-1@tropicalchevrolet.com';
    const groupCode = 'TPG';

    const url = `https://api.axelone.app/api/login/status/${sessionId}/${groupCode}/${username}`;
    console.log('📡 Calling backend:', url);

    try {
      const response = await fetch(url);
      const data = await response.json();
      console.log('✅ API response:', data);


      //If Success redirecting to dealer dashboard
      if(data.status == 200){
        const tkn = data.response
        const userPayload = this.getJwtPayload(tkn); 

        console.log('✅ API response 2:', userPayload);

        var selectedGrp = '';
        if (username.toLowerCase().includes('tropical')) {
          selectedGrp = 'TPG';
        } else {
          selectedGrp = '';
        }


         const userInfo = {
          group : selectedGrp,
          user_aou_AD_userid: username,
          fullName: userPayload.firstname+' '+userPayload.lastname, // Use actual value if available
          GU_URL: data?.groupInfo?.GU_URL || '',
          xtract_url : data?.groupInfo?.XTRACT_URL || '',
          user_Info : userPayload,
          xpertResp : ''
        };


        console.log('User info Obj : ', userInfo);

        this.networkService.setUser(userInfo);
        this.userInfo = userInfo;
        if (localStorage.getItem('userInfo')) {
          localStorage.removeItem('userInfo');
          localStorage.setItem('userInfo', JSON.stringify(userInfo));
        }else{
            localStorage.setItem('userInfo', JSON.stringify(userInfo));
        }

        this.understood();

      }
     

    }catch (error) {
      console.error('❌ Backend call failed:', error);
    }
  }

  /** Logout popup */
  async logout(): Promise<void> {
    const account = this.account || msalInstance.getActiveAccount();
    if (!account) return;

    try {
      await msalInstance.logoutPopup({ account });
      console.log('👋 Logged out successfully');
    } catch (err) {
      console.error('❌ Logout failed:', err);
    }
  }


getJwtPayload(token: string): any {
  const parts = token.split('.');
  if (parts.length !== 3) {
    throw new Error('Invalid JWT format');
  }
  const payload = this.decodeBase64Url(parts[1]);
  return JSON.parse(payload);
}


decodeBase64Url(base64Url: string): string {
  // Replace - with + and _ with / to make it valid Base64
  const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
  // Add padding if needed
  const padded = base64.padEnd(base64.length + (4 - (base64.length % 4)) % 4, '=');
  return atob(padded);
}


 async understood(chk=''){
    var  networkSt = true;
    const returnUrl = this.route.snapshot.queryParamMap.get('returnUrl');
    
    if(this.userInfo != undefined){
      const token = btoa(JSON.stringify(this.userInfo));
      if(this.userInfo?.group === 'TPG'){
        // this.router.navigate(['/dashboard']);

        const resp = await Promise.all([
          this.isReachable('https://tropical.axelone.app')
        ]);
        networkSt = resp[0];

        if(networkSt){
          localStorage.removeItem('userInfo');
          window.location.href = `https://tropical.axelone.app?token=${token}`
        }else{
          this.router.navigate(['/dashboard']);
        }
      }
    }else{
       window.location.href = environment.tropicalAzure.redirectUri;
    }
  }

  //Checking is url reachable or not
  isReachable(url: string, timeout = 3000): Promise<boolean> {
    return new Promise(resolve => {
      const img = new Image();
      let timer = setTimeout(() => {
        img.src = ''; // cancel if taking too long
        resolve(false);
      }, timeout);
        img.onload = () => {
          clearTimeout(timer);
          resolve(true);
        };
        img.onerror = () => {
          clearTimeout(timer);
          resolve(false);
        };
      img.src = `${url}/favicon.ico?ping=${Date.now()}`;
    });
  }

}
