import { HttpClient, HttpHeaders } from '@angular/common/http';
import { DOCUMENT, Inject, Injectable } from '@angular/core';
import {  Observable, throwError } from 'rxjs';
import { catchError, filter, map} from 'rxjs/operators';
import { Router, NavigationEnd } from '@angular/router';
// import { DOCUMENT } from '@angular/core';
// import { DOCUMENT } from '@angular/common';

@Injectable({
  providedIn: 'root'
})
export class UserActivityService {
  userID: any = '';
  private baseUrl = 'https://fbccapi.axelone.app/api/';


  constructor(private http: HttpClient, private router: Router,@Inject(DOCUMENT) private document: Document) {
  }

  getAllLocalStorage(): Record<string, string | null> {
    const result: Record<string, string | null> = {};
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key) {
        result[key] = localStorage.getItem(key);
      }
    }
    return result;
  }

  /* ---------------- CLIENT INFO ---------------- */
  private getClientInfo() {
    const ua = navigator.userAgent;

    const os =
      /Windows NT/.test(ua) ? 'Windows' :
        /Mac OS X/.test(ua) ? 'MacOS' :
          /Android/.test(ua) ? 'Android' :
            /iPhone|iPad/.test(ua) ? 'iOS' :
              'Unknown';

    const browser =
      /Chrome/.test(ua) && !/Edg/.test(ua) ? 'Chrome' :
        /Edg/.test(ua) ? 'Edge' :
          /Firefox/.test(ua) ? 'Firefox' :
            /Safari/.test(ua) ? 'Safari' :
              'Unknown';

    const device =
      /Mobi|Android/i.test(ua) ? 'Mobile' :
        /Tablet|iPad/i.test(ua) ? 'Tablet' :
          'Desktop';

    return { os, browser, device };
  }

  /* ---------------- IP ---------------- */
  private async getIp(): Promise<string> {
    const res = await fetch('https://api.ipify.org?format=json');
    return (await res.json()).ip;
  }

startRouteTracking() {
    this.router.events
      .pipe(filter(event => event instanceof NavigationEnd))
      .subscribe((event: NavigationEnd) => {
        const path = this.document.location.pathname;
        let section: any = path.split('/').filter(Boolean)[0];
        if (section == undefined) {
          section = '';
        }
        console.log(path);
        console.log(section);
       
        // console.log('modname', section);
        // console.log("User_Activity", event);
        const host = window.location.hostname;
        if (!host.includes('localhost') && !host.includes('127.0.0.1')) {
          this.trackScreen(event.urlAfterRedirects, section);
        }
 
      });
  }
  //////////// REFRESH TOKEN LOGIC ////////////

  async trackScreen(pageName: string, modname: string) {
    const remoteToken: any = localStorage.getItem('cdpRemoteBarerToken');
    const client = this.getClientInfo();
    const ip = await this.getIp();

    var currentDate = new Date();
    let timeStamp =
      ((currentDate.getMonth() + 1) < 10 ? "0" + (currentDate.getMonth() + 1) : (currentDate.getMonth() + 1)) + "-" + ((currentDate.getDate()) < 10 ? "0" + (currentDate.getDate()) : (currentDate.getDate())) + "-" + currentDate.getFullYear() + " " +
      currentDate.getHours() + ':' + currentDate.getMinutes() + ':' + currentDate.getSeconds();
    // console.log(timeStamp);

    const payload = {
      pagename: pageName,
      modname: modname,
      device: client.device,
      os: client.os,
      browser: client.browser,
      ip,
      screen: `${screen.width}x${screen.height}`,
      language: navigator.language,
      timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
      clientts: timeStamp,
    };

    const loginUserId = this.getLoginUserId();
    const bearerUserId = this.getUserIdFromToken(remoteToken);

    if (loginUserId && bearerUserId && loginUserId !== bearerUserId) {
      console.log('User mismatch — refreshing token');
      this.refreshToken().subscribe(() => this.sendTrackRequest(payload));
    } else {
      this.sendTrackRequest(payload);
    }
  }

  private getLoginUserId(): string | null {
    // const host = window.location.hostname;
    //  if (host.includes('localhost') || host.includes('127.0.0.1')) {
    //     return `eyJncm91cCI6IldFU1RFUk4iLCJ1c2VyX2FvdV9BRF91c2VyaWQiOiJwcmFzYWQuY2hhdmFsaUBheGVsYXV0b21vdGl2ZS5jb20iLCJmdWxsTmFtZSI6IlByYXNhZCBDaGF2YWxpIiwiR1VfVVJMIjoiaHR0cHM6Ly93ZXN0ZXJuYXV0by5heGVsb25lLmFwcCIsInh0cmFjdF91cmwiOiJodHRwczovL3dlc3Rlcm54dHJhY3QuYXhlbG9uZS5hcHAiLCJ1c2VyX0luZm8iOnsidXNlcmlkIjoxLCJmaXJzdG5hbWUiOiJQcmFzYWQiLCJsYXN0bmFtZSI6IkNoYXZhbGkiLCJlbWFpbCI6InByYXNhZC5jaGF2YWxpQGF4ZWxhdXRvbW90aXZlLmNvbSIsInBob25lIjoiMjEwNDg5MTQwMSIsInJvbGVpZCI6MTAwLCJzdGF0dXMiOiJZIiwicHJvZmlsZUltZyI6IiIsInRpdGxlIjoiU3VwZXIgQWRtaW4iLCJtb2R1bGVpZHMiOiIxLDIsNCw1LDcsOCwxMSwxMiwxNSwxNiwxOCwxOSwyMiwyMywyNCwyNSwyNywyOCwyOSwzMCwzNSwzNiw0MCw0MSw0OCw1Myw2Miw2Myw2NiwyNiw2OCw2OSw3MCw3MSw3Miw3Myw4MCw4MSw5NSw5Nyw5OSwxMDMsMTQsNTQsNzQsNzYsNzcsNzgsNzksODcsODgsODksOTAsOTEsOTIsOTMsOTQsOTYsOTgsMTAwLDEwMSw2Nyw3NSw4Myw4NCw4NSwxMDIsMTA0IiwidXN0b3JlcyI6IjEsNCwxNCw0Miw3MSIsIkFEdXNlcmlkIjoicHJhc2FkLmNoYXZhbGlAYXhlbGF1dG9tb3RpdmUuY29tIiwidWlkIjoxLCJYdHJhY3QiOjEsIlRvdWNoIjoxLCJYcGVyaWVuY2UiOjEsIlhjaGFuZ2UiOjMsIlhpb20iOjEsIlRyYWNzIjoxLCJpYXQiOjE3NjUxOTA4MjgsImV4cCI6MTc5NjcyNjgyOCwiZ3JvdXAiOiJXRVNURVJOIn0sInhwZXJ0UmVzcCI6IiIsInNpdGUiOiJwcm9kIn0`;
    //   }
    try {
      const storage: any = this.getAllLocalStorage();
      const token = storage?.finalUser || btoa(JSON.stringify(storage?.userInfo));
      if (!token) return null;

      const data = JSON.parse(atob(token));
      return data?.user_Info?.userid || null;
    } catch {
      return null;
    }
  }

  private getUserIdFromToken(token: string | null): string | null {
    // const host = window.location.hostname;
    // if (host.includes('localhost') || host.includes('127.0.0.1')) {
    //     return `eyJncm91cCI6IldFU1RFUk4iLCJ1c2VyX2FvdV9BRF91c2VyaWQiOiJwcmFzYWQuY2hhdmFsaUBheGVsYXV0b21vdGl2ZS5jb20iLCJmdWxsTmFtZSI6IlByYXNhZCBDaGF2YWxpIiwiR1VfVVJMIjoiaHR0cHM6Ly93ZXN0ZXJuYXV0by5heGVsb25lLmFwcCIsInh0cmFjdF91cmwiOiJodHRwczovL3dlc3Rlcm54dHJhY3QuYXhlbG9uZS5hcHAiLCJ1c2VyX0luZm8iOnsidXNlcmlkIjoxLCJmaXJzdG5hbWUiOiJQcmFzYWQiLCJsYXN0bmFtZSI6IkNoYXZhbGkiLCJlbWFpbCI6InByYXNhZC5jaGF2YWxpQGF4ZWxhdXRvbW90aXZlLmNvbSIsInBob25lIjoiMjEwNDg5MTQwMSIsInJvbGVpZCI6MTAwLCJzdGF0dXMiOiJZIiwicHJvZmlsZUltZyI6IiIsInRpdGxlIjoiU3VwZXIgQWRtaW4iLCJtb2R1bGVpZHMiOiIxLDIsNCw1LDcsOCwxMSwxMiwxNSwxNiwxOCwxOSwyMiwyMywyNCwyNSwyNywyOCwyOSwzMCwzNSwzNiw0MCw0MSw0OCw1Myw2Miw2Myw2NiwyNiw2OCw2OSw3MCw3MSw3Miw3Myw4MCw4MSw5NSw5Nyw5OSwxMDMsMTQsNTQsNzQsNzYsNzcsNzgsNzksODcsODgsODksOTAsOTEsOTIsOTMsOTQsOTYsOTgsMTAwLDEwMSw2Nyw3NSw4Myw4NCw4NSwxMDIsMTA0IiwidXN0b3JlcyI6IjEsNCwxNCw0Miw3MSIsIkFEdXNlcmlkIjoicHJhc2FkLmNoYXZhbGlAYXhlbGF1dG9tb3RpdmUuY29tIiwidWlkIjoxLCJYdHJhY3QiOjEsIlRvdWNoIjoxLCJYcGVyaWVuY2UiOjEsIlhjaGFuZ2UiOjMsIlhpb20iOjEsIlRyYWNzIjoxLCJpYXQiOjE3NjUxOTA4MjgsImV4cCI6MTc5NjcyNjgyOCwiZ3JvdXAiOiJXRVNURVJOIn0sInhwZXJ0UmVzcCI6IiIsInNpdGUiOiJwcm9kIn0`;
    //   }
    try {
      if (!token) return null;
      const payload = token.split('.')[1];
      const base64 = payload.replace(/-/g, '+').replace(/_/g, '/');
      const json = atob(base64.padEnd(base64.length + (4 - base64.length % 4) % 4, '='));
      return JSON.parse(json)?.userid || null;
    } catch {
      return null;
    }
  }

  private sendTrackRequest(payload: any) {
    this.http.post(`${this.baseUrl}cc/useractivity/trackuserscreen`, payload, {
      headers: this.getAuthHeaders()
    }).subscribe({
      next: res => console.log('User Activity Tracked', res),
      error: () => {
        console.log('Request failed — refreshing token');
        this.refreshToken().subscribe({
          next: () => {
            this.http.post(`${this.baseUrl}cc/useractivity/trackuserscreen`, payload, {
              headers: this.getAuthHeaders()
            }).subscribe({
              next: res => console.log('Retry success', res),
              error: err => console.error('Retry failed', err)
            });
          },
          error: err => console.error('Refresh token failed', err)
        });
      }
    });
  }


  private getAuthHeaders(): HttpHeaders {
    const token = localStorage.getItem('cdpRemoteBarerToken');
    return new HttpHeaders({
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      ...(token ? { Authorization: `Bearer ${token}` } : {})
    });
  }

  // 🔹 Refresh Token Logic
  loginToken: any = '';
  private refreshToken(): Observable<string> {
    console.log('refresh token called');
    const host = window.location.hostname;
    if (!host.includes('localhost') && !host.includes('127.0.0.1')) {
      //   this.loginToken = `eyJncm91cCI6IldFU1RFUk4iLCJ1c2VyX2FvdV9BRF91c2VyaWQiOiJwcmFzYWQuY2hhdmFsaUBheGVsYXV0b21vdGl2ZS5jb20iLCJmdWxsTmFtZSI6IlByYXNhZCBDaGF2YWxpIiwiR1VfVVJMIjoiaHR0cHM6Ly93ZXN0ZXJuYXV0by5heGVsb25lLmFwcCIsInh0cmFjdF91cmwiOiJodHRwczovL3dlc3Rlcm54dHJhY3QuYXhlbG9uZS5hcHAiLCJ1c2VyX0luZm8iOnsidXNlcmlkIjoxLCJmaXJzdG5hbWUiOiJQcmFzYWQiLCJsYXN0bmFtZSI6IkNoYXZhbGkiLCJlbWFpbCI6InByYXNhZC5jaGF2YWxpQGF4ZWxhdXRvbW90aXZlLmNvbSIsInBob25lIjoiMjEwNDg5MTQwMSIsInJvbGVpZCI6MTAwLCJzdGF0dXMiOiJZIiwicHJvZmlsZUltZyI6IiIsInRpdGxlIjoiU3VwZXIgQWRtaW4iLCJtb2R1bGVpZHMiOiIxLDIsNCw1LDcsOCwxMSwxMiwxNSwxNiwxOCwxOSwyMiwyMywyNCwyNSwyNywyOCwyOSwzMCwzNSwzNiw0MCw0MSw0OCw1Myw2Miw2Myw2NiwyNiw2OCw2OSw3MCw3MSw3Miw3Myw4MCw4MSw5NSw5Nyw5OSwxMDMsMTQsNTQsNzQsNzYsNzcsNzgsNzksODcsODgsODksOTAsOTEsOTIsOTMsOTQsOTYsOTgsMTAwLDEwMSw2Nyw3NSw4Myw4NCw4NSwxMDIsMTA0IiwidXN0b3JlcyI6IjEsNCwxNCw0Miw3MSIsIkFEdXNlcmlkIjoicHJhc2FkLmNoYXZhbGlAYXhlbGF1dG9tb3RpdmUuY29tIiwidWlkIjoxLCJYdHJhY3QiOjEsIlRvdWNoIjoxLCJYcGVyaWVuY2UiOjEsIlhjaGFuZ2UiOjMsIlhpb20iOjEsIlRyYWNzIjoxLCJpYXQiOjE3NjUxOTA4MjgsImV4cCI6MTc5NjcyNjgyOCwiZ3JvdXAiOiJXRVNURVJOIn0sInhwZXJ0UmVzcCI6IiIsInNpdGUiOiJwcm9kIn0`;
      // } else {
      const CCcaptureObject: any = this.getAllLocalStorage();
      const produserInfo = CCcaptureObject?.userInfo;
      console.log('Json',CCcaptureObject?.userInfo);
      const prodObjToken = btoa(JSON.stringify(produserInfo));
      console.log('prodObjToken', prodObjToken);
      const prodToken = CCcaptureObject?.finalUser;
      console.log('prodToken', prodToken);
      if (prodToken) {
        this.loginToken = prodToken;
      } else if (prodObjToken) {
        this.loginToken = prodObjToken;
      }      
    }
    console.log('login Token', this.loginToken);
    let httpOptions: any = {
      headers: new HttpHeaders({
        'Authorization': `Bearer ${this.loginToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      })
    };
    console.log('httpOptions', httpOptions);
    return this.http.post<{ token: string }>(`${this.baseUrl}secure/exchangeremotetoken`, {}, httpOptions).pipe(
      map((res: any) => {
        if (res.status == 200) {
          const newToken = res.token;
          localStorage.setItem('cdpRemoteBarerToken', newToken);
          return newToken;
        }

        // else {
        //   localStorage.clear();
        //   sessionStorage.clear();
        //   window.location.href = `https://axelone.app/`;
        //   //logout
        // }
      }),
      catchError((err) => {
        // localStorage.clear();
        // sessionStorage.clear();
        // window.location.href = `https://axelone.app/`;        //logout
        // console.error('Error refreshing token:', err);
        return throwError(() => err);
      })
    );
  }

}
