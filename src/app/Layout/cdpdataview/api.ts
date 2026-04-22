import { Injectable } from '@angular/core';
import { HttpClient, HttpErrorResponse, HttpHeaders } from '@angular/common/http';
import { BehaviorSubject, Observable, throwError } from 'rxjs';
import { tap, map, catchError, switchMap, filter, take } from 'rxjs/operators';

@Injectable({
  providedIn: 'root'
})
export class Api {
  private baseUrl = 'https://fbccapi.axelone.app/api/';
  fontUrl = `${this.baseUrl}resources/fonts/`;
  constructor(private http: HttpClient) { }

  ////////========== TOKEN AUTHENTICATION METHODS ======/////////

  //  username = JSON.parse(localStorage.getItem('userInfo')!)?.user_Info?.ADuserid;
  username = 'fbcc'
  password = 'fbcc@2k25#';
  base64Credentials = btoa(`${this.username}:${this.password}`);
  private httpOptions1 = {
    headers: new HttpHeaders({
      'Authorization': `Basic ${this.base64Credentials}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    })
  };

  post(url: string, data: any): Observable<any> {
    console.log(this.username, this.base64Credentials, 'UserName');

    return this.http.post(`${this.baseUrl}${url}`, data, this.httpOptions1);
  }


}
