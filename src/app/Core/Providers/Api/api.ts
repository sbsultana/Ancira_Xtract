import { Injectable } from '@angular/core';
import { BehaviorSubject, Observable } from 'rxjs';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { environment } from '../../../../environments/environment';
import { map } from 'rxjs/operators';



@Injectable({
  providedIn: 'root'
})
export class Api {
  private HeaderData = new BehaviorSubject<any>({
    obj: '',
  });
  private ExportToPDFResponse = new BehaviorSubject<any>({
    obj: '',
  });
  private ReportOpen = new BehaviorSubject<any>({
    obj: '',
  });
  private Reports = new BehaviorSubject<any>({
    obj: '',
  });

  private ExportToExcelAllReports = new BehaviorSubject<any>({
    obj: '',
  });
  private ExportToPDFAllReports = new BehaviorSubject<any>({
    obj: '',
  });
  private ExportToEmailPDFAllReports = new BehaviorSubject<any>({
    obj: '',
  });
  private ExportToPrintAllReports = new BehaviorSubject<any>({
    obj: '',
  });
  private Comments = new BehaviorSubject<any>({
    obj: '',
  });
  private stores = new BehaviorSubject<any>({
    obj: '',
  });
  private reconstores = new BehaviorSubject<any>({
    obj: '',
  });
  private accouningreconstores = new BehaviorSubject<any>({
    obj: '',
  });
  private storesVariables = new BehaviorSubject<any>({
    obj: '',
  })
  private storesAll = new BehaviorSubject<any>({
    obj: '',
  });

  private ReconServiceLaborStores = new BehaviorSubject<any>({
    obj: '',
  });
  private mobileservicestrs = new BehaviorSubject<any>({
    obj: '',
  });

  constructor(public http: HttpClient) { }

  postmethod(endpoint: string, obj: object): Observable<any> {
    return this.http.post(`${environment.apiUrl}${endpoint}`, obj).pipe(
      map((res: any) => {
        return res;
      })
    );
  }
  putmethod(endpoint: string, obj: object): Observable<any> {
    return this.http.put(`${environment.apiUrl}${endpoint}`, obj).pipe(
      map((res: any) => {
        return res;
      })
    );
  }

  deletemethod(endpoint: string, obj: object): Observable<any> {
    return this.http
      .request('delete', `${environment.apiUrl}${endpoint}`, { body: obj })
      .pipe(
        map((res: any) => {
          // console.log(res);
          return res;
        })
      );
  }
  postmethodOne(endpoint: string, obj: object): Observable<any> {
    return this.http.post(`${environment.axelOneUrl}${endpoint}`, obj).pipe(
      map((res: any) => {
        return res;
      })
    );
  }

  postmethodDev2(endpoint: string, obj: object): Observable<any> {
    return this.http.post(`${environment.apiDev2}${endpoint}`, obj).pipe(
      map((res: any) => {
        return res;
      })
    );
  }

  postmethodXpert(endpoint: string, obj: object): Observable<any> {
    return this.http.post(`${environment.xpert}${endpoint}`, obj).pipe(
      map((res: any) => {
        return res;
      })
    );
  }

  // postSupportmethod(endpoint: string, obj: object): Observable<any> {
  //   let httpOptions: any;

  //   httpOptions = {
  //     headers: new HttpHeaders({
  //       "Accept": "application/json"
  //     })
  //   }
  //   return this.http.post(endpoint, obj, httpOptions)
  //     .pipe(map((res: any) => {
  //       return res;
  //     }))
  // }

  checkAuthLogin() {
    let authToken = localStorage.getItem("userobj");
    if (authToken) {
      return true;
    } else {
      return false;
    }
  }


  getNotificationsUnreadedCount(endpoint: any): Observable<any> {
    //var endpoint = '/fcm/getNotifications?user_to_aou_id=1&user_group_code=CAV1';
    return this.http.get(`${environment.apiUrl}${endpoint}`);
  }




  public getBroswer() {
    const agent = window.navigator.userAgent.toLowerCase();
    const browser =
      agent.indexOf('edge') > -1 ? 'Microsoft Edge'
        : agent.indexOf('edg') > -1 ? 'Chromium-based Edge'
          : agent.indexOf('opr') > -1 ? 'Opera'
            : agent.indexOf('chrome') > -1 ? 'Chrome'
              : agent.indexOf('trident') > -1 ? 'Internet Explorer'
                : agent.indexOf('firefox') > -1 ? 'Firefox'
                  : agent.indexOf('safari') > -1 ? 'Safari'
                    : 'other';

    return browser
  }

  private payrollSubject = new BehaviorSubject<string>('N');
  payroll$ = this.payrollSubject.asObservable();

  setPayroll(value: string) {
    this.payrollSubject.next(value);
    localStorage.setItem('Payroll', value);
  }

  SetHeaderData(data: any) {
    this.HeaderData.next(data);
  }
  GetHeaderData() {
    return this.HeaderData.asObservable();
  }

  SetReportOpening(data: any) {
    this.ReportOpen.next(data);
  }
  GetReportOpening() {
    return this.ReportOpen.asObservable();
  }
  SetReports(data: any) {
    this.Reports.next(data);
  }
  GetReports() {
    return this.Reports.asObservable();
  }

  setExportToExcelAllReports(data: any) {
    this.ExportToExcelAllReports.next(data);
  }
  getExportToExcelAllReports() {
    return this.ExportToExcelAllReports.asObservable();
  }
  setExportToPrintAllReports(data: any) {
    this.ExportToPrintAllReports.next(data);
  }
  getExportToPrintAllReports() {
    return this.ExportToPrintAllReports.asObservable();
  }
  setExportToPDFAllReports(data: any) {
    this.ExportToPDFAllReports.next(data);
  }
  getExportToPDFAllReports() {
    return this.ExportToPDFAllReports.asObservable();
  }
  setExportToPDFResponse(data: any) {
    this.ExportToPDFResponse.next(data);
  }
  getExportToPDFResponse() {
    return this.ExportToPDFResponse.asObservable();
  }
  setExportToEmailPDFAllReports(data: any) {
    this.ExportToEmailPDFAllReports.next(data);
  }
  getExportToEmailPDFAllReports() {
    return this.ExportToEmailPDFAllReports.asObservable();
  }
  setComments(data: any) {
    this.Comments.next(data);
  }
  getComments() {
    return this.Comments.asObservable();
  }
  setStores(data: any) {
    this.stores.next(data);
  }
  getStores() {
    return this.stores.asObservable();
  }

  getIncomeTargets() {
    return this.storesVariables.asObservable();
  }

  setAllStores(data: any) {
    this.reconstores.next(data);
  }
  getAllStores() {
    return this.reconstores.asObservable();
  }

  setStoresAll(data: any) {
    this.storesAll.next(data);
  }
  getStoresAll() {
    return this.storesAll.asObservable();
  }
  setAccuntingAllStores(data: any) {
    this.accouningreconstores.next(data);
  }
  getAccountingAllStores() {
    return this.accouningreconstores.asObservable();
  }
  setReconServiceLaborStores(data: any) {
    this.ReconServiceLaborStores.next(data);
  }
  getReconServiceLaborStores() {
    return this.ReconServiceLaborStores.asObservable();
  }

  setMobileService(data: any) {
    this.mobileservicestrs.next(data);
  }
  getMobileService() {
    return this.mobileservicestrs.asObservable();
  }

  headermenu(roleId: any) {
    var httpOptions = {
      headers: new HttpHeaders({
        "Accept": "application/json"
      })
    }

    const data = {
      id: roleId
    }
    return this.http.post('https://ancinsightsapi.axelone.app/api/cmsmodules/GetCmsPermissionsMenu', data, httpOptions)
  }

  headsource = new BehaviorSubject(false);
  headval = this.headsource.asObservable();

  changehead(data: any) {
    this.headsource.next(data);
  }

  getsessionid(userid: any, url: any) {
    var httpOptions = {
      headers: new HttpHeaders({
        "Accept": "application/json"
      })
    }

    let body = { "userid": userid }
    return this.http.post(url, body, httpOptions);
  }
  getMethod(endpoint: any) {
    return this.http.get(`${environment.apiUrl1}${endpoint}`);
  }

  postmethodtoken(endpoint: string, obj: object): Observable<any> {
    return this.http.post(`${environment.apiconn}${endpoint}`, obj).pipe(
      map((res: any) => {
        return res;
      })
    );
  }

  private userTokenSource = new BehaviorSubject<any>(null);
  userToken$ = this.userTokenSource.asObservable();

  setUserToken(data: any) {
    this.userTokenSource.next(data);
  }


  postmethodHRDB(endpoint: string, obj: object): Observable<any> {
    return this.http.post(`${environment.apiconn}${endpoint}`, obj).pipe(
      map((res: any) => {
        return res;
      })
    );
  }


}
