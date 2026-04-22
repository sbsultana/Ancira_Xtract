import { Component, ElementRef, EventEmitter, HostListener, inject, Input, Output, PLATFORM_ID, Renderer2, ViewChild } from '@angular/core';
import { Sharedservice } from '../../Core/Providers/Shared/sharedservice';
import { Setdates } from '../../Core/Providers/SetDates/setdates';
import { ActivatedRoute, Router, NavigationStart, RouterModule, } from '@angular/router';
import { decode } from 'js-base64';
import { isPlatformBrowser } from '@angular/common';
import { SharedModule } from '../../Core/Providers/Shared/shared.module';
import { UserNetworkService } from '../../Core/Providers/Shared/global.services';


import { environment, UserInfo } from '../../../environments/environment';
import { NgFor, NgIf } from '@angular/common';
import { NgStyle, NgClass, NgForOf } from '@angular/common';
Api
ToastService
import { HttpClient } from '@angular/common/http';
NotificationService
import { RivIcon } from '../..//Core/Providers/Shared/riv-icon/riv-icon';
import { Api } from '../../Core/Providers/Api/api';
import { ToastService } from '../../Core/Providers/Shared/toast.service';
import { NotificationService } from '../../Core/Providers/Shared/notification.service';

import { AuthServices } from '../../Core/Providers/Shared/auth.service 1';
import { common } from '../../common';
import { Subscription } from 'rxjs';
import { PdfExportService } from '../../Core/Providers/Shared/export.service';
import { NgxSpinnerService } from 'ngx-spinner';

declare var bootstrap: any;
@Component({
  selector: 'app-header',
  imports: [SharedModule, RivIcon, NgIf, RouterModule],
  standalone: true,
  templateUrl: './header.html',
  styleUrl: './header.scss'
})


export class Header {
  @Output() ready = new EventEmitter<void>();

  componentData: any;
  // userInfo: any;
  hasUnread = false;
  matched = false;
  loading = false;
  selectedGroup: string | null = null;
  isDownload: any = true
  private readonly toast = inject(ToastService);
  private readonly router = inject(Router);
  private readonly http = inject(HttpClient);
  private readonly apiCall = inject(Api);
  private readonly shared = inject(Sharedservice);

  public readonly userNetworkInfo = inject(UserNetworkService);
  public readonly notificationService = inject(NotificationService);

  private readonly KeyWords = ['Prasad', 'Chavali', 'Pchavali'];

  private processingToken = false; // ✅ Prevent premature logout
  private sub!: Subscription;
  userData: any;
  constructor(private exportService: PdfExportService, private spinner: NgxSpinnerService) {
    this.apiCall.GetHeaderData().subscribe((res: any) => {
      this.componentData = res.obj;
      this.shared.common.pageName = this.componentData.title;
      this.exportFile()
      console.log(this.shared.common.pageName, 'Page Name');

    });
  }
  userInfo:any={
    "group": "ANCIRA",
    "user_aou_AD_userid": "prasad.chavali@axelautomotive.com",
    "fullName": "Prasad Chavali",
    "GU_URL": "https://anicra.axelautomotive.com",
    "xtract_url": "https://ancxtract.axelautomotive.com/",
    "user_Info": {
        "id": 1,
        "userid": 1,
        "firstname": "Prasad",
        "lastname": "Chavali",
        "email": "prasad.chavali@axelautomotive.com",
        "phone": "9999999999",
        "roleid": 100,
        "status": "Y",
        "profileImg": "",
        "title": "Super Admin",
        "Storeids":"7,5,8,9,204,6,107,104,10,4,13,12",
        "ustores": "7,5,8,9,204,6,107,104,10,4,13,12",
        "ADuserid": "prasad.chavali@axelautomotive.com",
        "Preferences": 1,
        "iat": 1774938969,
        "exp": 1806474969
    },
    "xpertResp": "",
    "site": "prod"
}
  ngOnInit(): void {
    if (window.location.pathname.startsWith('/schedules')) {
      return;
    }
    history.pushState(null, '', location.href);
    window.onpopstate = () => {
      history.pushState(null, '', location.href);
    };
    this.sub = this.apiCall.userToken$
      .subscribe(data => {
        if (data) {
          this.userData = data;
        }
      });

    localStorage.setItem('userInfo', JSON.stringify(this.userInfo))

    const match = window.location.href.match(/[?&]token=([^&#]+)/i);

    if (match) {
      console.log(match);
      const loadingTkn = match?.[1] ? decodeURIComponent(match[1]) : null!;
      const cleared = loadingTkn?.replaceAll('%253D', '%3D');
      const decoded = JSON.parse(atob(cleared));
      localStorage.setItem('flag', decoded.flag ? decoded.flag : 'M')
      localStorage.setItem('stime', decoded.Type ? decoded.Type : 'MTD')
      // decoded.user_Info = decoded.user_Info.user_Info
      this.shared.common.otherstoreids = []
      this.shared.common.otherstoreids = decoded?.other_stores ? decoded?.other_stores.split(",").map(Number) : []
      localStorage.setItem('otherstoreids', JSON.stringify(this.shared.common.otherstoreids))
      console.log(decoded, '................ Decoded', this.shared.common.otherstoreids);
      if (localStorage.getItem('flag') == 'V') {
        let UserData = JSON.parse(localStorage.getItem('userInfo')!);
        UserData.WidgetData = decoded
        localStorage.setItem('userInfo', JSON.stringify(UserData))
        // localStorage.setItem('WidgetData', JSON.stringify(decoded))
      }
      // localStorage.setItem('userInfo', JSON.stringify(decoded))
    } else {
      console.log(match);
      console.log(JSON.parse(localStorage.getItem('userInfo')!), '..................... Without token')
    }
    const storedUserInfo = localStorage.getItem('userInfo');

    if (storedUserInfo) {
      this.userInfo = JSON.parse(storedUserInfo);
      // ✅ Clean up token in URL
      let cleanUrl = window.location.href;
      cleanUrl = cleanUrl.replace(/[?&]token=[^&#]*/g, '').replace(/[?&]$/, '');

      if (cleanUrl.includes('#')) {
        const [base, hash] = cleanUrl.split('#');
        let cleanHash = hash.replace(/[?&]token=[^&#]*/g, '').replace(/[?&]$/, '');
        cleanUrl = base + (cleanHash ? `#${cleanHash}` : '');
      }

      window.history.replaceState({}, document.title, cleanUrl);

      console.log(JSON.parse(localStorage.getItem('userInfo')!), '..................... Without token')

      const adUserId = this.userInfo?.user_Info?.ADuserid || '';
      const isAxelEmp = adUserId.toLowerCase().endsWith('srikanth.bobba@axelautomotive.com');
      if (isAxelEmp) {
        this.getGruopsandStoresAxelOne()
      } else {
        this.getGruopsandStores();
      }
      this.afterUserLoad();
      this.getGruopsandStoresAll();
      // this.getMobileStores();
      // this.getAccountingReconStores()
      // this.getReconStores();
      // this.getStoreByGroupUserIdRecon();
      this.getUsersData()
      this.ready.emit();

      return;
    }

    const token = this.getTokenFromUrl();

    if (token) {
      this.processingToken = true;
      this.processToken(token);
      setTimeout(() => {
        this.processingToken = false;
        this.ready.emit();
      }, 800);
      return;
    }

    this.checkUserInfo();

    this.router.events.subscribe((event) => {
      if (event instanceof NavigationStart && !this.processingToken) {
        this.checkUserInfo();
      }
    });

    console.log('Header Component Data : ', this.componentData.title);
  }
  ngOnDestroy() {
    this.sub?.unsubscribe();
  }
  /** ✅ Extract token from URL if exists */
  private getTokenFromUrl(): string | null {
    //debuger;
    try {
      const match = window.location.href.match(/[?&]token=([^&#]+)/i);
      return match?.[1] ? decodeURIComponent(match[1]) : null;
    } catch (err) {
      return null;
    }
  }

  /** ✅ Decode token, store user info, and remove token from URL */
  private processToken(token: string): void {
    const localToken = localStorage.getItem('Login_token');
    if (!localToken) localStorage.setItem('Login_token', token);

    try {
      // const decoded = JSON.parse(atob(token));
      //    localStorage.setItem('userInfo', JSON.stringify(decoded));

      // this.userInfo = decoded;

      this.afterUserLoad();

      // ✅ Clean up token in URL
      let cleanUrl = window.location.href;
      cleanUrl = cleanUrl.replace(/[?&]token=[^&#]*/g, '').replace(/[?&]$/, '');

      if (cleanUrl.includes('#')) {
        const [base, hash] = cleanUrl.split('#');
        let cleanHash = hash.replace(/[?&]token=[^&#]*/g, '').replace(/[?&]$/, '');
        cleanUrl = base + (cleanHash ? `#${cleanHash}` : '');
      }

      window.history.replaceState({}, document.title, cleanUrl);
    } catch (e) {
    }
  }

  /** ✅ After user info available */
  private afterUserLoad(): void {
    this.evaluateUserMatch();
    this.loadHeaderData();
  }

  /** ✅ Keyword match */
  private evaluateUserMatch(): void {
    const info = this.userInfo?.user_Info || {};
    const fullName = (this.userInfo?.fullName || '').toLowerCase();
    const first = (info?.firstname || '').toLowerCase();
    const last = (info?.lastname || '').toLowerCase();

    this.matched = this.KeyWords.some(
      (k) =>
        fullName.includes(k.toLowerCase()) ||
        first.includes(k.toLowerCase()) ||
        last.includes(k.toLowerCase())
    );
  }

  /** ✅ Fetch header data */
  private loadHeaderData(): void {
    this.apiCall.GetHeaderData().subscribe((res: any) => (this.componentData = res.obj));
  }
  private checkUserInfo(): void {

    if (window.location.pathname.startsWith('/schedules')) {
      return;
    }

    if (this.processingToken) return;

    const stored = localStorage.getItem('userInfo');

    let user: any = null;

    if (stored) {
      try {
        user = JSON.parse(stored);
      } catch {
        user = null;
      }
    }

    if (!user || !Object.keys(user).length) {

      const redirectUrl = 'https://axelone.app/?logout=true';

      window.location.replace(redirectUrl);
      return;
    }

    this.getNotificationCount();
  }


  /** ✅ Notifications */
  private getNotificationCount(): void {
    this.notificationService.unreadCount$.subscribe((count) => (this.hasUnread = count > 0));
  }

  onNotificationIconClick(): void {
    this.notificationService.triggerFetch();
    this.toast.show('Notifications will be implemented soon', 'success', 'Coming Soon');
  }

  /** ✅ Settings navigation */
  onSettingsClick(): void {
    this.apiCall.headermenu(`${this.userInfo.user_Info.roleid}`).subscribe({
      next: (data: any) => {
        const settings = data?.response?.find(
          (m: any) => m.mod_name === 'Settings' || m.mod_name === 'SETTINGS' || m.Mod_ID == 26
        );
        if (!settings) return console.error('SETTINGS module not found');

        try {
          const xmlDataArray = JSON.parse(settings.xmlData);
          const firstItem = xmlDataArray[0];
          if (firstItem) this.gotoLink(firstItem.mod_filename);
        } catch (e) {
          console.error('Error parsing xmlData:', e);
        }
      },
      error: (err) => console.error('API Error:', err),
    });
  }
  gotoLink(url: string): void {
    if (!url) return;

    try {
      const uToken = localStorage.getItem('Login_token') || '';
      const currentHost = window.location.host;
      const isLocal = /localhost|127\.0\.0\.1/.test(currentHost);
      const target = new URL(url, window.location.origin);

      target.searchParams.set('token', uToken);

      if (target.host === currentHost || isLocal) {
        this.router.navigateByUrl(target.pathname + target.search).catch(() => {
          window.location.href = target.href;
        });
      } else {
        const currentBase = currentHost.split('.').slice(-2).join('.');
        const targetBase = target.host.split('.').slice(-2).join('.');
        targetBase === currentBase
          ? (window.location.href = target.href)
          : (window.location.href = target.href);
      }
    } catch (error) {
      console.error('Invalid URL:', error, 'for input:', url);
    }
  }

  dashbardRoute(): void {
    // this.router.navigate(['/dashboard']);
    window.location.href = 'https://ancira.axelone.app/'
  }

  // logoutSession(): void {

  //   const site = this.userInfo?.site ;
  //   localStorage.clear();
  //   window.location.href =
  //     site === 'demo' ? 'https://demo.axelone.app?logout=true' : 'http://axelone.app?logout=true';
  // }

  logoutSession(): void {
    const site = this.userInfo?.site;
    localStorage.clear();
    sessionStorage.clear();
    history.pushState(null, '', location.href);
    window.addEventListener('popstate', () => {
      history.pushState(null, '', location.href);
    });
    const target =
      site === 'demo' ? 'https://demo.axelone.app?logout=true' : 'https://axelone.app?logout=true';
    window.location.replace(target); // THIS prevents going back
  }

  onNetworkClick() {
    if (this.userNetworkInfo?.networkStatus) {
      this.toast.show('You are already connected to the office network.', 'success', 'Connected');
    } else {
      this.toast.show('Please connect to the office network.', 'warning', 'Not Connected');
    }
  }

  // onExcelClick() {
  //   const Data = {
  //     title: this.componentData.title,
  //     state: true,
  //   };
  //   this.apiCall.setExportToExcelAllReports({ obj: Data });
  // }
  // onPDFClick() {
  //   const Data = {
  //     title: this.componentData.title,
  //     statePDF: true,
  //   };
  //   this.apiCall.setExportToPDFAllReports({ obj: Data });
  // }

  showExport = false;
  showExportMail = false;
  toggleExport(event: Event) {
    event.stopPropagation();
    this.mailtooltip = false;
    this.showusersAutofill = false;
    this.showExport = !this.showExport;
  }
  toggleExportMail(event: Event) {
    event.stopPropagation();
    this.showExportMail = !this.showExportMail;
    this.showExport = false;
  }

  closeAllDropdowns() {
    this.showExport = false;
    this.showExportMail = false;
    this.mailtooltip = false;
    this.showusersAutofill = false;
  }

  isExporting = false;
  export(type: string, event: Event): void {
    event.stopPropagation();
    if (this.isExporting) return;
    this.isExporting = true;
    if (type === 'xlsx' || type === 'csv') {
      this.onExcelClick();
      this.showExport = false;
    } else if (type === 'pdf') {
      this.onPDFClick();
      this.showExport = false;
    } else if (type === 'print') {
      this.onPrintClick();
      this.showExport = false;
    }
  }
  onExcelClick() {
    const Data = {
      title: this.componentData.title,
      state: true,
    };
    this.apiCall.setExportToExcelAllReports({ obj: Data });
    setTimeout(() => {
      this.isExporting = false;
    }, 2000);
  }
  onPrintClick() {
    const Data = {
      title: this.componentData.title,
      statePrint: true,
    };
    this.apiCall.setExportToPrintAllReports({ obj: Data });
    setTimeout(() => {
      this.isExporting = false;
    }, 2000);
  }
  onPDFClick() {
    const Data = {
      title: this.componentData.title,
      statePDF: true,
    };
    this.apiCall.setExportToPDFAllReports({ obj: Data });
    setTimeout(() => {
      this.isExporting = false;
    }, 2000);
  }

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const target = event.target as HTMLElement;

    const clickedInsideExport = target.closest('.export-btn, .export-card-btn');
    const clickedInsideMail = target.closest('.export-mail, .export-card-mail, .autofillcls, .search');

    if (!clickedInsideExport) {
      this.showExport = false;
    }

    if (!clickedInsideMail) {
      this.mailtooltip = false;
      this.showusersAutofill = false;
    }
  }
  groupsandstoresdata: any = []
  getGruopsandStores() {
    this.shared.common.completeUserDetails = this.userInfo
    // alert('HI')
    const obj = {
      "userid": JSON.parse(localStorage.getItem('userInfo')!).user_Info.userid,
      // "userid": 44,
    }
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetStoresList', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.shared.common.allstores = res.response
          let data = res.response
          // .filter((val: any) => val.sg_id != 7)
          this.groupsandstoresdata = data.reduce((r: any, { sg_name, sg_id }: any) => {
            if (!r.some((o: any) => o.sg_name == sg_name)) {
              r.push({
                sg_name,
                sg_id,
                Stores: data.filter((v: any) => v.sg_name == sg_name),
              });
            }
            return r;
          }, []);
          console.log(this.groupsandstoresdata);
          // .filter((val:any)=>{ val.sg_name != 'Other Stores' })
          this.shared.common.groupsandstores = this.groupsandstoresdata.filter((val: any) => val.sg_name != 'OtherStores')
          this.shared.common.OtherStoresData = this.groupsandstoresdata.filter((val: any) => val.sg_name == 'OtherStores')
          // .filter((val:any)=>{ val.sg_name == 'Other Stores' })
          const obj = {
            storesData: this.shared.common.groupsandstores,
            otherStoresData: this.shared.common.OtherStoresData
          };
          this.shared.api.setStores({ obj: obj });
        }
      })
  }
  getGruopsandStoresAxelOne() {
    this.shared.common.completeUserDetails = this.userInfo
    // alert('HI')
    const obj = {
      "StoreIds": JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids,
    }
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetStoresListByStoreIds', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.shared.common.allstores = res.response
          let data = res.response
          // .filter((val: any) => val.sg_id != 7)
          this.groupsandstoresdata = data.reduce((r: any, { sg_name, sg_id }: any) => {
            if (!r.some((o: any) => o.sg_name == sg_name)) {
              r.push({
                sg_name,
                sg_id,
                Stores: data.filter((v: any) => v.sg_name == sg_name),
              });
            }
            return r;
          }, []);
          console.log(this.groupsandstoresdata);
          // .filter((val:any)=>{ val.sg_name != 'Other Stores' })
          this.shared.common.groupsandstores = this.groupsandstoresdata.filter((val: any) => val.sg_name != 'OtherStores')
          this.shared.common.OtherStoresData = this.groupsandstoresdata.filter((val: any) => val.sg_name == 'OtherStores')
          // .filter((val:any)=>{ val.sg_name == 'Other Stores' })
          const obj = {
            storesData: this.shared.common.groupsandstores,
            otherStoresData: this.shared.common.OtherStoresData
          };
          this.shared.api.setStores({ obj: obj });
        }
      })
  }
  groupsandstoresdataAll: any = []
  getGruopsandStoresAll() {
    const obj = {
      // "userid": JSON.parse(localStorage.getItem('UserDetails')!).userid,
      "userid": 0,
    }
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetStoresList', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.shared.common.allstores = res.response
          let data = res.response
          // .filter((val: any) => val.sg_id != 7)
          this.groupsandstoresdataAll = data.reduce((r: any, { sg_name, sg_id }: any) => {
            if (!r.some((o: any) => o.sg_name == sg_name)) {
              r.push({
                sg_name,
                sg_id,
                Stores: data.filter((v: any) => v.sg_name == sg_name),
              });
            }
            return r;
          }, []);
          console.log(this.groupsandstoresdata);
          // .filter((val:any)=>{ val.sg_name != 'Other Stores' })
          this.shared.common.groupsandstoresAll = this.groupsandstoresdataAll.filter((val: any) => val.sg_name != 'OtherStores')
          // this.shared.common.OtherStoresData = this.groupsandstoresdataAll.filter((val: any) => val.sg_name == 'OtherStores')
          // .filter((val:any)=>{ val.sg_name == 'Other Stores' })
          const obj = {
            storesData: this.shared.common.groupsandstoresAll,
            otherStoresData: this.shared.common.OtherStoresData
          };
          this.shared.api.setStoresAll({ obj: obj });
          console.log(this.shared.common.groupsandstoresAll, 'All Stores List');

        }
      })
  }

  getMobileStores() {
    const obj = {
      "userid": JSON.parse(localStorage.getItem('userInfo')!).user_Info.userid,
    };
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetStoresListMobileServiceGL', obj)
      .subscribe((res: any) => {
        //console.log(res.response);
        if (res.status == 200) {
          let data = res.response
          let withoutreduce = data.reduce((r: any, { sg_name, sg_id }: any) => {
            if (!r.some((o: any) => o.sg_name == sg_name)) {
              r.push({
                sg_name,
                sg_id,
                Stores: data.filter((v: any) => v.sg_name == sg_name),
              });
            }
            return r;
          }, []);
          this.shared.common.MobileServiceGL = res.response
          const obj = {
            storesData: this.shared.common.MobileServiceGL
          };
          this.shared.api.setMobileService({ obj: obj });
          console.log(this.shared.common.MobileServiceGL, obj);

        }
      });
  }
  getAccountingReconStores() {
    const obj = {
      "userid": JSON.parse(localStorage.getItem('userInfo')!).user_Info.userid,
    };
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetStoresListRecon', obj)
      .subscribe((res: any) => {
        //console.log(res.response);
        if (res.status == 200) {
          let data = res.response
          let withoutreduce = data.reduce((r: any, { sg_name, sg_id }: any) => {
            if (!r.some((o: any) => o.sg_name == sg_name)) {
              r.push({
                sg_name,
                sg_id,
                Stores: data.filter((v: any) => v.sg_name == sg_name),
              });
            }
            return r;
          }, []);
          this.shared.common.AccountingReconStores = res.response
          const obj = {
            storesData: this.shared.common.AccountingReconStores
          };
          this.shared.api.setAccuntingAllStores({ obj: obj });
        }
      });
  }
  getReconStores() {
    const obj = {
      "userid": JSON.parse(localStorage.getItem('userInfo')!).user_Info.userid,
      // "userid": 44,
    }
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetStoresListCustom', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          let data = res.response
          let withoutreduce = data.reduce((r: any, { sg_name, sg_id }: any) => {
            if (!r.some((o: any) => o.sg_name == sg_name)) {
              r.push({
                sg_name,
                sg_id,
                Stores: data.filter((v: any) => v.sg_name == sg_name),
              });
            }
            return r;
          }, []);
          this.shared.common.ReconStores = withoutreduce.filter((val: any) => val.sg_name != 'OtherStores')
          const obj = {
            storesData: this.shared.common.ReconStores
          };
          this.shared.api.setAllStores({ obj: obj });
        }
      })
  }

  getStoreByGroupUserIdRecon() {
    const obj = {
      id: '',
      "userid": JSON.parse(localStorage.getItem('userInfo')!).user_Info.userid,
    };
    this.shared.api
      .postmethodtoken(this.shared.common.routeEndpoint + 'GetStoresbyGroupuseridRecon', obj)
      .subscribe((res: any) => {
        //console.log(res.response);
        if (res.status == 200) {
          let data = res.response
          let withoutreduce = data.reduce((r: any, { masterStore, masterid }: any) => {
            if (!r.some((o: any) => o.masterid == masterid)) {
              r.push({
                masterStore,
                masterid,
                Stores: data.filter((v: any) => v.masterid == masterid),
              });
            }
            return r;
          }, []);
          this.shared.common.ReconServiceLaborStores = withoutreduce
          const obj = {
            storesData: this.shared.common.ReconServiceLaborStores
          };
          this.shared.api.setReconServiceLaborStores({ obj: obj });
        }
      });
  }
  @Input() title!: string;
  @Input() tableId!: string;


  // excel() {
  //   this.exportService.exportExcel(this.componentData?.title, this.componentData?.title);
  // }

  // pdf() {
  //   this.exportService.exportPDF(this.componentData?.title, this.componentData?.title);
  // }

  // print() {
  //   this.exportService.print(this.tableId, this.componentData?.title);
  // }

  // email() {
  //   this.exportService.emailPDF(this.tableId, this.componentData?.title);
  // }


  downloadPDF() {
    // alert('Hi')
    const Data = {
      title: this.componentData.title,
      state: true,
    };
    this.apiCall.setExportToPDFAllReports({ obj: Data });
  }

  users: any[] = [];          // all users (available list)
  filteredUsers: any[] = [];  // search result
  UsersList: any[] = [];      // selected emails
  UsersData: any[] = [];      // API data

  mailuser: string = '';
  notes: string = '';

  showusersAutofill = false;
  mailtooltip = false;
  selectedItem: any = null;

  getUsersData() {
    const obj = {
      userid: '0',
      flag: '',
      exp: ''
    };

    this.shared.api.postmethodHRDB('users/GetAOUsers', obj).subscribe((res) => {
      this.UsersData = res.response.filter((e: any) => e.Status === 'Y');

      this.users = [...this.UsersData];        // available users
      this.filteredUsers = [...this.UsersData]; // for search
    });
  }

  mailBlock() {
    this.users = [...this.UsersData];
    this.filteredUsers = [...this.UsersData];
    this.UsersList = [];
    this.mailuser = '';
    this.notes = '';
    this.showExport = false;
    this.showusersAutofill = false;
    this.mailtooltip = !this.mailtooltip;
  }
  onSearchChange() {
    const val = this.mailuser?.toLowerCase() || '';

    this.filteredUsers = this.users.filter((user: any) =>
      user.FullName.toLowerCase().includes(val) ||
      user.MailID.toLowerCase().includes(val)
    );

    this.showusersAutofill = true;
  }
  dropdownopen() {
    this.filteredUsers = [...this.users];
    this.showusersAutofill = true;
  }
  selectItem(item: any) {
    if (!item) return;

    // prevent duplicate
    const exists = this.UsersList.some(x => x.MailID === item.MailID);
    if (exists) return;

    this.UsersList.push(item);

    // remove from available list
    this.users = this.users.filter(x => x.Id !== item.Id);

    this.mailuser = '';
    this.showusersAutofill = false;
  }

  removeBrandTag(item: any, index: number) {
    this.UsersList.splice(index, 1);

    // add back to dropdown list
    if (item.Id) {
      this.users.push(item);
    }
  }

  addEmails() {
    if (!this.validateEmail(this.mailuser)) {
      this.toast.show('Please Enter Valid EmailId', 'warning', 'Warning');
      return;
    }

    const exists = this.UsersList.some(x => x.MailID === this.mailuser);
    if (exists) return;

    this.UsersList.push({
      Id: '',
      FullName: '',
      MailID: this.mailuser
    });

    this.mailuser = '';
  }
  validateEmail(email: string): boolean {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  }
  sendMail() {
    if (!this.UsersList.length) {
      this.toast.show('Please Enter Atleast One EmailId', 'warning', 'Warning');
      return;
    }
    const Data = {
      title: this.componentData.title,
      stateEmailPdf: true,
      Email: this.UsersList.map(x => x.MailID).join(','),
      notes: this.notes,
      from: JSON.parse(localStorage.getItem('userInfo')!).fullName
    };
    this.apiCall.setExportToEmailPDFAllReports({ obj: Data });
    setTimeout(() => {
      this.isExporting = false;
    }, 2000);

    this.mailBlock();
  }

  excludedPrintTitles: string[] = [
    'Tables Last Updated',
    'Service RO Details',
    'Buyer Ranking',
    'Washout Report',

    'Search Parts',
    'Parts Aging',
    'Parts Gross Profit Score Card',
    'Recon Transactional',
    'Recon GL Report',
    'Recon Open ROs',
    'Total Transactional Gross'
  ];
  excludedPdfTitles: string[] = [
    'Tables Last Updated',
    'Service RO Details',
    'Chargebacks',
    'Buyer Ranking',
    'Washout Report',
    'Total Transactional Gross'


  ];
  excludedExcelTitles: string[] = [
    'Buyer Ranking',
    'Washout Report',
    'Total Transactional Gross'

  ];
  excludedEmailToPdfTitles: string[] = [
    'Tables Last Updated',
    'Service RO Details',
    'Buyer Ranking',
    'Washout Report',

    'Search Parts',
    'Parts Aging',
    'Parts Gross Profit Score Card',
    'Recon Transactional',
    'Recon GL Report',
    'Recon Open ROs',
    'Total Transactional Gross'


  ];

  isActionVisible(actionType: 'print' | 'pdf' | 'excel' | 'emailToPdf'): boolean {
    let excludedTitles: string[];
    switch (actionType) {
      case 'print':
        excludedTitles = this.excludedPrintTitles;
        break;
      case 'pdf':
        excludedTitles = this.excludedPdfTitles;
        break;
      case 'excel':
        excludedTitles = this.excludedExcelTitles;
        break;
      case 'emailToPdf':
        excludedTitles = this.excludedEmailToPdfTitles;
        break;
    }
    return !excludedTitles.includes(this.componentData.title);
  }

  exportFile() {
    if (this.excludedPrintTitles.indexOf(this.componentData.title) > 0 && this.excludedPrintTitles.indexOf(this.componentData.title) > 0 &&
      this.excludedPrintTitles.indexOf(this.componentData.title) > 0 && this.excludedPrintTitles.indexOf(this.componentData.title) > 0) {
      this.isDownload = false
    } else {
      this.isDownload = true
    }
  }



}


