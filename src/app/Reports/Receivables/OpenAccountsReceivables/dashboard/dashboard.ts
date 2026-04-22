import { Component, ElementRef, HostListener, ViewChild } from '@angular/core';
import { NgbActiveModal, NgbDateParserFormatter, NgbModal, NgbModule } from '@ng-bootstrap/ng-bootstrap';
import { Workbook } from 'exceljs';
import { DatePipe } from '@angular/common';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { Subscription } from 'rxjs';
// import { DealRecapComponent } from 'src/app/Global/cdpdataview/deal/deal-recap/deal-recap.component';

// Shared / app-level utilities & modules (adapted to your project's shared service/module)
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { FilterPipe } from '../../../../Core/Providers/filterpipe/filter.pipe';
import { ChangeDetectorRef } from '@angular/core';
import { TimeConversionPipe } from '../../../../Core/Providers/pipes/timeconversion.pipe';
import { NgxSpinnerService } from 'ngx-spinner';
import { Router } from '@angular/router';
import { common } from '../../../../common';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import * as FileSaver from 'file-saver';
import autoTable from 'jspdf-autotable';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, Stores, NgbModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
  standalone: true,
})
export class Dashboard {
  /* ------------------------------- VARIABLES ------------------------------- */

  NoData: boolean = false;
  FloorPlanData: any = [];
  FloorPlanTotalData: any = [];
  LiabilitiesData: any = [
    {
      Liabilities: 'All',
      path: 'ALL',
      title: 'ALL',
      url: 'ALL',
    },
    {
      Liabilities: 'Inventory',
      path: 'Inventory',
      title: 'Inventory',
      url: 'Inventory',
    },
    {
      Liabilities: 'Sold',
      path: 'Sold',
      title: 'Sold',
      url: 'Sold',
    }
  ];
  allordebit: any = 'dr';
  selectedreceviabe: any = [];
  StoreVal: any = '71,53,8,7,4,35,1,32,40,50,25,18,31,3,70,72,2,17,41,55,42,51,12,73,54,9,15,5,14,30,11';
  spinnerLoader: boolean = true;
  Role: any = [];
  userid: any;

  index: string = '';
  groups: any = 1;
  financeManagerId: any = '0';
  AgeFrom: any = 0;
  AgeTo: any = 0;
  QISearchName: any = '';

  callLoadingState = 'FL';
  enablevehicle: boolean = false;
  commentsVisibility: boolean = true;

  hideVisibility: boolean = false;
  hideRecords: any = [];
  FinalArray: any = [];

  header: any = [
    {
      type: 'Bar',

      storeIds: this.StoreVal,
      groups: this.groups,
      financemanagers: this.financeManagerId,
      ageFrom: this.AgeFrom,
      ageTo: this.AgeTo,
    },
  ];

  reportgetting!: Subscription;

  popup: any = [{ type: 'Popup' }];
  // actionType: any = 'N';

  notesStageValue: any = '';
  notesStageText: any = '';
  notesstage: any = [];
  selecteddata: any = [];
  check: boolean = false;
  activePopover: number = -1;

  selectedFiManagersvalues: any = [];
  selectedFiManagersname: any = [];
  financeManager: any = [];
  storename: any = '';
  groupsArray: any = [];
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 8;
  storeIds: any = '';
  stores: any = [];

  storesFilterData: any = {
    groupsArray: this.groupsArray,
    groupId: this.groupId,
    storesArray: this.stores,
    storeids: '1',
    type: 'M',
    others: 'N',
    groupName: this.groupName,
    storename: this.storename,
    storecount: null,
    storedisplayname: this.storedisplayname,
  };
  // solutionurl: any = environment.apiUrl;
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest(
      '.dropdown-toggle, .dropdown-menu , .timeframe, .reportstores-card',
    );
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  /* ------------------------------- CONSTRUCTOR ------------------------------- */

  constructor(
    private ngbmodal: NgbModal,
    private comm: common,
    public shared: Sharedservice,
    private router: Router,
    private ngbmodalActive: NgbActiveModal,
    private spinner: NgxSpinnerService,
    private toast: ToastService,
    private datepipe: DatePipe,
  ) {
    if (localStorage.getItem('flag') == 'V') {
      this.storeIds = [];
      console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.groupid
      JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store.indexOf(',') > 0 ?
        this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store.split(',') :
        this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store)
      localStorage.setItem('flag', 'M')
    } else {
      if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
        this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
      }
    }
    this.commentsVisibility = true;

    /* --------- USER DETAILS --------- */
    if (localStorage.getItem('userInfo') != null) {
      // Keep same logic but don't break when redirectionFrom missing
      try {
        const ud: any = JSON.parse(localStorage.getItem('userInfo')!);
        this.Role = ud.user_Info.title;
        this.userid = ud.user_Info.userid;
        console.log(this.Role, this.userid);
      } catch {
        // ignore
      }

    }

    /* --------- STORE SETUP --------- */
    // let storeids = userData.Store_Ids;

    // if (storeids.toString().indexOf(',') > 0) {
    //   this.StoreVal = "";
    //   this.actionType = 'N';
    // } else {
    //   this.StoreVal = storeids;
    //   this.actionType = 'Y';
    // }

    /* --------- DEFAULT SELECTED TAB --------- */
    this.selectedreceviabe = this.LiabilitiesData[0];
    // this.selectedreceviabe ='NewFlooring A/P'

    /* --------- ROUTE CHECK FOR SELECTED TAB --------- */
    const path = this.router.url.split('?')[0].replace('/', '');

    const selectedpath = this.LiabilitiesData.find((e: any) => e.url === path);

    this.selectedreceviabe = selectedpath ? selectedpath : this.LiabilitiesData[0];

    if (path === 'Open Account Receivables') {
      this.selectedreceviabe = this.LiabilitiesData[0];
    } else {
      let selectedpath = this.LiabilitiesData.find((e: any) => e.url === path);
      this.selectedreceviabe = selectedpath ? selectedpath : this.LiabilitiesData[0];
    }

    this.shared.setTitle(this.comm.titleName + '-Open Account Receivables');

    /* --------- HEADER FOR REPORTING --------- */
    const data = {
      title: 'Open Account Receivables',
      stores: this.StoreVal,
      groups: this.groups,
      financemanagers: '',
      count: 0,
      AgeFrom: this.AgeFrom,
      AgeTo: this.AgeTo,
      search: this.QISearchName,
    };

    this.shared.api.SetHeaderData({ obj: data });

    this.header = [
      {
        type: 'Bar',
        storeIds: this.StoreVal,
        groups: this.groups,
        financemanagers: this.financeManagerId,
        ageFrom: this.AgeFrom,
        ageTo: this.AgeTo,
      },
    ];
    this.Getfloorplansdata(this.ARSearchName);
  }

  /* ------------------------------- LIFECYCLE ------------------------------- */

  ngOnInit(): void { }
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Open Account Receivables') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.shared.api.GetReportOpening().subscribe((res) => {
      if (res.obj.Module === 'Open Account Receivables') {
        document.getElementById('report')?.click();
      }
    });

    /* -------- REFRESH DATA FROM HEADER -------- */
    this.reportgetting = this.shared.api.GetReports().subscribe((data) => {
      if (data.obj.Reference === 'Open Account Receivables') {
        this.FloorPlanData = [];
        // this.actionType = 'Y';
        this.NoData = false;

        /* Update filters */
        if (!data.obj.header) {
          this.StoreVal = data.obj.storeValues;
          this.financeManagerId = data.obj.FIvalues;
          this.groups = data.obj.groups;
          this.AgeFrom = data.obj.AgeFrom;
          this.AgeTo = data.obj.AgeTo;
        }

        if (this.StoreVal != '') {
          this.Getfloorplansdata(this.ARSearchName);
        } else {
          this.NoData = true;
        }

        /* Update header for next reload */
        const headerdata = {
          title: 'Open Account Receivables',
          stores: this.StoreVal,
          groups: this.groups,
          financemanagers: this.financeManagerId,
          AgeFrom: this.AgeFrom,
          AgeTo: this.AgeTo,
        };

        this.shared.api.SetHeaderData({ obj: headerdata });

        this.header = [
          {
            type: 'Bar',
            storeIds: this.StoreVal,
            groups: this.groups,
            financemanagers: this.financeManagerId,
            ageFrom: this.AgeFrom,
            ageTo: this.AgeTo,
          },
        ];
      }
    });

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Open Account Receivables') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Open Account Receivables') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Open Account Receivables') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Open Account Receivables') return;
      if (obj.stateEmailPdf) {
        this.sendEmailData(obj.Email, obj.notes, obj.from);
      }
    });
  }
  ngOnDestroy(): void {
    this.excel?.unsubscribe();
    this.print?.unsubscribe();
    this.pdf?.unsubscribe();
    this.email?.unsubscribe();
  }
  //   formatBalance(val: number | null, decimals = 2): string {
  //   if (val === null || val === undefined) return '-';
  //   return val < 0
  //     ? `-$${Math.abs(val).toLocaleString(undefined, { minimumFractionDigits: decimals })}`
  //     : `$${val.toLocaleString(undefined, { minimumFractionDigits: decimals })}`;
  // }

  formatBalance(val: number | null): string {
    if (val === null || val === undefined) return '-';
    return val < 0 ? `-$${Math.abs(val).toFixed(2)}` : `$${val.toFixed(2)}`;
  }

  formatBalancetozero(val: number | null): string {
    if (val === null || val === undefined) return '-';
    if (val === 0) return '0';
    return val < 0 ? `-$${Math.abs(val).toFixed(0)}` : `$${val.toFixed(0)}`;
  }

  keyPressNumbers(event: any) {
    var charCode = event.which ? event.which : event.keyCode;
    // Only Numbers 0-9
    if (charCode < 48 || charCode > 57) {
      event.preventDefault();
      return false;
    } else {
      return true;
    }
  }

  /* ------------------------------- SEARCH HANDLER ------------------------------- */

  receiveMessage($event: any) {
    this.QISearchName = $event;
  }

  /* ------------------------------- SORT ------------------------------- */

  isDesc: boolean = false;
  column: string = 'CategoryName';

  sort(property: string, state?: any) {
    if (state === undefined) {
      this.isDesc = !this.isDesc;
    }

    this.callLoadingState = 'FL';
    this.column = property;

    const direction = this.isDesc ? 1 : -1;

    this.FloorPlanData.sort((a: any, b: any) => {
      let valA = a[property];
      let valB = b[property];

      // Normalize null / dash
      if (valA === null || valA === undefined || valA === '-') valA = '';
      if (valB === null || valB === undefined || valB === '-') valB = '';

      // 🔑 Detect numeric values (Control, Deal #, Stock # when numeric)
      const numA = Number(valA);
      const numB = Number(valB);

      const isNumA = !isNaN(numA);
      const isNumB = !isNaN(numB);

      if (isNumA && isNumB) {
        return (numA - numB) * direction;
      }

      return String(valA).localeCompare(String(valB), undefined, {
        numeric: true,
        sensitivity: 'base'
      }) * direction;
    });
  }

  /* ------------------------------- MAIN API CALL ------------------------------- */
  previousReportPath: string | null = null;
  ARSearchName: string = '';
  pageNumber: number = 1;
  pageSize: number = 500;
  loading: boolean = false;
  Scrollpercent: number = 0;
  scrollCurrentposition: number = 0;

  @ViewChild('scrollcent') scrollcent!: ElementRef;

  // ---------------------- Fetch Data ----------------------
  Getfloorplansdata(searchName: string) {
    this.spinner.show();
    this.goToFirstPage();
    const obj = {
      CUSTNAME: searchName || "",
      ASID: this.storeIds?.toString() || "",
      PageNumber: this.pageNumber,
      PageSize: this.pageSize.toString()
    };

    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetAropenAccountReceivables', obj)
      .subscribe({
        next: (res: any) => {
          this.spinner.hide();
          if (res?.status === 200 && Array.isArray(res.response)) {
            if (this.pageNumber === 1) {
              this.FloorPlanData = res.response;
            } else {
              this.FloorPlanData = [...this.FloorPlanData, ...res.response];
            }
            this.filteredFloorplanData = [...this.FloorPlanData];
            this.NoData = this.FloorPlanData.length === 0;
          } else {
            this.NoData = true;
            console.warn('No data received from API', res);
          }
        },
        error: (err) => {
          this.spinner.hide();
          this.NoData = true;
          console.error('API Error:', err);
          this.toast.show('502 Bad Gateway Error', 'danger', 'Error');
        }
      });
  }

  // ---------------------- Search ----------------------
  searchFilter() {
    this.pageNumber = 1;
    this.NoData = false;
    this.Getfloorplansdata(this.ARSearchName);
  }

  onKeypressEvent(event: KeyboardEvent) {
    if (event.key === 'Enter') {
      this.pageNumber = 1;
      this.NoData = false;
      this.Getfloorplansdata(this.ARSearchName);
    }
  }


  updateVerticalScroll(event: any): void {
    this.scrollCurrentposition = event.target.scrollTop;
    const scrollEl = this.scrollcent.nativeElement as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop / (scrollEl.scrollHeight - scrollEl.clientHeight)) * 100
    );

    const nearBottom = event.target.scrollTop + event.target.clientHeight >= event.target.scrollHeight - 10;

    if (nearBottom && this.filteredFloorplanData.length >= this.pageSize) {
      this.pageNumber++;
      this.Getfloorplansdata(this.ARSearchName);
    }
  }
  ////////////////////////////////////////////Pagination Code////////////////////////////////////

  searchText: string = '';
  currentPage: number = 1;
  itemsPerPage: number = 100;
  maxPageButtonsToShow: number = 5;
  clickedPage: number | null = null;

  filteredFloorplanData: any[] = [];
  setFloorplanData() {
    this.filteredFloorplanData = this.FloorPlanData || [];
    this.currentPage = 1; // reset page
  }
  get filteredData() {

    if (!this.searchText) {
      return this.filteredFloorplanData;
    }

    const searchTerms = this.searchText
      .split(',')
      .map(term => term.trim().toLowerCase())
      .filter(term => term.length > 0);

    const filtered = this.filteredFloorplanData.filter((x: any) =>
      searchTerms.some(term =>

        x.age?.toString().toLowerCase().includes(term) ||
        x.dealdate?.toString().toLowerCase().includes(term) ||

        x.stock?.toString().toLowerCase().includes(term) ||

        x.customername?.toLowerCase().includes(term) ||
        x.customer?.toString().toLowerCase().includes(term) ||

        x.deal?.toString().toLowerCase().includes(term) ||
        x.dealstatus?.toLowerCase().includes(term) ||

        x.year?.toString().toLowerCase().includes(term) ||
        x.make?.toLowerCase().includes(term) ||
        x.model?.toLowerCase().includes(term) ||

        x.vin?.toLowerCase().includes(term) ||

        x.invstatus?.toLowerCase().includes(term) ||
        x.titlerecd?.toString().toLowerCase().includes(term)

      )
    );
    return filtered.sort((a: any, b: any) => {
      const aIndex = searchTerms.findIndex(term =>
        Object.values(a).join(' ').toLowerCase().includes(term)
      );
      const bIndex = searchTerms.findIndex(term =>
        Object.values(b).join(' ').toLowerCase().includes(term)
      );
      return aIndex - bIndex;
    });
  }
  sortColumn: string = '';
  sortDirection: 'asc' | 'desc' = 'asc';
  sortTable(column: string) {
    if (this.sortColumn === column) {
      this.sortDirection = this.sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
      this.sortColumn = column;
      this.sortDirection = 'asc';
    }
    this.filteredFloorplanData.sort((a: any, b: any) => {
      let valA = a[column];
      let valB = b[column];
      valA = valA ?? '';
      valB = valB ?? '';
      if (!isNaN(valA) && !isNaN(valB)) {
        return this.sortDirection === 'asc'
          ? valA - valB
          : valB - valA;
      }
      return this.sortDirection === 'asc'
        ? valA.toString().localeCompare(valB.toString())
        : valB.toString().localeCompare(valA.toString());

    });

  }
  getSortIcon(column: string) {
    if (this.sortColumn !== column) return 'fa-sort';
    return this.sortDirection === 'asc'
      ? 'fa-sort-up'
      : 'fa-sort-down';
  }
  get paginatedItems() {
    const startIndex = (this.currentPage - 1) * this.itemsPerPage;
    const endIndex = startIndex + this.itemsPerPage;
    return this.filteredData.slice(startIndex, endIndex);
  }
  onPageSizeChange() {
    this.currentPage = 1;
  }
  getMaxPageNumber(): number {
    return Math.ceil(this.filteredData.length / this.itemsPerPage);
  }
  nextPage() {
    if (this.currentPage < this.getMaxPageNumber()) this.currentPage++;
  }
  prevPage() {
    if (this.currentPage > 1) this.currentPage--;
  }
  goToFirstPage() {
    this.currentPage = 1;
  }

  goToLastPage() {
    this.currentPage = this.getMaxPageNumber();
  }
  getStartRecordIndex(): number {
    return (this.currentPage - 1) * this.itemsPerPage;
  }
  getEndRecordIndex(): number {
    const endIndex = this.getStartRecordIndex() + this.itemsPerPage;
    return endIndex > this.filteredData.length
      ? this.filteredData.length
      : endIndex;
  }

  get BalanceTotal(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.amount) || 0);
    }, 0);
  }

  get BalFloorplanTotal(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.BalFloorplan) || 0);
    }, 0);
  }
  ////////////////////////////////////Pagination Code End////////////////////////////////////////////




  getStoresandGroupsValues() {
    this.storesFilterData.groupsArray = this.groupsArray;
    this.storesFilterData.groupId = this.groupId;
    this.storesFilterData.storesArray = this.stores;
    this.storesFilterData.storeids = this.storeIds;
    this.storesFilterData.groupName = this.groupName;
    this.storesFilterData.storename = this.storename;
    this.storesFilterData.storecount = this.storecount;
    this.storesFilterData.storedisplayname = this.storedisplayname;

    this.storesFilterData = {
      groupsArray: this.groupsArray,
      groupId: this.groupId,
      storesArray: this.stores,
      storeids: this.storeIds,
      groupName: this.groupName,
      storename: this.storename,
      storecount: this.storecount,
      storedisplayname: this.storedisplayname,
      'type': 'M', 'others': 'N'
    };

    // this.setHeaderData();
    // this.GetData();

  }
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
  }
  AllorDebit(e: any) {
    this.allordebit = []
    // if (e == 'All') {
    //   if (this.dealStatus.length == 3) {
    //     this.dealStatus = []
    //     this.toast.warning('Please Select Atleast One dealStatus', '');

    //   } else {
    //     this.dealStatus = []
    //     this.dealStatus = ['Booked', 'Finalized','Delivered'];
    //   }
    // } 
    // else {
    const index = this.allordebit.findIndex((i: any) => i == e);
    if (index >= 0) {
      this.allordebit.splice(index, 1);
    } else {
      this.allordebit.push(e);
    }
    // }
  }
  /* ------------------------------- NOTES LOGIC ------------------------------- */

  viewmoreAction(fp: any) {
    if (fp.Notesstate === '+') {
      fp.Notesstate = '-';
      fp.duplicateNotes = fp.Notes;
    } else {
      fp.Notesstate = '+';
      fp.duplicateNotes = fp.Notes.slice(0, 3);
    }
  }

  getDropDown(companyid: any) {
    const obj = {
      AssociatedReport: this.selectedreceviabe.path,
      CompanyID: companyid,
    };

    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetScheduleNoteStages', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.notesstage = res.response;
          this.notesStageValue = '';
        }
      });
  }

  addNotes(item: any) {
    this.selecteddata = item;
    this.getDropDown(item.companyid);
    this.notesStageText = '';
    this.notesStageValue = '';
  }

  formatDate(date?: number | string | Date): string {
    if (!date) return '-';

    const parsedDate = new Date(date);

    if (isNaN(parsedDate.getTime())) return '-';

    return new Intl.DateTimeFormat('en-US', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric',
    }).format(parsedDate);
  }
  save() {
    if (this.notesStageText.trim() === '') {
      this.toast.show('Please enter notes', 'warning', 'Warning');
      return;
    }

    const obj = {
      AS_ID: this.selecteddata.storeid,
      Account: this.selecteddata.Account,
      Control: this.selecteddata.Control,
      Notes: this.notesStageText,
      StageId: this.notesStageValue,
      UserID: this.userid,
    };

    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'AddScheduleNotesAction', obj)
      .subscribe((res: any) => {
        if (res.status == 200) {
          this.toast.show('Notes Added Successfully', 'success', 'Success');
          this.callLoadingState = 'ANS';
          (document.getElementById('close') as HTMLInputElement)?.click();
          this.oncloseone();
          this.commentsVisibility = true;

          /* Instant Frontend Update */
          const userName =JSON.parse(localStorage.getItem('userInfo')!).fullName
          const curDate = new Date();
          let nts = '';

          if (this.notesStageValue) {
            const filtered = this.notesstage.find(
              (item: any) => item.NS_ID == this.notesStageValue,
            );
            nts = `[${filtered.NS_Text}] ${this.notesStageText} - ${userName} - ${this.formatDate(curDate)}`;
            this.selecteddata.Stage = filtered.NS_Text;
          } else {
            nts = `${this.notesStageText} - ${userName} - ${this.formatDate(curDate)}`;
          }

          const newNote = {
            STAGE: '',
            NOTES: nts,
            NOTESDATE: this.formatDate(curDate),
            UserName: userName,
          };

          if (!this.selecteddata.duplicateNotes) {
            this.selecteddata.duplicateNotes = [];
          }

          this.selecteddata.duplicateNotes.unshift(newNote);
          this.selecteddata.NotesStatus = 'Y';
        } else {
          this.toast.show('Something went wrong. Please try again.', 'danger', 'Error');
        }
      });
  }

  /* ------------------------------- HIDE RECORD LOGIC ------------------------------- */

  collectHidevalues(e: any, val: any, confirmtemplate: any, ref: any, refval: any) {
    if (ref === 'multi') {
      if (this.hideRecords.length === 0) {
        this.toast.show('Please Select Atleast One Record to Hide', 'warning', 'Warning');
        (document.getElementById('symbol') as HTMLInputElement).checked = false;
        return;
      }

      if (e.target.checked) {
        this.hideVisibility = true;
        this.ngbmodalActive = this.ngbmodal.open(confirmtemplate, {
          size: 'sm',
          backdrop: 'static',
        });
      }
    } else {
      if (e.target.checked) {
        this.hideVisibility = true;
        this.hideRecords.push(val);
      } else {
        const index = this.hideRecords.findIndex((list: any) => list.StockNo == refval);
        this.hideRecords.splice(index, 1);
      }
    }
  }

  oncloseone() {
    (document.getElementById('symbol') as HTMLInputElement).checked = false;
    this.ngbmodalActive.close();
  }

  hideAdd() {
    if (this.hideRecords.length === 0) {
      this.toast.show('Please Select Atleast One Record to Hide', 'warning', 'Warning');
      return;
    }

    this.FinalArray = this.hideRecords.map((item: any) => ({
      Receivable_Type: this.selectedreceviabe.path,
      Account: item.Account,
      CompanyID: item.companyid,
      AS_ID: item.storeid,
      Control: item.Control,
      Stock: item.StockNo,
      Control_Status: 'Y',
      Deal: item.DealNo,
      UserID: this.userid,
    }));

    const obj = { receivableexcludecontrol: this.FinalArray };

    this.shared.api.postmethod('ReceivableExcludeControls', obj).subscribe((res) => {
      if (res.status == 200) {
        this.toast.show('This Control Hidden Successfully', 'success', 'Success');
        (document.getElementById('closeone') as HTMLElement).click();
        this.oncloseone();
        this.Getfloorplansdata(this.ARSearchName);
        this.hideRecords = [];
        this.hideVisibility = false;
      } else {
        this.toast.show('Failed to hide control.', 'danger', 'Error');
      }
    });
  }

  /* ------------------------------- UTIL ------------------------------- */

  public inTheGreen(value: number): boolean {
    return value >= 0;
  }
  onclose() {
    const element = <HTMLInputElement>document.getElementById('symbol');
    if (element) element.checked = false;
    this.ngbmodalActive.close();
  }

  onclosealert() {
    const element = <HTMLInputElement>document.getElementById('symbol');
    if (element) element.checked = false;
    this.ngbmodalActive.close();
    (document.getElementById('close') as HTMLInputElement)?.click();
  }
  viewDeal(dealData: any) {
    // const modalRef = this.ngbmodal.open(DealRecapComponent, {
    //   size: 'md',
    //   windowClass: 'connectedmodal'
    // });
    // modalRef.componentInstance.data = {
    //   dealno: dealData.DealNo,
    //   storeid: dealData.storeid,
    //   stock: dealData.StockNo,
    //   vin: dealData.vin,
    //   custno: dealData.CustomerNumber
    // };
  }

  openComments() {
    this.commentsVisibility = !this.commentsVisibility;
  }

  togglePopover(popoverIndex: number) {
    if (this.activePopover === popoverIndex) {
      // If the same popover is clicked, close it
      this.activePopover = -1;
    } else {
      // Open the selected popover and close others
      this.activePopover = popoverIndex;
    }
  }
  viewreport() {
    this.activePopover = -1;
    let missing = false;
    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show('Please Select Atleast One Store.', 'warning', 'Warning');
      missing = true;
    }
    if (missing) {
      return;
    }
    this.Getfloorplansdata(this.ARSearchName);
  }

  getEmployees(val?: any, ids?: any, count?: any, bar?: any) {
    const obj = {
      AS_ID: this.StoreVal,
      type: 'F',
    };
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetEmployeesDev', obj).subscribe(
      (res: any) => {
        if (res && res.status == 200) {
          // if (val == 'F') {
          this.financeManager = res.response.filter((e: any) => e.FiName != 'Unknown');
          this.selectedFiManagersvalues = this.financeManager.map(function (a: any) {
            return a.FiId;
          });

          // if (bar == 'Bar') {
          //   if (this.employeeschanges != '') {
          //     let fiids = (this.employeeschanges || '').toString().split(',');
          //     this.selectedFiManagersvalues = fiids;
          //   }
          //   if (this.employeeschanges == '0' || this.employeeschanges == 0) {
          //     this.selectedFiManagersvalues = this.financeManager.map(function (a: any) { return a.FiId; });
          //   }
          //   if (this.employeeschanges == '') {
          //     this.selectedFiManagersvalues = [];
          //   }
          // }
          // }
        } else {
          this.toast.show('Invalid Details', 'danger', 'Error');
        }
      },
      (error: any) => {
        /* ignore console errors */
      },
    );
  }
  employees(block: any, e: any, ename?: any) {
    // ========== SINGLE FM TOGGLE ========== //
    if (block === 'FM') {
      const index = this.selectedFiManagersvalues.findIndex((i: any) => i == e);

      if (index >= 0) {
        // already selected -> remove
        this.selectedFiManagersvalues.splice(index, 1);
      } else {
        // not selected -> add
        this.selectedFiManagersvalues.push(e);
      }

      // Optional: set last clicked name (if you use it in UI)
      const index1 = this.selectedFiManagersvalues.findIndex((i: any) => i == e);
      if (index1 >= 0) {
        this.selectedFiManagersname = ename;
      }

      return;
    }

    // ========== SELECT ALL / CLEAR ALL ========== //
    if (block === 'AllFM') {
      // e == 0 → Select All
      // e == 1 → Clear All

      if (e === 0) {
        // SELECT ALL
        this.selectedFiManagersvalues = this.financeManager.map((fm: any) => fm.FiId);
      } else if (e === 1) {
        // CLEAR ALL
        this.selectedFiManagersvalues = [];
      }

      return;
    }
  }



  formatDollar(value: any): string {
    if (value === 0 || value === null || value === undefined) return '-';
    return '$' + Number(value).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  formatDateMMDDYYYY(value: any): string {
    if (!value) return '-';
    const date = new Date(value);
    if (isNaN(date.getTime())) return '-';
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const year = date.getFullYear();
    return `${month}.${day}.${year}`;
  }


  private createPDF(): jsPDF {
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a2' });

    doc.setFontSize(14);
    doc.text('Open Account Receivables', 14, 12);

    const head = [
      'AR Company',
      'Amount',
      'Customer Name',
      'Invoice Number',
      'Journal Type',
      'Reference Date',
      'Source Journal',
      'Sale Company',
      'Target JRN Company'
    ];

    const body: any[] = [];
    this.filteredFloorplanData.forEach(row => {
      body.push([
        row.arcompany || '-',
        row.amount || 0,
        row.customerName || '-',
        row.invoicenumber || '-',
        row.journaltype || '-',
        row.referencedate ? this.formatDateMMDDYYYY(row.referencedate) : '-',
        row.sourcejournal || '-',
        row.salecompany || '-',
        row.targetJRNcompany || '-'
      ]);
    });

    // Totals row
    body.push(['Total', this.BalanceTotal || 0, '', '', '', '', '', '', '']);

    autoTable(doc, {
      tableWidth: 'auto',
      startY: 16,
      head: [head],
      body: body,
      theme: 'grid',
      styles: {
        fontSize: 10,
        cellPadding: 3,
        valign: 'middle'
      },
      headStyles: {
        fillColor: [5, 84, 239],
        textColor: 255,
        fontStyle: 'bold',
        halign: 'center'
      },
      columnStyles: {
        0: { halign: 'left' }, // AR Company
        1: { halign: 'right' }, // Amount
        2: { halign: 'left' }, // Customer Name
        3: { halign: 'left' }, // Invoice Number
        4: { halign: 'left' }, // Journal Type
        5: { halign: 'center' }, // Reference Date
        6: { halign: 'left' }, // Source Journal
        7: { halign: 'left' }, // Sale Company
        8: { halign: 'left' }  // Target JRN Company
      },
      didParseCell: (data: any) => {
        // Format Amount column
        if (data.section === 'body' && data.column.index === 1) {
          const val = parseFloat(data.cell.raw) || 0;
          data.cell.text = [val === 0 ? '-' : '$' + val.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })];
          data.cell.styles.textColor = val < 0 ? [220, 53, 69] : [0, 0, 0];
        }
        // Totals row styling
        if (data.row.index === body.length - 1) {
          data.cell.styles.fontStyle = 'bold';
          if (data.column.index !== 1) data.cell.text = [''];
        }
      }
    });

    return doc;
  }


  generatePDF() {
    const doc = this.createPDF();
    doc.save('Factory Incentive Receivables.pdf');
  }
  generatePDFBlob(): Blob | null {
    try {
      const doc = this.createPDF();
      return doc.output('blob');
    } catch (e) {
      console.error(e);
      this.toast.show('Error generating PDF', 'danger', 'Error');
      return null;
    }
  }
  public blobToFile = (theBlob: Blob, fileName: string): File => {
    return new File([theBlob], fileName, {
      lastModified: new Date().getTime(),
      type: theBlob.type,
    });
  };
  sendEmailData(Email: any, notes: any, from: any) {
    this.shared.spinner.show();
    try {
      const pdfBlob = this.generatePDFBlob(); // ✅ reuse

      if (!pdfBlob) {
        this.shared.spinner.hide();
        return;
      }
      const sizeMB = pdfBlob.size / (1024 * 1024);
      if (sizeMB > 10) {
        this.toast.show('PDF too large. Please reduce data.', 'warning', 'Warning');
        this.spinner.hide();
        return;
      }
      const pdfFile = this.blobToFile(pdfBlob, 'Factory Incentive Receivables.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Factory Incentive Receivables');
      formData.append('file', pdfFile);
      formData.append('notes', notes);
      formData.append('from', from);

      this.shared.api.postmethod(this.comm.routeEndpoint + 'mail', formData)
        .subscribe({
          next: (res: any) => {
            if (res.status === 200) {
              this.toast.show(res.response, 'success', 'Success');
            } else {
              this.toast.show('Invalid Details.', 'danger', 'Error');
            }
            this.shared.spinner.hide();
          },
          error: (error) => {
            console.error('Error:', error);
            this.shared.spinner.hide();
          }
        });

    } catch (error) {
      console.error(error);
      this.shared.spinner.hide();
    }
  }

  printPDF() {
    if (this.storeIds.length >= 6) {
      this.toast.show('Please select only up to 5 stores', 'warning', 'Warning');
      return;
    }
    try {
      const doc = this.createPDF();
     const pdfBlob = doc.output('blob');
      const pdfUrl = URL.createObjectURL(pdfBlob);
      const isEdge = /Edg/.test(navigator.userAgent);
      if (isEdge) {
        const win = window.open(pdfUrl);
        if (win) {
          win.onload = () => {
            setTimeout(() => {
              win.focus();
              win.onafterprint = () => {
                win.close();
                URL.revokeObjectURL(pdfUrl);
              };
              win.print();
            }, 500);
          };
        } else {
          this.toast.show('Popup blocked. Please allow popups.', 'warning', 'Warning');
        }
        return;
      }
      const iframe = document.createElement('iframe');
      iframe.style.position = 'fixed';
      iframe.style.right = '0';
      iframe.style.bottom = '0';
      iframe.style.width = '0';
      iframe.style.height = '0';
      iframe.style.border = 'none';
      document.body.appendChild(iframe);
      iframe.src = pdfUrl;
      iframe.onload = () => {
        setTimeout(() => {
          const cw = iframe.contentWindow;
          if (cw) {
            try {
              cw.focus();
              cw.onafterprint = () => {
                document.body.removeChild(iframe);
                URL.revokeObjectURL(pdfUrl);
              };
              cw.print();
            } catch (err) {
              console.warn('Iframe print failed, fallback triggered');
              const win = window.open(pdfUrl);
              if (win) {
                win.onload = () => {
                  setTimeout(() => {
                    win.focus();
                    win.onafterprint = () => {
                      win.close();
                      URL.revokeObjectURL(pdfUrl);
                    };
                    win.print();
                  }, 500);
                };
              }
            }
          }
        }, 500);
      };
    } catch (e) {
      console.error(e);
      this.toast.show('Error printing PDF', 'danger', 'Error');
    }
  }
  getSelectedStoreNames(): string {
    if (!this.storeIds || this.storeIds.length === 0) return '';
    const ids = this.storeIds.toString().split(',');
    const selectedStores = this.stores.filter((s: any) =>
      ids.includes(s.ID.toString())
    );
    return selectedStores.map((s: any) => s.storename).join(', ');
  }
  exportToExcel() {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Open Account Receivables');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 11,
        topLeftCell: 'A12',
        showGridLines: false,
      },
    ];
    // Title
    worksheet.addRow(['Open Account Receivables']).font = { size: 14, bold: true };
    worksheet.addRow([]);
    worksheet.addRow([new Date().toLocaleString()]);
    worksheet.addRow([]);
    worksheet.addRow(['Report Controls :']).font = { name: 'Arial', family: 4, size: 12, bold: true };
    worksheet.addRow([]);
    // Group and Stores info
    worksheet.getCell('A6').value = 'Group :';
    const groups = worksheet.getCell('B6');
    groups.value = this.groupName || 'All';
    groups.font = { name: 'Arial', size: 10 };
    groups.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

    worksheet.getCell('A7').value = 'Stores :';
    const stores = worksheet.getCell('B7');
    stores.value = this.getSelectedStoreNames() || 'All Stores';
    stores.font = { name: 'Arial', size: 10 };
    stores.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.mergeCells('B7', 'K9');

    worksheet.addRow([]);

    // Header row
    const headers = [
      'AR Company', 'Amount', 'Customer Name', 'Invoice Number', 'Journal Type',
      'Reference Date', 'Source Journal', 'Sale Company', 'Target JRN Company'
    ];
    const headerRow = worksheet.addRow(headers);

    headerRow.eachCell((cell, i) => {
      cell.font = { bold: true, color: { argb: 'FFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0554EF' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    const columnWidths = [20, 15, 30, 30, 15, 30, 15, 35, 30, 30];
    worksheet.columns.forEach((col, idx) => {
      col.width = columnWidths[idx];
    });

    // Data rows
    this.filteredFloorplanData.forEach(row => {
      const r = worksheet.addRow([
        row.arcompany || '-',
        row.amount || 0,
        row.customerName || '-',
        row.invoicenumber || '-',
        row.journaltype || '-',
        row.referencedate ? this.formatDateMMDDYYYY(row.referencedate) : '-',
        row.sourcejournal || '-',
        row.salecompany || '-',
        row.targetJRNcompany || '-'
      ]);

      r.eachCell((cell: any, i) => {
        cell.alignment = { horizontal: i === 2 ? 'right' : 'left', vertical: 'middle' };
        if (i === 8) {
          cell.numFmt = '"$" * #,##0.00;[Red]"$" * -#,##0.00';
          if (cell.value < 0) {
            cell.font = {
              name: 'Calibri',
              color: cell.value < 0 ? { argb: 'DC3545' } : { argb: '000000' }
            };
          }
        }
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    const totalRow = worksheet.addRow(['Total', this.BalanceTotal || 0]);
    totalRow.eachCell((cell, i) => {
      cell.font = { bold: true };
      cell.alignment = { horizontal: i === 2 ? 'right' : 'left' };
    });

    workbook.xlsx.writeBuffer().then(buffer => {
      const blob = new Blob([buffer]);
      FileSaver.saveAs(blob, 'Open Account Receivables_Report.xlsx');
    });
  }

}
