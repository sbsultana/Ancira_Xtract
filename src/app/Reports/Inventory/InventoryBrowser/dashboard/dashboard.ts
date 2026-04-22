import { Component, ElementRef, HostListener, ViewChild } from '@angular/core';
import { NgbActiveModal, NgbDateParserFormatter, NgbModal, NgbModule } from '@ng-bootstrap/ng-bootstrap';
import { Workbook } from 'exceljs';
import { DatePipe } from '@angular/common';
import html2canvas from 'html2canvas';

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
import { Notes } from '../../../../Layout/notes/notes';
import * as ExcelJS from 'exceljs';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, Stores, NgbModule, Notes],
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
  AgeTo: any = 300;
  QISearchName: any = '';

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

    if (path === 'Inventory Browser') {
      this.selectedreceviabe = this.LiabilitiesData[0];
    } else {
      let selectedpath = this.LiabilitiesData.find((e: any) => e.url === path);
      this.selectedreceviabe = selectedpath ? selectedpath : this.LiabilitiesData[0];
    }

    this.shared.setTitle(this.comm.titleName + '-Inventory Browser');

    /* --------- HEADER FOR REPORTING --------- */
    const data = {
      title: 'Inventory Browser',
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
    this.Getfloorplansdata();
  }

  /* ------------------------------- LIFECYCLE ------------------------------- */

  ngOnInit(): void { }
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Inventory Browser') {
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
      if (res.obj.Module === 'Inventory Browser') {
        document.getElementById('report')?.click();
      }
    });

    /* -------- REFRESH DATA FROM HEADER -------- */
    this.reportgetting = this.shared.api.GetReports().subscribe((data) => {
      if (data.obj.Reference === 'Inventory Browser') {
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
          this.Getfloorplansdata();
        } else {
          this.NoData = true;
        }

        /* Update header for next reload */
        const headerdata = {
          title: 'Inventory Browser',
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
      if (!obj || obj.title !== 'Inventory Browser') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Inventory Browser') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Inventory Browser') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Inventory Browser') return;
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

  StockType: any = ['N', 'U', 'C'];

  multipleorsingle(ref: any, e: any) {
    if (ref == 'ST') {
      const index = this.StockType.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.StockType.splice(index, 1);
      } else {
        this.StockType.push(e);
      }
    }
  }
  getStockLabel(code: string) {
    switch (code) {
      case 'N': return 'New';
      case 'U': return 'Used';
      case 'C': return 'CPO';
      default: return code;
    }
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
  Getfloorplansdata() {
    this.spinner.show();
    this.goToFirstPage();
    let obj = {};
    if (localStorage.getItem('IBObject') != null) {
      const InvObj = JSON.parse(localStorage.getItem('IBObject')!);
      obj = {
        DealerId: InvObj.invobj.store,
        DealType: InvObj.invobj.DealType,
        FromAge: InvObj.invobj.FromAge,
        ToAge: InvObj.invobj.ToAge,
        UserID: 0,
      };
    } else {
      obj = {
        DealerId: this.storeIds.toString(),
        DealType: this.StockType.toString(),
        FromAge: this.AgeFrom,
        ToAge: this.AgeTo,
        UserID: 0,
      };
    }
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetQuickInventory', obj)
      .subscribe({
        next: (res: any) => {
          this.spinner.hide();
          if (res?.status === 200 && Array.isArray(res.response)) {
            this.FloorPlanData = res.response.map((v: any) => ({
              ...v, notesView: '+'
            }));
            this.FloorPlanData.some(function (x: any) {
              if (x.Notes != undefined && x.Notes != null && x.Notes != '') {
                x.Notes = JSON.parse(x.Notes);
              }
              if (x.Comments != undefined && x.Comments != null && x.Comments != '') {
                x.Comments = JSON.parse(x.Comments);
              }
            });
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

  searchFilter() {
    this.pageNumber = 1;
    this.NoData = false;
    this.Getfloorplansdata();
  }

  onKeypressEvent(event: KeyboardEvent) {
    if (event.key === 'Enter') {
      this.pageNumber = 1;
      this.NoData = false;
      this.Getfloorplansdata();
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
      this.Getfloorplansdata();
    }
  }
  toggleView(data: any) {
    if (data.notesView == '+') {
      data.notesView = '-'
    } else {
      data.notesView = '+'
    }
  }
  notesViewState: boolean = true
  notesView() {
    this.notesViewState = !this.notesViewState
  }
  notesData: any = {}
  Notespopup: any
  addNotes(data: any, ref: any) {
    this.notesData = {
      // store: data.StoreID,
      // mainkey: data.StockNo,
      // module: 'IN',
      // apiRoute:'AddNotesAction'
      store: data.StoreID,
      title1: data.StockNo,
      title2: '',
      apiRoute: 'AddGeneralNotes',
       moduleCode:'IN'
    }
    this.Notespopup = this.ngbmodal.open(ref, { size: 'xxl', backdrop: 'static' });
  }
  callLoadingState = 'FL'
  CurrentPageSetting: number = 1
  closeNotes(e: any) {
    // this.ngbmodalActive.dismiss()
    this.Notespopup.close()
    if (e == 'S') {
      console.log(this.currentPage);
      this.CurrentPageSetting = this.currentPage
      this.callLoadingState = 'ANS'
      this.Getfloorplansdata();
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
    this.currentPage = 1;
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

        x.Store?.toString().toLowerCase().includes(term) ||
        x.StockNo?.toString().toLowerCase().includes(term) ||

        x.DealType?.toString().toLowerCase().includes(term) ||

        x.Year?.toString().toLowerCase().includes(term) ||

        x.Make?.toLowerCase().includes(term) ||
        x.Model?.toLowerCase().includes(term) ||

        x.Color?.toLowerCase().includes(term) ||
        x['Interior Color']?.toLowerCase().includes(term) ||

        x.Age?.toString().toLowerCase().includes(term) ||

        x.ListPrice?.toString().toLowerCase().includes(term) ||
        x.Internet?.toString().toLowerCase().includes(term) ||

        x.Miles?.toString().toLowerCase().includes(term) ||

        x.VIN?.toLowerCase().includes(term)
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

  get TotalListPrice(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.ListPrice) || 0);
    }, 0);
  }

  get TotalInternet(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.Internet) || 0);
    }, 0);
  }

  get TotalMiles(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.Miles) || 0);
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
    //   if (this.StockType.length == 3) {
    //     this.StockType = []
    //     this.toast.warning('Please Select Atleast One StockType', '');

    //   } else {
    //     this.StockType = []
    //     this.StockType = ['Booked', 'Finalized','Delivered'];
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
          const userName = JSON.parse(localStorage.getItem('userInfo')!).fullName;
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
        this.Getfloorplansdata();
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
    if (!this.StockType || this.StockType.length === 0) {
      this.toast.show('Please Select Atleast One Stock Type.', 'warning', 'Warning');
      missing = true;
    }
    if (missing) {
      return;
    }
    ((this.AgeFrom = this.AgeFrom), (this.AgeTo = this.AgeTo));
    this.Getfloorplansdata();
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


  exportToExcel(): void {
    const workbook: any = this.shared.getWorkbook?.();
    if (!workbook) {
      this.toast.show('Workbook helper not available', 'danger', 'Error');
      return;
    }

    /* ================= SAFE NAME HELPERS ================= */

    const getSafeSheetName = (name: string): string => {
      const cleaned = (name || 'Sheet1')
        .replace(/[\\/*?:[\]]/g, '')
        .substring(0, 31)
        .trim();

      return cleaned || 'Sheet1';
    };

    const getSafeFileName = (name: string): string => {
      return (name || 'Report')
        .replace(/[\\/*?:[\]]/g, '')
        .trim();
    };

    /* ================= REPORT TITLE ================= */

    const reportTitle =
      this.selectedreceviabe?.title
        ? `${this.selectedreceviabe.title} Inventory Browser`
        : 'Inventory Browser';

    const safeSheetName = getSafeSheetName(reportTitle);

    // Prevent duplicate sheet names
    let finalSheetName = safeSheetName;
    let counter = 1;
    while (workbook.getWorksheet(finalSheetName)) {
      finalSheetName = `${safeSheetName}_${counter++}`;
    }

    const summarySheet = workbook.addWorksheet(finalSheetName);

    /* ================= TITLE ================= */

    const titleRow = summarySheet.addRow([reportTitle]);
    titleRow.font = { bold: true, size: 14 };

    /* ================= FILTERS ================= */

    const filters = [
      { label: 'Age Range', value: `${this.AgeFrom || 0} - ${this.AgeTo || 0}` },
      { label: 'Inventory Browser', value: this.selectedreceviabe?.Liabilities || '-' },
    ];

    let startIndex = 2;
    filters.forEach((f) => {
      startIndex++;
      summarySheet.addRow([]);
      summarySheet.mergeCells(`B${startIndex}:C${startIndex}`);
      summarySheet.getCell(`A${startIndex}`).value = f.label;
      summarySheet.getCell(`B${startIndex}`).value = f.value;
    });

    summarySheet.addRow([]);

    /* ================= AGING HEADER ================= */

    const agingHeader = summarySheet.addRow(['', 'Total', '0-5', '6-10', '11-15', '15+']);
    agingHeader.font = { bold: true, color: { argb: 'FFFFFFFF' } };

    agingHeader.eachCell((cell: any) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2A91F0' } };
      cell.alignment = { horizontal: 'center' };
    });

    if (this.FloorPlanData?.length) {
      const A = this.FloorPlanData[0].AgeData?.[0];

      const agingValueRow = summarySheet.addRow([
        `${this.selectedreceviabe?.title || 'Inventory Browser'} Aging`,
        A?.TOTAL ?? null,
        A?.D1 ?? null,
        A?.D2 ?? null,
        A?.D3 ?? null,
        A?.D4 ?? null,
      ]);

      [2, 3, 4, 5, 6].forEach((col) => {
        const cell = agingValueRow.getCell(col);
        if (typeof cell.value === 'number') {
          cell.numFmt = '$#,##0.00;[Red]-$#,##0.00';
          cell.alignment = { horizontal: 'right' };
        }
      });

      agingValueRow.font = { bold: true };
    }

    summarySheet.addRow([]);

    /* ================= DETAIL HEADERS ================= */

    const headers = [
      'Age', 'Date', 'Account', 'Control', 'Control 2', 'Balance', 'Customer',
      'Number', 'Sale Date', 'Sale Age', 'Stock #', 'Deal #', 'Stage',
      'F&I Mgr', 'New/Used', 'Type', 'Status', 'Bank Name',
      'Year', 'Make', 'Model',
    ];

    const headerRow = summarySheet.addRow(headers);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };

    headerRow.eachCell((cell: any) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2A91F0' } };
      cell.alignment = { horizontal: 'center' };
    });

    /* ================= DATE FORMATTER ================= */

    const formatDate = (d: any) => {
      if (!d) return '-';
      const dt = new Date(d);
      if (isNaN(dt.getTime())) return '-';
      return (
        `${String(dt.getMonth() + 1).padStart(2, '0')}.` +
        `${String(dt.getDate()).padStart(2, '0')}.` +
        `${dt.getFullYear()}`
      );
    };

    /* ================= DATA ROWS ================= */

    (this.FloorPlanData || []).forEach((fp: any) => {
      const row = summarySheet.addRow([
        fp.AGE || '-',
        formatDate(fp.FundedDate),
        fp.AccountDesc2 || '-',
        fp.Control || '-',
        fp.control2 || '-',
        fp.Balance ?? null,
        fp.CustomerName || '-',
        fp.CustomerNumber || '-',
        formatDate(fp.SaleDate),
        fp.SaleAge || '-',
        fp.StockNo || '-',
        fp.DealNo || '-',
        fp.Stage || '-',
        fp.FIManager || '-',
        fp.DealType || '-',
        fp.SaleType || '-',
        fp.DealStatus || '-',
        fp.BankName || '-',
        fp.VehicleYear || '-',
        fp.VehicleMake || '-',
        fp.VehicleModel || '-',
      ]);

      const balanceCell = row.getCell(6);
      if (typeof balanceCell.value === 'number') {
        balanceCell.numFmt = '$#,##0.00;[Red]-$#,##0.00';
        balanceCell.alignment = { horizontal: 'right' };
      }

      if (fp.duplicateNotes?.length) {
        summarySheet.addRow([]);
        const notesHeader = summarySheet.addRow(['Notes']);
        notesHeader.font = { bold: true };

        fp.duplicateNotes.forEach((n: any) => {
          summarySheet.addRow([n.NOTES || '-']);
        });
      }
    });

    summarySheet.columns.forEach((c: any) => (c.width = 20));

    /* ================= TOTALS ROW ================= */

    if (this.FloorPlanTotalData?.length > 0) {
      const totalBalance = this.FloorPlanTotalData[0].Balance ?? 0;

      summarySheet.addRow([]);

      const totalsRow = summarySheet.addRow([
        'Totals', '', '', '', '',
        totalBalance,
        '', '', '', '', '', '', '', '', '', '', '', '', '', '',
      ]);

      totalsRow.font = { bold: true };
      summarySheet.mergeCells(`A${totalsRow.number}:E${totalsRow.number}`);
      totalsRow.getCell(1).alignment = { horizontal: 'right' };

      const balanceCell = totalsRow.getCell(6);
      balanceCell.numFmt = '$#,##0.00;[Red]-$#,##0.00';
      balanceCell.alignment = { horizontal: 'right' };
    }

    /* ================= EXPORT ================= */

    workbook.xlsx.writeBuffer().then(() => {
      const rawFileName =
        `${this.selectedreceviabe?.title || 'Inventory Browser'} Title Tracking_Report`;

      const safeFileName = getSafeFileName(rawFileName);

      this.shared.exportToExcel(workbook, safeFileName);
    });
  }

  private createPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });

    doc.setFontSize(14);
    doc.text('Inventory Browser', 14, 12);

    const headers = [[
      'Store',
      'Stock No',
      'Stock Type',
      'Year',
      'Make',
      'Model',
      'Color',
      'Interior Color',
      'Age',
      'List Price',
      'MSRP',
      'Miles',
      'VIN'
    ]];

    const rows: any[] = [];

    this.filteredData.forEach((QI: any) => {

      rows.push([
        QI.Store || '-',
        QI.StockNo || '-',
        QI.DealType === 'N' ? 'New' :
          QI.DealType === 'U' ? 'Used' :
            QI.DealType ? 'CPO' : '-',
        QI.Year || '-',
        QI.Make || '-',
        QI.Model || '-',
        QI.Color || '-',
        QI['Interior Color'] || '-',
        QI.Age || '-',
        QI.ListPrice ?? '',
        QI.Internet ?? '',
        QI.Miles ?? '',
        QI.VIN || '-'
      ]);

    });

    rows.push([
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      'Totals',
      this.TotalListPrice ?? 0,
      this.TotalInternet ?? 0,
      this.TotalMiles ?? 0,
      ''
    ]);

    autoTable(doc, {
      head: headers,
      body: rows,
      startY: 18,

      styles: {
        fontSize: 8,
        cellPadding: 2,
        valign: 'middle',
        halign: 'center',
        lineWidth: 0.3,
        lineColor: [200, 200, 200]
      },

      headStyles: {
        fillColor: [30, 80, 200],
        textColor: 255,
        fontStyle: 'bold'
      },

      columnStyles: {
        0: { halign: 'left' },
        1: { halign: 'left' },
        2: { cellWidth: 20, halign: 'center' },
        3: { halign: 'center' },
        4: { halign: 'left' },
        5: { halign: 'left' },
        6: { halign: 'left' },
        7: { halign: 'left' },
        8: { halign: 'center' },
        9: { cellWidth: 30, halign: 'right' },
        10: { cellWidth: 30, halign: 'right' },
        11: { cellWidth: 30, halign: 'right' },
        12: { halign: 'left' }
      },

      didParseCell: (data: any) => {

        if (data.section !== 'body') return;
        const val = data.cell.raw;
        const col = data.column.index;
        if (data.row.index === rows.length - 1) {
          data.cell.styles.fillColor = [220, 235, 255];
          data.cell.styles.fontStyle = 'bold';
          if ((col === 9 || col === 10) && val !== '' && val !== null && !isNaN(val)) {
            data.cell.text = [`$ ${Number(val).toLocaleString()}`];
            data.cell.styles.halign = 'right';
          }
          else if (col === 11 && val !== '' && val !== null && !isNaN(val)) {
            data.cell.text = [Number(val).toLocaleString()];
            data.cell.styles.halign = 'right';
          }
          return;
        }
        if ((col === 9 || col === 10) && val !== '' && val !== null && !isNaN(val)) {
          data.cell.text = [`$ ${Number(val).toLocaleString()}`];
          data.cell.styles.halign = 'right';
        }
        else if (col === 11 && val !== '' && val !== null && !isNaN(val)) {
          data.cell.text = [Number(val).toLocaleString()];
          data.cell.styles.halign = 'right';
        }
        else if (col === 0) {
          data.cell.styles.halign = 'center';
        }
        else {
          data.cell.styles.halign = 'left';
        }
      }
    });

    return doc;
  }

  generatePDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save('Inventory Browser.pdf'); // ✅ only here
    } catch (e) {
      console.error(e);
      this.toast.show('Error generating PDF', 'danger', 'Error');
    } finally {
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
    this.spinner.show();
    try {
      const pdfBlob = this.generatePDFBlob();

      if (!pdfBlob) {
        this.spinner.hide();
        return;
      }
      const sizeMB = pdfBlob.size / (1024 * 1024);
      if (sizeMB > 10) {
        this.toast.show('PDF too large. Please reduce data.', 'warning', 'Warning');
        this.spinner.hide();
        return;
      }
      const pdfFile = this.blobToFile(pdfBlob, 'Inventory Browser.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Inventory Browser');
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
            this.spinner.hide();
          },
          error: (error) => {
            console.error('Error:', error);
            this.spinner.hide();
          }
        });

    } catch (error) {
      console.error(error);
      this.spinner.hide();
    }
  }
}
