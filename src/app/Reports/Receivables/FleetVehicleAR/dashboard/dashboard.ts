import { Component, HostListener } from '@angular/core';
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
  imports: [SharedModule, BsDatepickerModule, TimeConversionPipe, Stores, NgbModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
  standalone: true,
})
export class Dashboard {
  /* ------------------------------- VARIABLES ------------------------------- */

  NoData: boolean = false;
  FloorPlanData: any = [];
  FloorPlanTotalData: any = [];

  dealStatus: any = ['Capped', 'Delivered'];
  dealType: any = ['Retail', 'Fleet', 'Wholesale'];
  allordebit: any = 'dr';
  selectedreceviabe: any = [];
  StoreVal: any = '71,53,8,7,4,35,1,32,40,50,25,18,31,3,70,72,2,17,41,55,42,51,12,73,54,9,15,5,14,30,11';
  spinnerLoader: boolean = false;
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
  actionType: any = 'N';

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
  DefaultLoad: any = 'E'
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
    storedisplayname: this.storedisplayname, 'DefaultLoad': this.DefaultLoad
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
      this.actionType = 'Y';
      localStorage.setItem('flag', 'M')
    } else {
      if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
        this.storeIds = ''
        this.StoreVal = ""
        this.actionType = 'N';
      }
    }
    this.commentsVisibility = true;

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

    this.shared.setTitle(this.comm.titleName + '-Fleet Vehicle');

    /* --------- HEADER FOR REPORTING --------- */
    const data = {
      title: 'Fleet Vehicle',
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

    /* --------- AUTO LOAD FIRST TAB --------- */
    if (this.StoreVal != '') {
      this.Getfloorplansdata(this.selectedreceviabe);
      this.getEmployees();
    }
  }

  /* ------------------------------- LIFECYCLE ------------------------------- */

  ngOnInit(): void { }
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Fleet Vehicle') {
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
      if (res.obj.Module === 'Fleet Vehicle') {
        document.getElementById('report')?.click();
      }
    });

    /* -------- REFRESH DATA FROM HEADER -------- */
    this.reportgetting = this.shared.api.GetReports().subscribe((data) => {
      if (data.obj.Reference === 'Fleet Vehicle') {
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
          this.Getfloorplansdata(this.selectedreceviabe);
        } else {
          this.NoData = true;
        }

        /* Update header for next reload */
        const headerdata = {
          title: 'Fleet Vehicle',
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
      if (!obj || obj.title !== 'Fleet Vehicle') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Fleet Vehicle') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Fleet Vehicle') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Fleet Vehicle') return;
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


  async openSalesModal(dealnumber: any, vin: any, storeid: any, stock: any, source: any, custno: any) {
    const module = await import('../../../../Layout/cdpdataview/deal/deal-module');
    const component = module.Deal;

    const modalRef = this.shared.ngbmodal.open(component, { size: 'xl', windowClass: 'connectedmodal' });
    modalRef.componentInstance.data = { dealno: dealnumber, vin: vin, storeid: storeid, stock: stock, source: source, custno: custno }; // Pass data to the modal component    
    modalRef.result.then((result) => {
      // this.topScroll()
      console.log(result); // Handle modal close result
    }, (reason) => {
      // this.topScroll()
      console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    });
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

  // sort(property: any, state?: any) {
  //   if (state === undefined) {
  //     this.isDesc = !this.isDesc;
  //   }

  //   this.callLoadingState = 'FL';
  //   this.column = property;

  //   let direction = this.isDesc ? 1 : -1;

  //   this.FloorPlanData.sort((a: any, b: any) => {
  //     if (a[property] < b[property]) return -1 * direction;
  //     if (a[property] > b[property]) return 1 * direction;
  //     return 0;
  //   });
  // }

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

      // ✅ Numeric comparison when both are numbers
      if (isNumA && isNumB) {
        return (numA - numB) * direction;
      }

      // ✅ String comparison fallback
      return String(valA).localeCompare(String(valB), undefined, {
        numeric: true,
        sensitivity: 'base'
      }) * direction;
    });
  }

  /* ------------------------------- MAIN API CALL ------------------------------- */
  previousReportPath: string | null = null;
  Getfloorplansdata(path: any) {
    this.goToFirstPage();
    this.DefaultLoad = ''
    if (this.previousReportPath !== path.path) {
      this.AgeFrom = 0;
      this.AgeTo = 0;
      this.previousReportPath = path.path;
    }
    this.selectedreceviabe = path;
    this.NoData = false;
    this.FloorPlanData = [];
    this.FloorPlanTotalData = [];
    this.filteredFloorplanData = [];
    if (
      this.AgeFrom !== null &&
      this.AgeTo !== null &&
      Number(this.AgeFrom) > Number(this.AgeTo)
    ) {
      this.toast.show(
        'Please Enter Valid Age Range',
        'warning',
        'Warning'
      );
      return; // ⛔ stop execution
    }
    this.spinner.show();


    console.log('hi')

    const obj = {
      AS_ID: this.storeIds.toString(),
      // DealStatus: this.dealStatus.toString(),
      FIManagerID: this.financeManagerId,
      UserID: "1",
      // DealType2: this.dealType.toString(),
    };

    let startFrom = new Date().getTime();

    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetCITVehicleReport', obj).subscribe(
      (res) => {
        if (res.status == 200 && res.response) {
          this.spinner.hide();

          if (res.response.length > 0) {
            this.FloorPlanData = res.response.filter((e: any) => e.store !== 'TOTAL');
            this.FloorPlanTotalData = res.response.filter((e: any) => e.store === 'TOTAL');
            this.FloorPlanData.forEach((x: any) => {
              x.AgeData = JSON.parse(x.AgeData);

              if (x.Comments) x.Comments = JSON.parse(x.Comments);
              if (x.Notes) {
                x.Notes = JSON.parse(x.Notes);

                x.allNotes = [...x.Notes];        // ✅ full data
                x.duplicateNotes = x.allNotes.slice(0, 3); // ✅ first 3 only

                x.Notesstate = '+';
              }
            });

            if (this.callLoadingState == 'ANS') {
              this.sort(this.column, this.callLoadingState);
            }
            this.filteredFloorplanData = this.FloorPlanData || [];
            this.NoData = this.FloorPlanData.length == 0;
          } else {
            this.NoData = true;
          }
        } else {
          this.NoData = true;
          this.spinner.hide();
        }
      },
      () => {
        this.toast.show(
          '502 Bad Gateway Error',
          'danger',
          'Error'
        );
        this.spinner.hide();
        this.NoData = true;
      },
    );
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
      searchTerms.some(term => {

        const t = term.toLowerCase();

        return (

          // ===== BASIC =====
          this.normalize(x.CTAge).includes(t) ||
          this.normalize(x.CTDate).includes(t) ||
          this.normalize(x.Dealno).includes(t) ||

          this.normalize(x.Control).includes(t) ||
          this.normalize(x.Control2).includes(t) ||

          // ===== BALANCES =====
          this.normalize(x.Balance).includes(t) ||
          this.normalize(x.CIT).includes(t) ||
          this.normalize(x.CITAccount).includes(t) ||
          this.normalize(x.CashOSF).includes(t) ||
          this.normalize(x.CashOSFAccount).includes(t) ||
          this.normalize(x['W/O Credits']).includes(t) ||

          // ===== CUSTOMER =====
          this.normalize(x.CustomerName).includes(t) ||
          this.normalize(x.CustomerNo).includes(t) ||

          // ===== STORE =====
          this.normalize(x.store).includes(t) ||

          // ===== DEAL INFO =====
          this.normalize(x.Stage).includes(t) ||
          this.normalize(x.Stockno).includes(t) ||
          this.normalize(x.FIManager).includes(t) ||
          this.normalize(x.DealType).includes(t) ||
          this.normalize(x.SaleType).includes(t) ||
          this.normalize(x.DealStatus).includes(t) ||
          this.normalize(x.BankName).includes(t) ||

          // ===== VEHICLE =====
          this.normalize(x.VehicleYear).includes(t) ||
          this.normalize(x.VehicleMake).includes(t) ||
          this.normalize(x.VehicleModel).includes(t) ||

          // ===== DATES =====
          this.normalize(x.SaleDate).includes(t) ||
          this.normalize(x.FundedDate).includes(t)

        );
      })
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
  normalize(value: any): string {
    if (value === null || value === undefined) return '';
    return value.toString().toLowerCase().trim();
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

  get CIT(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.CIT) || 0);
    }, 0);
  }

  get Balance(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.Balance) || 0);
    }, 0);
  }

  get CashOSF(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.CashOSF) || 0);
    }, 0);
  }

  get WOCredits(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item['W/O Credits']) || 0);
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
      'type': 'M', 'others': 'N', 'DefaultLoad': this.DefaultLoad
    };


    // this.setHeaderData();
    // this.GetData();

  }
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.StoreVal = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
    this.getEmployees()
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


  multipleorsingle(ref: any, e: any) {
    if (ref == 'DS') {
      const index = this.dealStatus.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.dealStatus.splice(index, 1);
      } else {
        this.dealStatus.push(e);
      }
    }
    if (ref == 'DT') {
      const index = this.dealType.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.dealType.splice(index, 1);
      } else {
        this.dealType.push(e);
      }
    }
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
      "AssociatedReport": "CIT",
      "CompanyID": companyid
    }
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
      this.toast.show(
        'Please enter notes',
        'warning',
        'Warning'
      );
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
          this.toast.show(
            'Notes Added Successfully',
            'success',
            'Success'
          )
          this.callLoadingState = 'ANS';
          (document.getElementById('close') as HTMLInputElement)?.click();
          this.oncloseone();
          this.commentsVisibility = true;

          /* Instant Frontend Update */
          const userName = JSON.parse(localStorage.getItem('userInfo')!).fullName
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
          this.toast.show(
            'Something went wrong. Please try again.',
            'danger',
            'Error'
          );
        }
      });
  }

  /* ------------------------------- HIDE RECORD LOGIC ------------------------------- */

  collectHidevalues(e: any, val: any, confirmtemplate: any, ref: any, refval: any) {
    if (ref === 'multi') {
      if (this.hideRecords.length === 0) {
        this.toast.show(
          'Please Select Atleast One Record to Hide',
          'warning',
          'Warning'
        );
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
      this.toast.show(
        'Please Select Atleast One Record to Hide',
        'warning',
        'Warning'
      );
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
        this.toast.show(
          'This Control Hidden Successfully',
          'success',
          'Success'
        );
        (document.getElementById('closeone') as HTMLElement).click();
        this.oncloseone();
        this.Getfloorplansdata(this.selectedreceviabe);
        this.hideRecords = [];
        this.hideVisibility = false;
      } else {
        this.toast.show(
          'Failed to hide control.',
          'danger',
          'Error'
        );
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

  hidestoreIds: any = '';
  viewreport() {
    this.activePopover = -1;
    let missing = false;
    if (!this.dealStatus || this.dealStatus.length === 0) {
      this.toast.show('Please Select Atleast One Deal Status.', 'warning', 'Warning');
      missing = true;
    }
    if (!this.dealType || this.dealType.length === 0) {
      this.toast.show('Please Select Atleast One Deal Type.', 'warning', 'Warning');
      missing = true;
    }
    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show('Please Select Atleast One Store.', 'warning', 'Warning');
      missing = true;
    }
    if (!this.selectedFiManagersvalues || this.selectedFiManagersvalues.length === 0) {
      this.toast.show('Please Select Atleast One People.', 'warning', 'Warning');
      missing = true;
    }
    if (missing) {
      return;
    }
    this.hidestoreIds = Array.isArray(this.storeIds)
      ? [...this.storeIds]
      : this.storeIds;
    this.financeManagerId =
      this.selectedFiManagersvalues.length === this.financeManager.length
        ? '0'
        : this.selectedFiManagersvalues.toString();
    this.AgeFrom = this.AgeFrom;
    this.AgeTo = this.AgeTo;
    const payload = {
      DealStatus: this.dealStatus.toString(),
      DealType2: this.dealType.toString(),
      StoreIds: this.storeIds.toString(),
      FinanceManagerId: this.financeManagerId,
    };
    this.actionType = 'Y'
    this.Getfloorplansdata(this.selectedreceviabe);
  }

  getEmployees(val?: any, ids?: any, count?: any, bar?: any) {
    const obj = {
      AS_ID: this.storeIds.toString(),
      type: 'F',
    };
    this.spinnerLoader = true;
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetEmployeesDev', obj).subscribe(
      (res: any) => {
        if (res && res.status == 200) {
          this.spinnerLoader = false;
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
          this.spinnerLoader = false;
          this.toast.show(
            'Invalid Details',
            'danger',
            'Error'
          );

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


  showStoreColumn(): boolean {
    return this.hidestoreIds?.length > 1;
  }

  isValidDealNo(value: any): boolean {
    return value !== null && value !== undefined && value !== '' && value !== 0;
  }
  formatDollar(value: any): string {
    if (value === 0 || value === null || value === undefined) {
      return '-';
    }
    return '$' + Number(value).toLocaleString('en-US', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
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
  expandedNotes: { [key: number]: boolean } = {};
  private createPDF(): jsPDF {
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a2' });
    doc.setFontSize(14);
    doc.text('Fleet Vehicle', 14, 12);

    const aging = this.FloorPlanData?.[0]?.AgeData?.[0] ?? {};
    autoTable(doc, {
      startY: 16,
      theme: 'grid',
      margin: { left: 6, right: 6 },
      tableWidth: 'auto',
      styles: {
        fontSize: 10,
        cellPadding: 2,
        halign: 'center',
        textColor: [0, 0, 0]
      },
      headStyles: {
        fillColor: [5, 84, 239],
        textColor: 255,
        fontStyle: 'bold'
      },
      head: [['Booked Aging', 'TOTAL', '0-5', '6-10', '11-15', '15+']],
      body: [[
        '',
        this.formatDollar(aging?.TOTAL),
        this.formatDollar(aging?.D1),
        this.formatDollar(aging?.D2),
        this.formatDollar(aging?.D3),
        this.formatDollar(aging?.D4)
      ]],
      didParseCell: (data) => {
        if (data.section === 'body') {
          const val = parseFloat(String(data.cell.raw).replace(/[$,]/g, '')) || 0;
          data.cell.styles.textColor = val < 0 ? [220, 53, 69] : [0, 0, 0];
        }
      }
    });
    const head2 = [
      'Age', 'Date', 'Deal #', 'Control',
      'Balance',
      'Name', 'Number',
      'Store',
      'Stock #', 'Deal Type', 'Stage', 'F&I Mgr', 'Sales Mgr', 'Sales Person',
      'Type', 'New/Used', 'Status', 'Bank Name',
      'Trade'
    ];

    const alignments = [
      'left', 'center', 'center', 'center',
      'right',
      'left', 'center',
      'left',
      'left', 'center', 'left', 'left', 'left', 'left',
      'center', 'center', 'center', 'left',
      'left'
    ];
    const rows: any[] = [];

    (this.filteredData || []).forEach((fp: any, i: number) => {

      const rowData = [
        fp.CTAge || '-',
        fp.CTDate ? this.formatDateMMDDYYYY(fp.CTDate) : '-',
        fp.Dealno || '-',
        fp.Control || '-',

        fp.Balance || 0,

        fp.CustName || '-',
        fp.CustId || '-',

        fp.store || '-',

        fp.Stockno || '-',
        fp.DealType2 || '-',
        fp.Stage || '-',
        fp.FIMgr_Name || '-',
        fp.SalesMgr_Name || '-',
        fp.SP1_Name || '-',

        fp.SaleType || '-',
        fp.DealType || '-',
        fp.DealStatus || '-',
        fp.Lender || '-',

        `${fp.trade1year || ''} ${fp.trade1modelname || '-'}`
      ];

      rows.push(rowData);
      if (fp.duplicateNotes && this.commentsVisibility) {
        const colSpan = head2.length;

        rows.push([{
          content: 'Notes:',
          colSpan,
          styles: {
            fillColor: [217, 231, 255],
            textColor: [5, 84, 239],
            halign: 'left',
            fontStyle: 'bold',
            fontSize: 8
          }
        }, ...Array(colSpan - 1).fill({})]);

        fp.duplicateNotes.forEach((note: any) => {
          rows.push([{
            content: '   ' + (note?.NOTES || '-'),
            colSpan,
            styles: {
              fillColor: [243, 246, 255],
              halign: 'left',
              fontSize: 8
            }
          }, ...Array(colSpan - 1).fill({})]);
        });
      }
    });
    const borderStyle = {
      lineWidth: { left: .3, right: .3, top: 0, bottom: 0 },
      lineColor: [180, 180, 180]
    };

    const head1 = [
      { content: '', colSpan: 4, styles: { fillColor: [5, 84, 239], ...borderStyle } },
      { content: 'Balances', colSpan: 1, styles: { fillColor: [5, 84, 239], textColor: 255, halign: 'center', ...borderStyle } },
      { content: 'Customer', colSpan: 2, styles: { fillColor: [5, 84, 239], textColor: 255, halign: 'center', ...borderStyle } },
      { content: 'Store', colSpan: 1, styles: { fillColor: [5, 84, 239], textColor: 255, halign: 'center', ...borderStyle } },
      { content: 'Deal Info', colSpan: 10, styles: { fillColor: [5, 84, 239], textColor: 255, halign: 'center', ...borderStyle } },
      { content: '', colSpan: 1, styles: { fillColor: [5, 84, 239], ...borderStyle } }
    ];
    const baseWidths = [
      12, 15, 20, 20,
      18,
      30, 20,
      20,
      18, 18, 18, 22, 22, 22,
      18, 18, 18, 28,
      25
    ];

    const pageW = doc.internal.pageSize.getWidth();
    const printableW = pageW - 12;
    const scale = printableW / baseWidths.reduce((a, b) => a + b, 0);

    const columnStyles = baseWidths.reduce((acc: any, w, i) => {
      acc[i] = { cellWidth: Math.max(8, w * scale) };
      return acc;
    }, {});
    autoTable(doc, {
      startY: (doc as any).lastAutoTable?.finalY + 8,
      head: [head1 as any, head2 as any],
      body: rows,
      theme: 'grid',
      margin: { left: 6, right: 6 },

      styles: {
        fontSize: 8,
        cellPadding: 1.5,
        overflow: 'linebreak',
        valign: 'top',
        textColor: [0, 0, 0]
      },

      headStyles: {
        fillColor: [217, 231, 255],
        textColor: [0, 0, 0],
        fontStyle: 'bold',
        halign: 'center'
      },

      columnStyles,

      didParseCell: (data: any) => {

        if (data.section === 'body') {

          // Alignment
          const align = alignments[data.column.index];
          if (align) data.cell.styles.halign = align;

          // Currency formatting (Balance column)
          if (data.column.index === 4) {
            const val = data.cell.raw;
            if (typeof val === 'number') {
              data.cell.text = [
                `$ ${val.toLocaleString('en-US', {
                  minimumFractionDigits: 2,
                  maximumFractionDigits: 2
                })}`
              ];
              data.cell.styles.halign = 'right';

              if (val < 0) {
                data.cell.styles.textColor = [220, 53, 69];
              }
            }
          }

          // Notes styling
          if (data.row.raw[0]?.content?.startsWith('Notes')) {
            data.cell.styles.fontStyle = 'bold';
            data.cell.styles.textColor = [5, 84, 239];
            data.cell.styles.halign = 'left';
          }
        }
      }
    });

    return doc;
  }
  generatePDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save(`Fleet Vehicle.pdf`);
    } catch (e) {
      console.error(e);
      this.toast.show('Error generating PDF', 'danger', 'Error');
    } finally {
      this.shared.spinner.hide();
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
      const pdfBlob = this.generatePDFBlob(); // ✅ reuse

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
      const pdfFile = this.blobToFile(pdfBlob, 'Fleet Vehicle.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Fleet Vehicle');
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
  printPDF() {
    if (this.storeIds.length >= 6) {
      this.toast.show('Please Select Only 5 Stores', 'warning', 'Warning');
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

  private exportToExcel(): void {

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Fleet Vehicle');

    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 15,
        topLeftCell: 'A16',
        showGridLines: false,
      },
    ];
    const titleRow = worksheet.addRow(['Fleet Vehicle']);
    titleRow.font = { size: 16, bold: true };
    worksheet.mergeCells(1, 1, 1, 19);
    titleRow.getCell(1).alignment = { horizontal: 'left' };

    worksheet.addRow([]);

    worksheet.addRow(['Report Controls :']).font = { name: 'Arial', family: 4, size: 12, bold: true };
    worksheet.addRow([]);

    worksheet.getCell('A6').value = 'Group :';
    const groups = worksheet.getCell('B6');
    groups.value = this.groupName || 'All';
    groups.font = { name: 'Arial', family: 4, size: 10 };
    groups.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

    worksheet.mergeCells('B7', 'K9');
    worksheet.getCell('A7').value = 'Stores :';
    const stores = worksheet.getCell('B7');
    stores.value = this.getSelectedStoreNames() || 'All Stores';
    stores.font = { name: 'Arial', family: 4, size: 10 };
    stores.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };

    worksheet.addRow([]);
    const aging = this.FloorPlanData?.[0]?.AgeData?.[0] ?? {};

    const agingHeader = worksheet.addRow(['Booked Aging', 'TOTAL', '0-5', '6-10', '11-15', '15+']);
    agingHeader.eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0554EF' } };
      cell.alignment = { horizontal: 'center' };
      cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
    });

    const agingRow = worksheet.addRow([
      '',
      this.formatDollar(aging?.TOTAL),
      this.formatDollar(aging?.D1),
      this.formatDollar(aging?.D2),
      this.formatDollar(aging?.D3),
      this.formatDollar(aging?.D4)
    ]);

    agingRow.eachCell(cell => {
      cell.alignment = { horizontal: 'center' };
      cell.border = { bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
    });

    worksheet.addRow([]);

    /* ================= GROUP HEADER ================= */
    const groupHeaders = [
      { label: '', span: 4 },
      { label: 'Balances', span: 1 },
      { label: 'Customer', span: 2 },
      { label: 'Store', span: 1 },
      { label: 'Deal Info', span: 10 },
      { label: 'Trade', span: 1 }
    ];

    let col = 1;
    const groupRow = worksheet.addRow([]);

    groupHeaders.forEach(g => {
      const cell = groupRow.getCell(col);
      cell.value = g.label;
      cell.font = { bold: true, color: { argb: 'FFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0554EF' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };

      if (g.span > 1) {
        worksheet.mergeCells(groupRow.number, col, groupRow.number, col + g.span - 1);
      }
      col += g.span;
    });

    /* ================= COLUMN HEADERS ================= */
    const headers = [
      'Age', 'Date', 'Deal #', 'Control',
      'Balance',
      'Name', 'Number',
      'Store',
      'Stock #', 'Deal Type', 'Stage', 'F&I Mgr', 'Sales Mgr', 'Sales Person',
      'Type', 'New/Used', 'Status', 'Bank Name',
      'Trade'
    ];

    const headerRow = worksheet.addRow(headers);

    headerRow.eachCell(cell => {
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9E7FF' } };
      cell.alignment = { horizontal: 'center' };
      cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
    });

    /* ================= ALIGNMENTS ================= */
    const alignments = [
      'center', 'center', 'center', 'center',
      'right',
      'left', 'center',
      'left',
      'left', 'center', 'left', 'left', 'left', 'left',
      'center', 'center', 'center', 'left',
      'left'
    ];

    /* ================= DATA ================= */
    (this.filteredData || []).forEach((fp: any, rowIndex: number) => {
      const rowData = [
        fp.CTAge || '-',
        fp.CTDate ? this.formatDateMMDDYYYY(fp.CTDate) : '-',
        fp.Dealno || '-',
        fp.Control || '-',

        fp.Balance || 0,

        fp.CustName || '-',
        fp.CustId || '-',

        fp.store || '-',

        fp.Stockno || '-',
        fp.DealType2 || '-',
        fp.Stage || '-',
        fp.FIMgr_Name || '-',
        fp.SalesMgr_Name || '-',
        fp.SP1_Name || '-',

        fp.SaleType || '-',
        fp.DealType || '-',
        fp.DealStatus || '-',
        fp.Lender || '-',

        `${fp.trade1year || ''} ${fp.trade1modelname || '-'}`
      ]
      const row = worksheet.addRow(rowData);

      row.eachCell((cell: any, i) => {
        cell.alignment = { horizontal: alignments[i - 1], vertical: 'middle' };

        // Currency
        if (i === 5) {
          cell.numFmt = '"$" * #,##0.00;[Red]"$" * -#,##0.00';
          if (cell.value < 0) {
            cell.font = { color: { argb: 'DC3545' } };
          }
        }

        // Borders
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };
      });

      if (fp.Notes && this.commentsVisibility) {
        const notes = Array.isArray(fp.Notes) ? fp.Notes : [fp.Notes];
        const notesRow = worksheet.addRow(['Notes:']);
        worksheet.mergeCells(notesRow.number, 1, notesRow.number, rowData.length);
        const cell = notesRow.getCell(1);
        cell.font = { bold: true, color: { argb: '0554EF' }, size: 11 };
        cell.alignment = { horizontal: 'left', vertical: 'top' };
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };
        notes.forEach((note: any) => {
          const nRow = worksheet.addRow(['   ' + (note?.NOTES || 'No notes')]);
          worksheet.mergeCells(nRow.number, 1, nRow.number, rowData.length);
          const nCell = nRow.getCell(1);
          nCell.font = { size: 11 };
          nCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
          nCell.border = {
            left: { style: 'thin' },
            right: { style: 'thin' },
            bottom: { style: 'thin' }
          };
        });
      }
    });

    /* ================= WIDTH ================= */
    const widths = [
      15, 15, 20, 20,
      18,
      30, 20,
      20,
      18, 18, 18, 22, 22, 22,
      18, 18, 18, 28,
      25
    ];

    widths.forEach((w, i) => worksheet.getColumn(i + 1).width = w);

    /* ================= DOWNLOAD ================= */
    workbook.xlsx.writeBuffer().then(buffer => {
      const blob = new Blob([buffer]);
      FileSaver.saveAs(blob, 'Booked_Deals.xlsx');
    });
  }
}
