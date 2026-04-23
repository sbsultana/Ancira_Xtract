

import { Component, ElementRef, ViewChild, HostListener, OnInit, EventEmitter, Output, Input, Renderer2 } from '@angular/core';
import { NgbActiveModal, NgbDateParserFormatter, NgbModal } from '@ng-bootstrap/ng-bootstrap';
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
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import * as FileSaver from 'file-saver';
import autoTable from 'jspdf-autotable';
import { common } from '../../../../common';
// import { ToastrService } from 'ngx-toastr';

// declare let require: any;
// const pdfMake = require('pdfmake/build/pdfmake');
// const pdfFonts = require('pdfmake/build/vfs_fonts');
// (pdfMake as any).vfs =
//   (pdfFonts as any).vfs ||
//   ((pdfFonts as any).pdfMake && (pdfFonts as any).pdfMake.vfs);

const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, TimeConversionPipe, Stores],
  templateUrl: './dashboard.html',
  styleUrls: ['./dashboard.scss'],
})
export class Dashboard implements OnInit {
  // Dates & filters (appointment-style)
  FromDate: any;
  ToDate: any;
  CurrentDate = new Date();

  // user / store
  storeIds: any = '';
  Role: any = [];
  userid: any;

  // UI state
  spinnerLoader: boolean = false;
  enablevehicle: any = false;
  vehiclear: any = 'WOAR';
  noData: boolean = false;
  NoData: boolean = false;
  LogCount = 1;
  solutionurl: any = (window as any)['environment']?.apiUrl || '';
  groups: any = 1;
  callLoadingState = 'FL';
  check: boolean = false;
  Viewmore: boolean = false;

  // Data
  FloorPlanData: any = [];
  FloorPlanTotalData: any = [];
  TotalFloorPlanData: any = [];
  QISearchName: any = '';
  commentdata: any = [];
  notesstage: any = [];

  // filter defaults
  dealStatus: any = ['Booked', 'Finalized', 'Delivered'];
  saleType: any = ['Retail', 'Fleet'];
  tempPriorityFilter: any = ['All'];
  allordebit: any = 'dr';
  financeManagerId: any = '0';
  priorityRecords: any = [];
  priorityVisibility: boolean = false;
  prioritycheck: boolean = false;
  originalFloorPlanData: any[] = [];
  // header & widgets
  header: any = [
    {
      type: 'Bar',
      storeIds: this.storeIds,
      vehiclear: this.vehiclear,
      dealStatus: this.dealStatus,
      saleType: this.saleType,
      allordebit: this.allordebit,
      groups: this.groups,
      financemanagers: this.financeManagerId,
    },
  ];
  popup: any = [{ type: 'Popup' }];

  // notes / hide records
  notesStageValue: any = '';
  notesStageValueGrid: any = '';
  notesStageText: any = '';
  selecteddata: any = [];
  FinalArray: any = [];
  hideRecords: any = [];

  // misc
  today: any;
  startDate: any;
  endDate: any;
  column: string = 'CategoryName';
  isDesc: boolean = false;
  callExportState!: Subscription;
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  CITstate: any;

  // view helpers
  public isCollapsed: boolean = false;
  public isCollapsable: boolean = false;
  maxHeight: number = 10;
  Favreports: any = [];

  // notes view / comments
  notesViewState: boolean = true;
  commentsVisibility: boolean = false;
  hideVisibility: boolean = false;

  // scroll
  @ViewChild('scrollcent') scrollcent!: ElementRef;
  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0;

  StoreVal: any = '71,53,8,7,4,35,1,32,40,50,25,18,31,3,70,72,2,17,41,55,42,51,12,73,54,9,15,5,14,30,11';
  StoresIds: any = [];

  Performance: any = 'Load';
  maxDate: any;
  // vehiclear already defined above (string); keep an array for the report controls too:
  vehiclearArray: any = ['WAR'];
  @Input() headerInput: any;
  @Input() popupInput: any;
  Bar: boolean = false;
  storeName: any = '';
  employeeschanges: any = '';
  @Input() QISearchNameInput: any;
  @Output() messageEvent = new EventEmitter<string>();
  selectedstorevalues: any = [];
  AllStores: boolean = true;
  selectedGroups: any = [];
  AllGroups: boolean = true;
  groupstate: boolean = false;

  month: any;
  selectedFiManagersvalues: any = [];
  selectedFiManagersname: any = [];
  financeManager: any = [];
  helpdata: any;
  activePopover: number = -1;
  bsConfig!: Partial<BsDatepickerConfig>;

  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  stores: any = []
  //  storeIds: any = '';
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .dropdown-menu , .timeframe, .reportstores-card');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  constructor(
    public shared: Sharedservice,
    public setdates: Setdates,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private elementRef: ElementRef,
    private renderer: Renderer2,
    private cdRef: ChangeDetectorRef,
    private datepipe: DatePipe,
    public formatter: NgbDateParserFormatter,
    public toast: ToastService,
    private comm: common

  ) {
    // set defaults and mirrored logic from appointment style
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      if (localStorage.getItem('flag') == 'V') {
        this.storeIds = [];
        console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.groupid
        JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store.indexOf(',') > 0 ?
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store.split(',') :
          this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store)
        this.dealStatus = this.dealStatus;
        this.financeManagerId = 0;
        JSON.parse(localStorage.getItem('userInfo')!).WidgetData.flag2 == 'A' ? this.allordebit = 'all' : this.allordebit = 'dr';
        localStorage.setItem('flag', 'M')
      } else {
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
        this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
      }
      this.Role = JSON.parse(localStorage.getItem('userInfo')!).user_Info.title;
      this.userid = JSON.parse(localStorage.getItem('userInfo')!).user_Info.userid;
    }

    this.commentsVisibility = true;



    this.shared.setTitle(this.shared.common.titleName + '-CIT');


    const data = {
      title: 'CIT',
      stores: this.storeIds,
      groups: this.groups,
      financemanagers: this.financeManagerId,
      dealStatus: this.dealStatus,
      saleType: this.saleType,
      allordebit: this.allordebit,
      vehiclear: this.vehiclear,
    };
    this.shared.api.SetHeaderData({ obj: data });

    if (this.storeIds != '') {
      this.Getfloorplansdata();
      this.getEmployees();
    }

  }

  ngOnInit(): void { }

  // Messaging (report -> parent)
  sendMessage() {
    this.messageEvent.emit(this.QISearchName);
  }

  /* ------------------------------
     Core: Get floor plan / CIT data
     ------------------------------ */
  selectedAgingFrom: number | null = 0;
  selectedAgingTo: number | null = 0;
  agingFilter: { min: number; max: number | null } | null = null;
  applyAgingFilter(min: number, max: number | null) {
    this.agingFilter = { min, max };
    this.selectedAgingFrom = min;
    this.selectedAgingTo = max;

    console.log(this.agingFilter)
    this.Getfloorplansdata()

  }
  Getfloorplansdata() {
    this.goToFirstPage();
    this.NoData = false;
    this.FloorPlanData = [];
    this.FloorPlanTotalData = [];
    this.filteredFloorplanData = [];
    try { this.shared.spinner.show(); } catch { /* ignore */ }

    const obj = {
      AS_ID: this.storeIds.toString(),
      DealStatus: this.dealStatus.toString(),
      ValueType: this.allordebit && this.allordebit.toString ? (this.allordebit.toString() == 'all' ? '' : this.allordebit.toString()) : '',
      FIManagerID: this.financeManagerId,
      UserID: 0,
      GetPriority: this.tempPriorityFilter && this.tempPriorityFilter.toString ? (this.tempPriorityFilter.toString() == 'All' ? '' : this.tempPriorityFilter.toString()) : '',
      // Aging_From: this.agingFilter?.min ?? null,
      // Aging_To: this.agingFilter?.max ?? null,
      // Retail_Fleet: this.saleType.toString(),
    };

    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetCITFloorPlanData', obj).subscribe(
      (res: any) => {
        try { this.shared.spinner.hide(); } catch { }
        if (res && res.status == 200) {
          if (res.response && res.response.length > 0) {
            this.FloorPlanData = res.response.filter((e: any) => e.store != 'TOTAL');
            this.FloorPlanTotalData = res.response.filter((e: any) => e.store == 'TOTAL');

            // Parse AgeData / Comments / Notes safely
            this.FloorPlanData.forEach((x: any) => {
              if (x && x.AgeData && typeof x.AgeData === 'string') {
                try { x.AgeData = JSON.parse(x.AgeData); } catch { }
              }
              if (x && x.Comments && typeof x.Comments === 'string') {
                try { x.Comments = JSON.parse(x.Comments); } catch { }
              }
              if (x.Notes) {
                x.Notes = JSON.parse(x.Notes);

                x.allNotes = [...x.Notes];        // ✅ full data
                x.duplicateNotes = x.allNotes.slice(0, 3); // ✅ first 3 only

                x.Notesstate = '+';
              }
            });
            this.originalFloorPlanData = JSON.parse(
              JSON.stringify(this.FloorPlanData)
            );
            this.filteredFloorplanData = this.FloorPlanData || [];
            this.NoData = this.FloorPlanData.length === 0;
          } else {
            this.NoData = true;
          }
        } else {
          this.NoData = true;
        }
      },
      (error: any) => {
        try { this.shared.spinner.hide(); } catch { }
        this.NoData = true;
      }
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
          this.normalize(x.AGE).includes(t) ||
          this.normalize(x.AgeDate).includes(t) ||
          this.normalize(x.Control).includes(t) ||

          // ===== BALANCES =====
          this.normalize(x.BalCIT).includes(t) ||
          this.normalize(x.CITAccount).includes(t) ||
          this.normalize(x.BalCashOSF).includes(t) ||
          this.normalize(x.CashOSFAccount).includes(t) ||

          // ===== CUSTOMER =====
          this.normalize(x.CustomerName).includes(t) ||
          this.normalize(x.CustomerNumber).includes(t) ||
          this.normalize(x.store).includes(t) ||

          // ===== DEAL INFO =====
          this.normalize(x.StockNumner).includes(t) ||
          this.normalize(x.Deal).includes(t) ||
          this.normalize(x.BankName).includes(t) ||
          this.normalize(x.FIManager).includes(t) ||
          this.normalize(x.SalesManager).includes(t) ||
          this.normalize(x.SP1).includes(t) ||
          this.normalize(x.SaleType).includes(t) ||
          this.normalize(x.DealType).includes(t) ||
          this.normalize(x.DealStatus).includes(t) ||

          // ===== VEHICLE =====
          this.normalize(x.VehicleYear).includes(t) ||
          this.normalize(x.VehicleMake).includes(t) ||
          this.normalize(x.VehicleModel).includes(t)

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

  get BalCITTotal(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.BalCIT) || 0);
    }, 0);
  }

  get CITAccount(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.CITAccount) || 0);
    }, 0);
  }
  get BalCashOSF(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.BalCashOSF) || 0);
    }, 0);
  }
  get CashOSFAccount(): number {
    return this.filteredData.reduce((total, item) => {
      return total + (parseFloat(item.CashOSFAccount) || 0);
    }, 0);
  }
  ////////////////////////////////////Pagination Code End////////////////////////////////////////////


  priorityAction: 'Y' | 'N' | null = null;
  selectedPriorityRecord: any = null;
  previousHasPriority: 'Y' | 'N' | null = null;
  priority(e: any, val: any, confirmtemplate: any, ref: any, refval: any) {
    this.previousHasPriority = val.HasPriority;
    console.log('previousHasPriority set to:', this.previousHasPriority);
    this.selectedPriorityRecord = val;
    if (ref == 'multi') {
      if (this.priorityRecords.length == 0) {
        this.toast.show(
          'Please Select Atleast One record to prioritize',
          'warning',
          'Warning'
        );
        const element = <HTMLInputElement>document.getElementById('Priority');
        if (element) element.checked = false;
      } else {
        if (e.target.checked) {
          this.priorityAction = 'Y';
          this.priorityVisibility = true;
          this.ngbmodalActive = this.ngbmodal.open(confirmtemplate, { size: 'sm', backdrop: 'static' });
        }
      }
    } else {

      if (e.target.checked) {
        this.priorityAction = 'Y';
        this.priorityVisibility = true;
        this.priorityRecords.push(val);
      } else {
        this.priorityAction = 'N';

        // Remove from array if exists
        const index = this.priorityRecords.findIndex(
          (list: any) => list.StockNumner === refval
        );
        if (index > -1) {
          this.priorityRecords.splice(index, 1);
        }

        if (!e.target.checked && this.previousHasPriority == 'Y') {
          // Open confirmation modal
          this.ngbmodalActive = this.ngbmodal.open(confirmtemplate, {
            size: 'sm',
            backdrop: 'static'
          });
        }
      }
    }
  }
  confirmPriorityAction() {

    if (this.priorityAction === 'Y') {
      // ✅ Bulk / single PRIORITIZE
      this.priorityAdd();
      return;
    }

    // ❌ UNPRIORITIZE (single record)
    const payload = {
      schedulecontrolpriority: [
        {
          AS_ID: this.selectedPriorityRecord.storeid,
          Account: this.selectedPriorityRecord.Account,
          Scheduletype: 'CIT',
          Control: this.selectedPriorityRecord.Control,
          Stockno: this.selectedPriorityRecord.StockNumner,
          Dealno: this.selectedPriorityRecord.Deal,
          Custno: this.selectedPriorityRecord.CustomerNumber,
          UserID: this.userid,
          SetPriority: 'N'
        }
      ]
    };

    this.shared.api.postmethod('ReceivableExcludeControls/AddScheduleControlPriority', payload).subscribe(
      (res: any) => {
        if (res && res.status === 200) {
          this.toast.show(
            'Priority removed successfully',
            'success',
            'Success'
          );
          (document.getElementById('closeadd') as HTMLInputElement)?.click();
          this.onclose()
          this.Getfloorplansdata();
          this.prioritycheck = false;
        } else {
          this.toast.show(
            'Failed to unprioritize',
            'danger',
            'Error'
          );
        }
      },
      () => {
        this.toast.show(
          '502 Bad Gateway Error',
          'danger',
          'Error'
        );
      }
    );
  }
  priorityAdd() {
    const payload = { schedulecontrolpriority: [] as any[] };

    for (let i = 0; i < this.priorityRecords.length; i++) {
      payload.schedulecontrolpriority.push({
        AS_ID: this.priorityRecords[i].storeid,
        Account: this.priorityRecords[i].Account,
        Scheduletype: "CIT",
        Control: this.priorityRecords[i].Control,
        Stockno: this.priorityRecords[i].StockNumner,
        Dealno: this.priorityRecords[i].Deal,
        Custno: this.priorityRecords[i].CustomerNumber,
        UserID: this.userid,
        SetPriority: "Y"
      });
    }
    if (payload.schedulecontrolpriority.length == 0) {
      this.toast.show(
        'Please Select Atleast One record to prioritize',
        'warning',
        'Warning'
      );
      return;
    }
    this.shared.api.postmethod('ReceivableExcludeControls/AddScheduleControlPriority', payload).subscribe((res: any) => {
      if (res && res.status == 200) {
        this.toast.show(
          'This control prioritized successfully',
          'success',
          'Success'
        );
        (document.getElementById('closeadd') as HTMLInputElement)?.click();
        this.onclose()
        this.Getfloorplansdata();
        this.prioritycheck = false;
        this.priorityRecords = [];
        this.priorityVisibility = false;
      } else {
        this.toast.show(
          res.status,
          'danger',
          'Error'
        );
        try { this.shared.spinner.hide(); } catch { }
        this.NoData = true;
      }
    }, (error: any) => {
      this.toast.show(
        '502 Bad Gateway Error',
        'danger',
        'Error'
      );
      try { this.shared.spinner.hide(); } catch { }
      this.NoData = true;
    });
  }

  currentCheckboxEvent: any = null;
  previousPriorityState: boolean = false;
  previousHeaderCheck: boolean = false;

  onPriorityCancel(modal: any) {
    this.FloorPlanData = JSON.parse(
      JSON.stringify(this.originalFloorPlanData)
    );

    this.prioritycheck = false
    this.priorityAction = null;
    this.selectedPriorityRecord = null;
    this.previousHasPriority = null;


    modal.dismiss();
  }

  // -------------------------------------
  // Merged report component methods (groups/stores/people)
  // -------------------------------------
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'CIT') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    // subscribe to reports and export triggers
    this.reportGetting = this.shared.api.GetReports().subscribe((data: any) => {
      if (data && data.obj && data.obj.Reference == 'CIT') {
        if (data.obj.header == undefined) {
          this.storeIds = data.obj.storeValues;
          this.vehiclear = data.obj.vehiclear;
          this.dealStatus = data.obj.dealStatus;
          this.saleType = data.obj.saleType;
          this.financeManagerId = data.obj.FIvalues;
          this.allordebit = data.obj.allordebit;

          this.groups = data.obj.groups;
          this.enablevehicle = this.vehiclear === 'WAR';
        } else {
          if (data.obj.header == 'Yes') {
            this.storeIds = data.obj.storeValues;
          }
        }

        if (this.storeIds != '') {
          this.Getfloorplansdata();
        } else {
          this.NoData = true;
          this.FloorPlanData = [];
        }

        const headerdata = {
          title: 'CIT',
          stores: this.storeIds,
          groups: this.groups,
          financemanagers: this.financeManagerId,
          dealStatus: this.dealStatus,
          saleType: this.saleType,
          allordebit: this.allordebit,
          vehiclear: this.vehiclear,
        };
        this.shared.api.SetHeaderData({ obj: headerdata });
        this.header = [{
          type: 'Bar',
          storeIds: this.storeIds,
          vehiclear: this.vehiclear,
          dealStatus: this.dealStatus,
          saleType: this.saleType,
          allordebit: this.allordebit,
          groups: this.groups,
          financemanagers: this.financeManagerId,
        }];
      }
    });

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'CIT') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'CIT') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'CIT') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'CIT') return;
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
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
    this.getEmployees()
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
    if (ref == 'ST') {
      const index = this.saleType.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.saleType.splice(index, 1);
      } else {
        this.saleType.push(e);
      }
    }
    if (ref == 'PT') {
      this.tempPriorityFilter = e;
    }
  }


  viewmoreAction(fp: any) {
    if (!fp.Notes) return;
    if (fp.Notesstate === '+') {
      fp.Notesstate = '-';
      fp.duplicateNotes = [...fp.Notes]; // full list (spread = new reference)
    } else {
      fp.Notesstate = '+';
      fp.duplicateNotes = fp.Notes.slice(0, 3); // first 3
    }
    this.cdRef.detectChanges();
  }
  viewDeal(dealData: any) {
    // const modalRef = this.ngbmodal.open(DealRecapComponent, { size: 'md', windowClass: 'connectedmodal' });
    // modalRef.componentInstance.data = { dealno: dealData.Deal, storeid: dealData.storeid, stock: dealData.StockNumner, vin: dealData.vin, custno: dealData?.CustomerNumber }; // Pass data to the modal component    
    // modalRef.result.then((result) => {
    //   console.log(result); // Handle modal close result
    // }, (reason) => {
    //   console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    // });
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

        } else {
          this.toast.show(
            'Invalid Details',
            'danger',
            'Error'
          );
        }
      },
      (error: any) => {
      },
    );
  }
  employees(block: any, e: any, ename?: any) {
    if (block === 'FM') {
      const index = this.selectedFiManagersvalues.findIndex((i: any) => i == e);

      if (index >= 0) {
        this.selectedFiManagersvalues.splice(index, 1);
      } else {
        this.selectedFiManagersvalues.push(e);
      }
      const index1 = this.selectedFiManagersvalues.findIndex((i: any) => i == e);
      if (index1 >= 0) {
        this.selectedFiManagersname = ename;
      }

      return;
    }
    if (block === 'AllFM') {
      if (e === 0) {
        this.selectedFiManagersvalues = this.financeManager.map((fm: any) => fm.FiId);
      } else if (e === 1) {
        this.selectedFiManagersvalues = [];
      }

      return;
    }
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

    this.dealStatus = this.dealStatus;
    this.saleType = this.saleType;
    ((this.groups = this.groups),
      (this.financeManagerId =
        this.selectedFiManagersvalues.length == this.financeManager.length
          ? '0'
          : this.selectedFiManagersvalues.toString()));
    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show(
        'Please Select Atleast One Store',
        'warning',
        'Warning'
      );
      return;
    }

    // ✅ Finance Manager (People) validation
    if (!this.selectedFiManagersvalues || this.selectedFiManagersvalues.length === 0) {
      this.toast.show(
        'Please Select Atleast One People',
        'warning',
        'Warning'
      );
      return;
    }
    this.Getfloorplansdata();
  }

  closeReportModal() {
    try { this.ngbmodal.dismissAll(); } catch { }
    this.Performance = 'Unload';
  }

  // ---------------------------------------------------------
  // Notes / hide logic (already in component)
  // ---------------------------------------------------------
  addNotes(item: any) {
    this.selecteddata = item;
    this.getDropDown(this.selecteddata.companyid);
    this.notesStageText = '';
    this.notesStageValue = '';
  }

  getDropDown(companyid: any) {
    const obj = {
      AssociatedReport: 'CIT',
      CompanyID: companyid
    };
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetScheduleNoteStages', obj).subscribe((res: any) => {
      if (res && res.status == 200) {
        this.notesstage = res.response;
        this.notesStageValue = '';
      }
    });
  }

  save() {
    if (this.notesStageText == '') {

      this.toast.show('Please enter notes', 'warning', 'Warning');
      return;
    }
    const obj = {
      AS_ID: this.selecteddata.storeid,
      Account: this.selecteddata.CIT_Account,
      Control: this.selecteddata.Control,
      Notes: this.notesStageText,
      StageId: this.notesStageValue,
      UserID: JSON.parse(localStorage.getItem('userInfo')!).userid
    };

    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'AddScheduleNotesAction', obj).subscribe((res: any) => {
      if (res && res.status == 200) {
        try { this.shared.spinner.hide(); } catch { }

        this.toast.show('Notes Added Successfully', 'success', 'Success');
        this.callLoadingState = 'ANS';
        document.getElementById('close')?.click();
        this.onclose();
        this.commentsVisibility = true;

        const userName = JSON.parse(localStorage.getItem('userInfo')!)?.fullName;
        const curDate = new Date();
        let stageText = '';
        let nts = '';

        // ✅ Get Stage Text safely
        if (this.notesStageValue) {
          const filteredData = this.notesstage.filter((item: any) => item.NS_ID == this.notesStageValue);
          if (filteredData?.length) {
            stageText = filteredData[0].NS_Text.trim();
            this.selecteddata.Stage = stageText;
          }
        }

        // ✅ Build formatted note
        if (stageText) {
          nts = `${stageText} - ${this.notesStageText} - ${userName} - (${this.formatDate(curDate)})`;
        } else {
          nts = `${this.notesStageText} - ${userName} - (${this.formatDate(curDate)})`;
        }

        // ✅ Create note object
        const noteObj = {
          STAGE: stageText || '',
          NOTES: nts,
          NOTESDATE: this.formatDate(curDate),
          UserName: userName
        };

        // ✅ Update both Notes (full list) and duplicateNotes (display subset)
        if (!this.selecteddata.Notes) {
          this.selecteddata.Notes = [];
        }

        // Add new note at the top of full list
        this.selecteddata.Notes.unshift(noteObj);

        // Refresh the 3-note preview
        this.selecteddata.duplicateNotes = this.selecteddata.Notes.slice(0, 3);
        this.selecteddata.Notesstate = '+';
        this.selecteddata.NotesStatus = 'Y';

        this.cdRef.detectChanges(); // ✅ Force view refresh
      } else {

        this.toast.show('Something went wrong, please try again.', 'danger', 'Error');
      }
    });



  }

  collectHidevalues(e: any, val: any, confirmtemplate: any, ref: any, refval: any) {

    if (ref == 'multi') {
      if (this.hideRecords.length == 0) {

        this.toast.show('Please Select Atleast One Record to Hide', 'warning', 'Warning');
        const element = <HTMLInputElement>document.getElementById('symbol');
        if (element) element.checked = false;
      } else {
        if (e.target.checked) {
          this.hideVisibility = true;
          this.ngbmodalActive = this.ngbmodal.open(confirmtemplate, { size: 'sm', backdrop: 'static' });
        }
      }
    } else {

      if (e.target.checked) {
        this.hideVisibility = true;
        this.hideRecords.push(val);
      } else {
        const index = this.hideRecords.findIndex((list: { StockNumner: any }) => list.StockNumner == refval);
        this.hideRecords.splice(index, 1);
      }
    }
  }

  hideAdd() {

    this.FinalArray = [];
    for (let i = 0; i < this.hideRecords.length; i++) {
      const account = [this.hideRecords[i].CIT_Account, this.hideRecords[i].Vehicle_Account]
        .filter(Boolean)
        .join(', ');
      this.FinalArray.push({
        Receivable_Type: 'CIT',
        Account: account,
        CompanyID: this.hideRecords[i].companyid,
        AS_ID: this.hideRecords[i].storeid,
        Control: this.hideRecords[i].Control,
        Stock: this.hideRecords[i].StockNumner,
        Deal: this.hideRecords[i].Deal,
        Control_Status: 'Y'
      });
    }
    if (this.FinalArray.length == 0) {

      this.toast.show('Please Select Atleast One Record to Hide', 'warning', 'Warning');
      return;
    }
    const obj = { receivableexcludecontrol: this.FinalArray };
    this.shared.api.postmethod('ReceivableExcludeControls', obj).subscribe((res: any) => {
      if (res && res.status == 200) {

        this.toast.show('This control hidden successfully', 'success', 'Success');
        (document.getElementById('closeadd') as HTMLInputElement)?.click();
        this.onclose()
        this.Getfloorplansdata();
        this.hideRecords = [];
        this.hideVisibility = false;
      } else {

        this.toast.show(res.status, 'danger', 'Error');
        try { this.shared.spinner.hide(); } catch { }
        this.NoData = true;
      }
    }, (error: any) => {

      this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
      try { this.shared.spinner.hide(); } catch { }
      this.NoData = true;
    });
  }

  public inTheGreen(value: number): boolean {
    return value >= 0;
  }

  // ---------------------------------------------------------
  // UI helpers / scrolling / sorting / popovers
  // ---------------------------------------------------------
  updateVerticalScroll(event: any): void {
    this.scrollCurrentposition = event.target.scrollTop;
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop / (event.target.scrollHeight - (scrollDemo?.clientHeight || 1))) * 100
    );
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

  sort(property: any, state?: any) {
    if (state == undefined) {
      this.isDesc = !this.isDesc;
    }
    this.callLoadingState = 'FL';
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    this.FloorPlanData.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }

  openMenu(Object: any) { /* placeholder */ }

  notesView() {
    this.notesViewState = !this.notesViewState;
  }

  openComments() {
    this.commentsVisibility = !this.commentsVisibility;
  }
  index = '';
  clean(value: any) {
    if (Array.isArray(value)) {
      return value.join(', ');
    }
    if (typeof value === 'string' && value.startsWith('[')) {
      try {
        return JSON.parse(value).join(', ');
      } catch {
        return value.replace(/[\[\]"]+/g, '');
      }
    }
    return value || '-';
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

    /* ================= TITLE ================= */
    doc.setFontSize(14);
    doc.text('CIT Floorplan', 14, 12);

    const aging = this.FloorPlanData?.[0]?.AgeData?.[0] ?? {};

    /* ================= AGING TABLE ================= */
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
      head: [[
        'CIT Aging', 'TOTAL', '0-5', '6-10', '11-15', '15+'
      ]],
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

    /* ================= HEADERS ================= */
    const head2 = [
      'Age', 'Date', 'Account', 'Control', 'Control 2',
      'CIT', 'Floorplan',
      'Name', 'Number',
      'Stock #', 'Deal #', 'Stage', 'F&I Mgr', 'Sales Mgr', 'Sales Person',
      'Type', 'New/Used', 'Status', 'Bank Name',
      'Year', 'Make', 'Model',
      'Delivery', 'Sale', 'Funding',
      'Trade'
    ];

    const alignments = [
      'left', 'center', 'center', 'center', 'left',
      'right', 'right',
      'left', 'center',
      'left', 'center', 'left', 'left', 'left', 'left',
      'center', 'center', 'center', 'left',
      'center', 'left', 'left',
      'center', 'center', 'center',
      'left'
    ];

    /* ================= DATA ROWS ================= */
    const rows: any[] = [];

    (this.filteredData || []).forEach((fp: any, i: number) => {

      const rowData = [
        fp.AGE || '-',
        this.formatDateMMDDYYYY(fp.AgeDate),
        fp.Account || '-',
        fp.Control || '-',
        fp.Control2 || '-',

        fp.BalCIT || 0,
        fp.BalFloorplan || 0,

        fp.CustomerName || '-',
        fp.CustomerNumber || '-',

        fp.StockNumner || '-',
        fp.Deal || '-',
        fp.Stage || '-',
        fp.FIManager || '-',
        fp.SalesManager || '-',
        fp.SP1 || '-',

        fp.SaleType || '-',
        fp.DealType || '-',
        fp.DealStatus || '-',
        fp.BankName || '-',

        fp.VehicleYear || '-',
        fp.VehicleMake || '-',
        fp.VehicleModel || '-',

        this.formatDateMMDDYYYY(fp.DeliveryDate),
        this.formatDateMMDDYYYY(fp.DateSale),
        this.formatDateMMDDYYYY(fp.FundingDate),

        fp.trade1modelname || '-'
      ];

      rows.push(rowData);

      /* ================= NOTES ================= */
      if (fp.Notes && fp.Notes.length > 0 && this.commentsVisibility) {
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

        let notes = this.expandedNotes[i] || fp.Notes;
        if (!Array.isArray(notes)) notes = [notes];

        notes.forEach((note: any) => {
          rows.push([{
            content: '   ' + (note?.NOTES ?? 'No additional notes provided.'),
            colSpan,
            styles: {
              fillColor: [243, 246, 255],
              halign: 'left',
              // fontStyle: 'italic',
              fontSize: 8
            }
          }, ...Array(colSpan - 1).fill({})]);
        });
      }
    });

    /* ================= GROUP HEADER ================= */
    const borderStyle = {
      lineWidth: { left: .3, right: .3, top: 0, bottom: 0 },
      lineColor: [180, 180, 180]
    };

    const head1 = [
      { content: '', colSpan: 5, styles: { fillColor: [5, 84, 239], ...borderStyle } },
      { content: 'Balances', colSpan: 2, styles: { fillColor: [5, 84, 239], textColor: 255, halign: 'center', ...borderStyle } },
      { content: 'Customer', colSpan: 2, styles: { fillColor: [5, 84, 239], textColor: 255, halign: 'center', ...borderStyle } },
      { content: 'Deal Info', colSpan: 10, styles: { fillColor: [5, 84, 239], textColor: 255, halign: 'center', ...borderStyle } },
      { content: 'Vehicle', colSpan: 3, styles: { fillColor: [5, 84, 239], textColor: 255, halign: 'center', ...borderStyle } },
      { content: 'Dates', colSpan: 3, styles: { fillColor: [5, 84, 239], textColor: 255, halign: 'center', ...borderStyle } },
      { content: '', colSpan: 1, styles: { fillColor: [5, 84, 239], ...borderStyle } }
    ];

    /* ================= COLUMN WIDTH SCALING ================= */
    const baseWidths = [
      10, 15, 30, 20, 20,
      18, 18,
      30, 20,
      18, 18, 18, 22, 22, 22,
      18, 18, 18, 28,
      14, 18, 22,
      18, 18, 18,
      25
    ];

    const pageW = doc.internal.pageSize.getWidth();
    const printableW = pageW - 12;

    const scale = printableW / baseWidths.reduce((a, b) => a + b, 0);

    const columnStyles = baseWidths.reduce((acc: any, w, i) => {
      acc[i] = { cellWidth: Math.max(8, w * scale) };
      return acc;
    }, {});

    /* ================= MAIN TABLE ================= */
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

          // Currency formatting
          if ([5, 6].includes(data.column.index)) {
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
      doc.save(`CIT.pdf`);
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
        this.shared.spinner.hide();
        return;
      }
      const pdfFile = this.blobToFile(pdfBlob, 'CIT.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'CIT');
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
    const worksheet = workbook.addWorksheet('CIT');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 15,
        topLeftCell: 'A16',
        showGridLines: false,
      },
    ];

    worksheet.addRow(['CIT']).font = { size: 14, bold: true };
    worksheet.addRow([]);

    worksheet.addRow([new Date().toLocaleString()]);
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

    const agingHeader = ['CIT Aging', 'TOTAL', '0-5', '6-10', '11-15', '15+'];
    const agingRow = [
      '',
      this.formatDollar(aging?.TOTAL),
      this.formatDollar(aging?.D1),
      this.formatDollar(aging?.D2),
      this.formatDollar(aging?.D3),
      this.formatDollar(aging?.D4)
    ];
    const agHeaderRow = worksheet.addRow(agingHeader);
    agHeaderRow.eachCell(cell => {
      cell.font = { bold: true };
      cell.alignment = { horizontal: 'center' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0554EF' } };
      cell.font = { bold: true, color: { argb: 'FFFFFF' } };
      cell.border = {
        left: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' }
      };
    });
    const agDataRow = worksheet.addRow(agingRow);
    agDataRow.eachCell(cell => {
      cell.alignment = { horizontal: 'center' };
      cell.border = {
        left: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' }
      };
    });
    worksheet.addRow([]);

    /* ================= HEADER GROUP (FIXED) ================= */

    const groupHeaders = [
      { label: '', span: 5 },
      { label: 'Balances', span: 2 },
      { label: 'Customer', span: 2 },
      { label: 'Store', span: 1 },
      { label: 'Deal Info', span: 10 },
      { label: 'Vehicle', span: 3 },
      { label: 'Dates', span: 3 },
      { label: 'Trade', span: 1 }
    ];

    let col = 1;
    const groupRow = worksheet.addRow([]);

    groupHeaders.forEach(g => {
      const cell = groupRow.getCell(col);
      cell.value = g.label;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { bold: true, color: { argb: 'FFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0554EF' } };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
      if (g.span > 1) {
        worksheet.mergeCells(groupRow.number, col, groupRow.number, col + g.span - 1);
      }

      col += g.span;
    });

    /* ================= COLUMN HEADERS ================= */

    const headers = [
      'Age', 'Date', 'Account', 'Control', 'Control 2',
      'CIT', 'Floorplan',
      'Name', 'Number',
      'Store',
      'Stock #', 'Deal #', 'Stage', 'F&I Mgr', 'Sales Mgr', 'Sales Person',
      'Type', 'New/Used', 'Status', 'Bank Name',
      'Year', 'Make', 'Model',
      'Delivery', 'Sale', 'Funding',
      'Trade'
    ];

    const headerRow = worksheet.addRow(headers);

    headerRow.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9E7FF' } };
      cell.font = { bold: true };
      cell.alignment = { horizontal: 'center' };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    /* ================= ALIGNMENTS ================= */

    const alignments = [
      'center', 'center', 'center', 'center', 'left',
      'right', 'right',
      'left', 'center',
      'left',
      'left', 'center', 'left', 'left', 'left', 'left',
      'center', 'center', 'center', 'left',
      'center', 'left', 'left',
      'center', 'center', 'center',
      'left'
    ];

    /* ================= DATA ================= */

    (this.filteredData || []).forEach((fp: any) => {

      const rowData = [
        fp.AGE ?? '-',
        this.formatDateMMDDYYYY(fp.AgeDate),
        fp.Account ?? '-',
        fp.Control ?? '-',
        fp.Control2 ?? '-',
        fp.BalCIT ?? 0,
        fp.BalFloorplan ?? 0,
        fp.CustomerName ?? '-',
        fp.CustomerNumber ?? '-',
        fp.store ?? '-',
        fp.StockNumner ?? '-',
        fp.Deal ?? '-',
        fp.Stage ?? '-',
        fp.FIManager ?? '-',
        fp.SalesManager ?? '-',
        fp.SP1 ?? '-',
        fp.SaleType ?? '-',
        fp.DealType ?? '-',
        fp.DealStatus ?? '-',
        fp.BankName ?? '-',
        fp.VehicleYear ?? '-',
        fp.VehicleMake ?? '-',
        fp.VehicleModel ?? '-',
        this.formatDateMMDDYYYY(fp.DeliveryDate),
        this.formatDateMMDDYYYY(fp.DateSale),
        this.formatDateMMDDYYYY(fp.FundingDate),
        `${fp.trade1year || ''} ${fp.trade1modelname || '-'}`
      ];

      const row = worksheet.addRow(rowData);

      row.eachCell((cell: any, i) => {

        cell.alignment = {
          horizontal: alignments[i - 1] as any,
          vertical: 'middle'
        };

        // Currency columns
        if (i === 6 || i === 7) {
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
    const widths = [
      15, 15, 15, 15, 15,
      14, 14,
      25, 13,
      20,
      15, 15, 15, 18, 18, 18,
      12, 12, 14, 30,
      12, 15, 16,
      12, 12, 12,
      20
    ];

    widths.forEach((w, i) => worksheet.getColumn(i + 1).width = w);
    workbook.xlsx.writeBuffer().then(buffer => {
      const blob = new Blob([buffer]);
      FileSaver.saveAs(blob, 'CIT_Report.xlsx');
    });
  }

  isNaN(value: any): boolean {
    return value !== null && value !== undefined && !isNaN(Number(value));
  }


  formatDate(date: number | Date | undefined) {
    const options: Intl.DateTimeFormatOptions = {
      month: '2-digit' as '2-digit',
      day: '2-digit' as '2-digit',
      year: 'numeric' as 'numeric',
    };
    return new Intl.DateTimeFormat('en-US', options).format(date as Date);
  }
}
