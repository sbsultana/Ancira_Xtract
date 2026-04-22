import {
  Component,
  OnInit,
  ViewChild,
  TemplateRef,
  ChangeDetectorRef,
  HostListener,
  Directive,
} from '@angular/core';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { DatePipe, CommonModule } from '@angular/common';
import { NgbActiveModal, NgbModal, NgbModalRef } from '@ng-bootstrap/ng-bootstrap';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { common } from '../../../../common';

import { HttpClient } from '@angular/common/http';
import * as ExcelJS from 'exceljs';
import * as FileSaver from 'file-saver';
import { Subscription } from 'rxjs';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
declare var bootstrap: any;

@Component({
  selector: 'app-dashboard',
  imports: [CommonModule, SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard implements OnInit {
  modalInstance: any;
  @ViewChild('detailModal') detailModal!: TemplateRef<any>;
  item: any;
  activePopover: number = -1;
  @HostListener('document:click', ['$event'])
  clickOutside(event: MouseEvent) {
    const target = event.target as HTMLElement;

    if (!target.closest('.filter-dropdowns')) {
      this.activePopover = -1;
    }
  }

  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = 0;
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };


  DateType: any;
  maxToMonth!: Date;
  minToMonth!: Date;


  storeName = '';
  storesList: any[] = [];


  allStoresData: any[] = [];
  selectedVars: any = { var1: 'All' };
  isPopoverOpen = false;
  storeButtonLabel: string = 'All';
  FromDate!: string;
  ToDate!: string;
  tomonth: Date | null = null;
  previous: Date = new Date();
  previousmultiplemax: Date = new Date();
  toprevious: any;
  fromDate: any;
  toDate: any;
  filteredMonths: string[] = [];

  // NEW / integrated fields (Client A functionality)
  // table data (original)
  searchTerm: string = '';
  isMapped: boolean | null = null;
  filteredTableData: any[] = [];
  tableData: any[] = [];
  selectedCategory: string = 'Sale';
  apiUrl = 'https://tcxtractapi.axelone.app/api/Ancira/GetGLAccountsInfo';

  // Account mapping specific
  AccountingMapping: any[] = []; // used in integrated UI
  Pagination: boolean = false;
  PageCount: number = 1;
  LastCount: boolean = false;
  NoData: boolean = false;
  NoDatainitail: boolean = false;

  // bulk edit / modal
  bulkEditRecords: any[] = [];
  Bulk: boolean = false;
  ignorebtn: boolean = false;

  // form fields (ngModel bind)
  actbal: string = '';
  accounttype: any = 0;
  selectedaccounttypedetail: any = 0;
  departmenttype: any = 0;
  department: any = 0;
  subtype: any = 0;
  subtypedetail: any = 0;
  subtypefulldetail: any = 0;
  flexcol1: any = 0;
  financialsummary: any = 0;
  submitted: boolean = false;

  Refchange: string = '';
  defaultresponse: any = {};
  BIref: any = '';

  Accounttype: any[] = [];
  Accounttypedetail: any[] = [];
  Departmenttype: any[] = [];
  Department: any[] = [];
  Subtype: any[] = [];
  SubtypeDetail: any[] = [];
  SubtypeFullDetail: any[] = [];
  Flexcol1: any[] = [];
  Financialsummary: any[] = [];
  Balancesummary: any[] = [];

  missingFields: string[] = [];

  // modal reference if using NgbModal programmatically
  private modalRef: NgbModalRef | null = null;

  // ---------------- Header Tabs (Client A Logic) ----------------
  selectedheadertab: any[] = ['S']; // default "Sale" selected

  // ---------------- Mapped / Unmapped & Search (Client A) ----------------
  searchText: string = '';
  statusFilter: any[] = ['M']; // Default shows both
  @ViewChild('Alert') Alert!: TemplateRef<any>;

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  constructor(
    public shared: Sharedservice,
    public setdates: Setdates,
    private comm: common,
    private cdr: ChangeDetectorRef,
    private datepipe: DatePipe,
    public modalService: NgbModal,
    private http: HttpClient,
    private ngbmodalActive: NgbActiveModal,
    private toast: ToastService,
  ) {
    this.shared.setTitle(this.shared.common.titleName + '-AccountMapping');
    if (typeof window !== 'undefined') {

      this.shared.setTitle(this.shared.common.titleName + '-AccountMapping');
      if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
        this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
      }
      if (this.shared.common.groupsandstores.length > 0) {
        this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
        this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
        this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
        this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
        this.getStoresandGroupsValues()
      }

      if (localStorage.getItem('Fav') != 'Y') {
        const data = {
          title: 'Account Mapping',
          stores: this.storeIds,
          datetype: 'MTD',
          fromdate: this.FromDate,
          todate: this.ToDate,
        };
        this.shared.api.SetHeaderData({
          obj: data,
        });
      }
    }
  }

  getMonthYear(): string {
    if (!this.frommonth) return '';

    const dateObj = this.frommonth instanceof Date ? this.frommonth : new Date(this.frommonth);
    const month = (dateObj.getMonth() + 1).toString().padStart(2, '0');
    const year = dateObj.getFullYear();

    return `${month}-${year}`;
  }

  togglePopover(popoverId: number) {
    this.activePopover = this.activePopover === popoverId ? -1 : popoverId;
  }
  ngOnInit(): void {
    this.shared.setTitle('Account Mapping');

    const today = new Date();
    const fromMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    const toMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);

    this.frommonth = fromMonth;
    this.tomonth = toMonth;

    this.minToMonth = this.frommonth;
    this.maxToMonth = today;
    this.getGLAccountsInfo();
  }
  onApply(): void {
    this.getGLAccountsInfo(this.searchTerm);
  }


  headerselection(type: string) {
    this.selectedheadertab = [type];

    // Change category in your existing structure
    switch (type) {
      case 'S':
        this.selectedCategory = 'Sale';
        break;
      case 'C':
        this.selectedCategory = 'Cost';
        break;
      case 'X':
        this.selectedCategory = 'Non-Operating Expense';
        break;
      case 'I':
        this.selectedCategory = 'Non-Operating Income';
        break;
      case 'E':
        this.selectedCategory = 'Expense';
        break;
      case 'A':
        this.selectedCategory = 'Asset';
        break;
      case 'L':
        this.selectedCategory = 'Liability';
        break;
      case 'Q':
        this.selectedCategory = 'Equity';
        break;
    }

    // Refresh data using your original API
    this.getGLAccountsInfo();
  }


  frommonth: Date | null = null;
  Dupfrommonth: Date | null = null;

  minFromMonth: Date = new Date(2000, 0, 1);
  maxFromMonth: Date = new Date();

  monthPickerConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM YYYY',
    minMode: 'month' as 'month',
    showWeekNumbers: false,
    minDate: this.minFromMonth,
    maxDate: this.maxFromMonth,
  };
  changeDatefrom(e: Date | null) {
    this.activePopover = -1;
    if (!e) return;

    this.frommonth = e;

    const year = e.getFullYear();
    const month = e.getMonth();
    this.fromDate = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    this.toprevious = e;
    this.minToMonth = e;
    if (this.tomonth && this.tomonth < e) {
      this.tomonth = lastDay;
    }
  }

  // filterMapped(mapped: boolean) {
  //   this.isMapped = mapped;
  //   this.applyFilters();
  // }

  getGLAccountsInfo(accountNumber: string = '') {
    this.shared.spinner.show();
    this.AccountingMapping = [];
    this.NoData = false;
    const body = {
      map_type: this.statusFilter[0],
      store_Id: this.storeIds.length ? this.storeIds.toString() : '0',
      account_Type: this.mapCategoryToType(this.selectedCategory),
      accountnumber: accountNumber,
      UserID: 1,
      Date: this.frommonth ? this.formatDateForApi(this.frommonth) : '10-12-2025',
    };

    console.log('API Request Body:', body);

    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetGLAccountsInfo', body)
      .subscribe({
        next: (response: any) => {
          if (response?.status === 200) {
            const rows = response?.response || [];

            // If API returned rows
            if (rows.length > 0) {
              console.log('API Response:', rows);

              this.tableData = rows; // raw data
              this.filteredTableData = [...rows]; // filtered copy
              this.AccountingMapping = [...rows]; // grid binding
              this.shared.spinner.hide();

              this.NoData = false;
            } else {
              // Empty response
              this.NoData = true;
              this.tableData = [];
              this.filteredTableData = [];
              this.AccountingMapping = [];
              this.shared.spinner.hide();
            }
          } else {
            // Invalid response
            this.NoData = true;
            this.tableData = [];
            this.filteredTableData = [];
            this.AccountingMapping = [];
          }

          // Pagination logic after filters
          this.Pagination = this.AccountingMapping.length > 10;

          this.cdr.detectChanges();
          //  this.shared.spinner.hide();
        },

        error: (err: any) => {
          console.error('API Error:', err);
          this.shared.spinner.hide();

          this.NoData = true;
          this.tableData = [];
          this.filteredTableData = [];
          this.AccountingMapping = [];
        },
      });
  }

  searchFilter() {
    this.searchTerm = this.searchText.trim(); // sync with backend usage
    this.getGLAccountsInfo(this.searchTerm);
  }

  onSearchInput() {
    // if search field becomes empty, reload original API
    if (!this.searchText.trim()) {
      this.searchTerm = '';
      this.getGLAccountsInfo('');
    }
  }

  StatusFiltering(type: string) {
    this.statusFilter = [type];

    this.getGLAccountsInfo(this.searchTerm);
  }

  applyFilters() {
    this.activePopover = -1;
    this.isPopoverOpen = false;

    let data = this.tableData;

    // apply date filter if needed
    if (this.frommonth) {
      const year = this.frommonth.getFullYear();
      const month = this.frommonth.getMonth();
      this.fromDate = new Date(year, month, 1);
    }

    if (this.tomonth) {
      const year = this.tomonth.getFullYear();
      const month = this.tomonth.getMonth();
      this.toDate = new Date(year, month + 1, 0);
    }

    this.filteredTableData = data;
    // update AccountingMapping used in template
    this.AccountingMapping = [...this.filteredTableData];
    this.Pagination = this.AccountingMapping.length > 10;
    this.NoData = this.AccountingMapping.length === 0;

    this.cdr.detectChanges();
  }

  formatDateForApi(date: Date): string {
    const d = new Date(date);
    const month = ('0' + (d.getMonth() + 1)).slice(-2);
    const day = ('0' + d.getDate()).slice(-2);
    const year = d.getFullYear();
    return `${month}-${day}-${year}`;
  }

  mapCategoryToType(category: string): string {
    switch (category) {
      case 'Sale':
        return 'S';
      case 'Cost':
        return 'C';
      case 'Non-Operating Expense':
        return 'X';
      case 'Non-Operating Income':
        return 'I';
      case 'Expense':
        return 'E';
      case 'Asset':
        return 'A';
      case 'Liability':
        return 'L';
      case 'Equity':
        return 'Q';
      default:
        return 'S';
    }
  }

  setCategory(category: string) {
    this.selectedCategory = category;
    this.getGLAccountsInfo();
  }

  // -------------- CLIENT A MAPPING FUNCTIONS -----------------

  /**
   * bulkEdit(BulkFlag, mode, p3, p4, am, actionType)
   * This mirrors Client A function signature and behavior.
   * - BulkFlag: whether multiple records are being edited
   * - am: record or array of records
   */
  bulkEdit(e: any, ref: any, template: any, multiorsingle: any, record: any, BI?: any) {
    this.BIref = BI;

    if (BI === 'MappedEditMapping') {
      this.Refchange = 'editmapping';
      this.bulkEditRecords = [];
      this.submitted = false;

      // Identify Single or Multiple Selection
      if (ref === 'E') {
        this.Bulk = true;
        this.ignorebtn = true;
      } else {
        if (multiorsingle === 'S') {
          // Insert selected record
          this.bulkEditRecords.push(record);

          // Load all fields + dropdown values just like Client A
          this.OnChange('Default', '', '', '', '', '', '', '', '', '', '');

          // Reset all local dropdown selections
          if (this.bulkEditRecords.length > 0) {
            this.accounttype = '0';
            this.selectedaccounttypedetail = '0';
            this.departmenttype = '0';
            this.department = '0';
            this.subtype = '0';
            this.subtypedetail = '0';
            this.subtypefulldetail = '0';
            this.flexcol1 = '0';
            this.financialsummary = '0';

            // Empty all dependent dropdown lists
            this.Accounttypedetail = [];
            this.Departmenttype = [];
            this.Department = [];
            this.Subtype = [];
            this.SubtypeDetail = [];
            this.SubtypeFullDetail = [];
            this.Flexcol1 = [];
            this.Financialsummary = [];

            // Flags reset just like Client A
            // this.Asubtype = false;
            // this.Asubtypedetail = false;
            // this.Asubtypefulldetail = false;
          }
        }
      }
    }
  }

  OnChange(
    ref: string,
    storeId?: any,
    accountNo?: any,
    accountTypeRaw?: any,
    accountTypeDetail?: any,
    opArea?: any,
    dept?: any,
    subtype?: any,
    subtypeDetail?: any,
    subtypeFullDetail?: any,
    flexcol1?: any,
    refChange?: any,
    actbal?: any
  ) {
    if (refChange === 'AE') {
      this.Refchange = '';
      this.resetAllSelections();
    }

    const body = {
      FilterType: ref === 'AT' ? this.actbal : ref,
      store_ID: this.bulkEditRecords[0].companyId,
      accountnumber: this.bulkEditRecords[0].accountNumber,
      ACCOUNTTYPERAW: accountTypeRaw,
      ACCOUNTTYPEDETAIL: accountTypeDetail,
      OperationalArea: opArea,
      DEPARTMENT: dept,
      SUBTYPE: subtype,
      SUBTYPEDETAIL: subtypeDetail,
      SUBTYPEFULLDETAIL: subtypeFullDetail,
      FLEXCOL1: flexcol1,
    };

    console.log(accountTypeRaw, 'Account Type Raw');
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetAccountMappingFiltersV2', body)
      .subscribe((res: any) => {
        if (res.status !== 200) return;

        const list = res.response;

        /* -----------------------
         DEFAULT (initial load)
      ------------------------*/
        if (ref === 'Default') {
          this.resetAllSelections();
          this.defaultresponse = res;

          const first = list.find((e: any) => e.OperationalType);
          this.actbal = first?.OperationalType || 'Activity';

          this.OnChange('AT');
          return;
        }

        /* -----------------------
         AT — Account Type
      ------------------------*/
        if (ref === 'AT') {
          this.resetLowerLevels('AT');
          this.Accounttype = list.filter((e: { accountType: any }) => e.accountType);

          if (this.Refchange === 'editmapping') {
            this.accounttype = this.bulkEditRecords[0].accountType;
          } else {
            // auto-select if only one
            if (this.Accounttype.length === 1) {
              this.accounttype = this.Accounttype[0].accountType;
            } else {
              this.accounttype = '0';
            }
          }

          // trigger next level
          if (this.actbal === 'Activity') this.OnChange('OA', '', '', this.accounttype);
          else this.OnChange('ATD', '', '', this.accounttype);

          return;
        }

        //        // ALWAYS pass accountType
        // if (this.actbal === 'Activity'){
        //   this.OnChange('OA', '', '', this.accounttype);
        // } else {
        //   this.OnChange('ATD', '', '', this.accounttype);
        //    return;
        // }



        /* -----------------------
         ATD — Account Type Detail
      ------------------------*/
        if (ref === 'ATD') {
          this.resetLowerLevels('ATD');
          this.Accounttypedetail = list.filter(
            (e: { AccountTypeDetail: any }) => e.AccountTypeDetail
          );

          if (this.Refchange === 'editmapping') {
            this.selectedaccounttypedetail = this.bulkEditRecords[0].acctSubtype;
          } else {
            this.selectedaccounttypedetail =
              this.Accounttypedetail.length === 1
                ? this.Accounttypedetail[0].AccountTypeDetail
                : '0';
          }

          this.OnChange('OA', '', '', this.accounttype);
          return;
        }

        /* -----------------------
         OA — Operational Area
      ------------------------*/
        if (ref === 'OA') {
          console.log(this.accounttype, 'Account Type');

          this.resetLowerLevels('OA');
          this.Departmenttype = list.filter((e: { OperationalArea: any }) => e.OperationalArea);

          if (this.Refchange === 'editmapping') {
            this.departmenttype = this.bulkEditRecords[0].operationalArea;
          } else {
            this.departmenttype =
              this.Departmenttype.length === 1 ? this.Departmenttype[0].OperationalArea : '0';
          }

          this.OnChange('DT', '', '', this.accounttype, '', this.departmenttype);
          return;
        }

        /* -----------------------
         DT — Department
      ------------------------*/
        if (ref === 'DT') {
          this.resetLowerLevels('DT');
          this.Department = list.filter((e: any) => e.Department);

          if (this.Refchange === 'editmapping') {
            this.department = this.bulkEditRecords[0].department;
          } else {
            this.department = this.Department.length === 1 ? this.Department[0].Department : '0';
          }

          this.OnChange('ST', '', '', this.accounttype, '', '', this.department);

          return;
        }

        /* -----------------------
         ST — SubType
      ------------------------*/
        if (ref === 'ST') {
          this.resetLowerLevels('ST');
          this.Subtype = list.filter((e: { SubType: any }) => e.SubType);

          if (this.Refchange === 'editmapping') {
            this.subtype = this.bulkEditRecords[0].subType;
          } else {
            this.subtype = this.Subtype.length === 1 ? this.Subtype[0].SubType : '0';
          }

          this.finsummary();
          this.OnChange('STD', '', '', this.accounttype, '', '', this.department, this.subtype);
          return;
        }

        /* -----------------------
         STD — SubType Detail
      ------------------------*/
        if (ref === 'STD') {
          this.resetLowerLevels('STD');
          this.SubtypeDetail = list.filter((e: { SubTypeDetail: any }) => e.SubTypeDetail);

          if (this.Refchange === 'editmapping') {
            this.subtypedetail = this.bulkEditRecords[0].subTypeDetail;
          } else {
            this.subtypedetail =
              this.SubtypeDetail.length === 1 ? this.SubtypeDetail[0].SubTypeDetail : '0';
          }

          this.OnChange(
            'STFD',
            '',
            '',
            this.accounttype,
            '',
            '',
            this.department,
            this.subtype,
            this.subtypedetail
          );
          return;
        }

        /* -----------------------
         STFD — SubType Full Detail
      ------------------------*/
        if (ref === 'STFD') {
          this.SubtypeFullDetail = list.filter((e: any) => e.SubTypeFullDetail);
          return;
        }

        /* -----------------------
         FS — Financial Summary
      ------------------------*/
        if (ref === 'FS') {
          this.Financialsummary = list.filter((e: any) => e.Fin_Summary);

        }
        else {
          this.financialsummary =
            this.Financialsummary.length === 1
              ? this.Financialsummary[0].Fin_Summary_Key
              : '0';
        }

        /* -----------------------
         BS — Balance Summary
      ------------------------*/
        if (ref === 'BS') {
          this.Balancesummary = list.filter((e: any) => e.Balance_Subtype);
          return;
        }
      });
  }

  finsummary() {
    const obj = {
      Accounttype: this.accounttype,
      Department: this.department,
    };
    // this.apiSrvc
    //   .postmethod

    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetAccountFinSummary', obj)
      .subscribe((res: any) => {
        if (!res) return;
        // guard and normalize response before accessing properties
        const resp: any = res;
        if (resp.status == 200) {
          const responseArray = resp.response || [];
          this.Financialsummary = responseArray.filter(
            (e: any) =>
              e.Fin_Summary_Key != '' && e.Fin_Summary_Key != null && e.Fin_Summary_Key != 0
          );

          console.log(this.Financialsummary);
          if (this.Financialsummary.length == 1) {
            if (this.Refchange == 'editmapping') {
              this.financialsummary = this.bulkEditRecords[0].Fin_Summary;
            } else {
              this.financialsummary = this.Financialsummary[0].Fin_summary;
            }
          } else {
            if (this.Refchange == 'editmapping') {
              this.financialsummary = this.bulkEditRecords[0].Fin_Summary;
            } else {
              this.financialsummary = '0';
            }
          }
        }
      });
  }

  resetAllSelections() {
    this.accounttype = '0';
    this.selectedaccounttypedetail = '0';
    this.departmenttype = '0';
    this.department = '0';
    this.subtype = '0';
    this.subtypedetail = '0';
    this.subtypefulldetail = '0';
    this.flexcol1 = '0';
  }

  resetLowerLevels(level: string) {
    if (level === 'AT') {
      this.selectedaccounttypedetail = '0';
      this.departmenttype = '0';
      this.department = '0';
      this.subtype = '0';
      this.subtypedetail = '0';
      this.subtypefulldetail = '0';
      this.flexcol1 = '0';
      return;
    }

    if (level === 'ATD') {
      this.departmenttype = '0';
      this.department = '0';
      this.subtype = '0';
      this.subtypedetail = '0';
      this.subtypefulldetail = '0';
      this.flexcol1 = '0';
      return;
    }

    if (level === 'OA') {
      this.department = '0';
      this.subtype = '0';
      this.subtypedetail = '0';
      this.subtypefulldetail = '0';
      this.flexcol1 = '0';
      return;
    }

    if (level === 'DT') {
      this.subtype = '0';
      this.subtypedetail = '0';
      this.subtypefulldetail = '0';
      this.flexcol1 = '0';
      return;
    }

    if (level === 'ST') {
      this.subtypedetail = '0';
      this.subtypefulldetail = '0';
      this.flexcol1 = '0';
      return;
    }

    if (level === 'STD') {
      this.subtypefulldetail = '0';
      this.flexcol1 = '0';
      return;
    }

    if (level === 'STFD') {
      this.flexcol1 = '0';
      return;
    }
  }

  closeBootstrapModal() {
    const modalEl = document.getElementById('mapModal');
    if (!modalEl) return;

    const modal = bootstrap.Modal.getInstance(modalEl)
      || new bootstrap.Modal(modalEl);

    modal.hide();

    // clean backdrop & body lock
    document.querySelectorAll('.modal-backdrop').forEach(el => el.remove());
    document.body.classList.remove('modal-open');
    document.body.style.overflow = '';
  }

  save(): void {
    this.submitted = true;

    this.missingFields = this.getMissingFields();
    if (this.missingFields.length > 0) {
      this.modalService.open(this.Alert, { centered: true });
      return;
    }

    const payload = {
      companyid: this.bulkEditRecords[0].companyId,
      accountnumber: this.bulkEditRecords[0].accountNumber,
      accountdescription: this.bulkEditRecords[0].accountDescription,
      accounttype: this.accounttype,
      accountTypeDetail: this.actbal == 'Activity' ? '' : this.selectedaccounttypedetail,
      operationalArea: this.departmenttype,
      department: this.department,
      subType: this.subtype,
      subTypeDetail: this.subtypedetail,
      subTypeFullDetail: '',
      OperationalType: this.actbal,
      flexCol1: '',
      flexCol2: '',
      finSummary: this.actbal == 'Balance' ? '' : this.financialsummary,
    };

    this.shared.spinner.show();

    this.shared.api.postmethod('accountmapping/AccountMappingActionV2', payload).subscribe({
      next: (res: any) => {
        this.shared.spinner.hide();

        if (res && (res.status == 200)) {

          this.toast.show('This Account Mapped successfully', 'success', 'Success');

          this.getGLAccountsInfo();

          (<HTMLInputElement>document.getElementById('forceCloseModal')).click();

          this.onclose();
        } else {

          this.toast.show(res?.message || 'Failed to save mapping', 'danger', 'Error');
        }
      },
      error: (err) => {
        this.shared.spinner.hide();
        console.error('UpdateAccountMapping error', err);

        this.toast.show('Error saving mapping.', 'danger', 'Error');
      },
    });
  }

  MapAccountCancel(): void {
    this.submitted = false;
    this.onclose();
  }

  onclose(): void {
    // Reset modal fields and state
    this.bulkEditRecords = [];
    this.actbal = '';
    this.accounttype = 0;
    this.selectedaccounttypedetail = 0;
    this.departmenttype = 0;
    this.department = 0;
    this.subtype = 0;
    this.subtypedetail = 0;
    this.financialsummary = 0;
    this.submitted = false;

    // programmatically dismiss modal if using NgbModal
    try {
      this.modalService.dismissAll();
    } catch (err) {
      // ignore if not using NgbModal programmatic control
    }
  }

  // ---------- Pagination helpers ----------
  Prev(): void {
    if (this.PageCount > 1) {
      this.PageCount--;
      this.getGLAccountsInfo();
    }
  }

  Next(): void {
    if (this.LastCount) {
      this.PageCount++;
      this.getGLAccountsInfo();
    }
  }

  getMissingFields(): string[] {
    const fields = [];

    if (!this.actbal || this.actbal === '0') fields.push('Account Balance');
    if (!this.accounttype || this.accounttype === '0') fields.push('Account Type');
    if (!this.departmenttype || this.departmenttype === '0') fields.push('Operational Area');
    if (!this.department || this.department === '0') fields.push('Department');
    if (!this.subtype || this.subtype === '0') fields.push('Sub Type');
    if (!this.subtypedetail || this.subtypedetail === '0') fields.push('Sub Type Detail');

    return fields;
  }
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit(): void {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Account Mapping') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Account Mapping') return;
      if (obj.state) {
        this.exportToExcelAccountingMapping();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Account Mapping') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Account Mapping') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Account Mapping') return;
      if (obj.stateEmailPdf) {
        this.sendEmailData(obj.Email, obj.notes, obj.from);
      }
    });
    const modalEl = document.getElementById('mapModal');
    if (modalEl) {
      this.modalInstance = new bootstrap.Modal(modalEl);
    }
  }
  ngOnDestroy(): void {
    this.excel?.unsubscribe();
    this.print?.unsubscribe();
    this.pdf?.unsubscribe();
    this.email?.unsubscribe();
  }
  getReportFilters(): { title: string; filters: any[] } {
    return {
      title: 'Account Mapping',
      filters: [
        {
          label: 'Type',
          value:
            this.statusFilter[0] === 'M'
              ? 'Mapped'
              : this.statusFilter[0] === 'U'
                ? 'Unmapped'
                : 'All'
        },
        {
          label: 'Account Type',
          value: this.selectedCategory || 'All'
        },
        {
          label: 'Month',
          value: this.datepipe.transform(this.frommonth, 'MMMM yyyy')
        }
      ]
    };
  }

  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
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
  }


  addExcelFiltersSection(worksheet: any): number {
    let rowCount = 0;

    const report = this.getReportFilters();

    /*  TITLE (LEFT ALIGNED) */
    const titleRow = worksheet.addRow([report.title]);
    titleRow.font = { bold: true, size: 14 };
    worksheet.mergeCells(`A${rowCount + 1}:G${rowCount + 1}`);
    titleRow.alignment = { horizontal: 'left', vertical: 'middle' };
    rowCount++;

    /* FILTERS */
    report.filters.forEach((filter: any) => {
      const row = worksheet.addRow([`${filter.label}:`, filter.value]);
      row.getCell(1).font = { bold: true };
      rowCount++;
    });

    /* SPACE */
    worksheet.addRow([]);
    rowCount++;

    return rowCount;
  }
  exportToExcelAccountingMapping() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Accounting Mapping');

    /* ================= FILTER SECTION ================= */
    let rowCount = this.addExcelFiltersSection(worksheet);

    /* ================= HEADER ================= */
    const headerRowIndex = rowCount + 1;

    const headerRow = worksheet.addRow([
      'Store',
      'Acct Desc',
      'Acct #',
      'Acct Type',
      'Opera Area',
      'Dept',
      'Sub Type',
      'Sub Type Detail',
      'Fin Category',
      'MTD',
      'YTD'
    ]);

    /* ================= HEADER STYLE ================= */
    headerRow.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'D9E7FF' }
      };
      cell.font = { bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    /* ================= FORMAT FUNCTION ================= */
    const formatRow = (row: any) => {
      row.eachCell((cell: any, colNumber: number) => {

        cell.font = {
          name: 'Calibri',
          size: 11,
          ...(cell.value < 0 ? { color: { argb: 'FF0000' } } : {})
        };

        if ([1, 2, 8, 9, 10].includes(colNumber)) {
          cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
        } else if ([3, 4, 5, 6, 7].includes(colNumber)) {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        } else {
          cell.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
        }

        if (typeof cell.value === 'number') {
          if (colNumber === 11 || colNumber === 12) {
            cell.numFmt = '"$" * #,##0;[Red]"$" * -#,##0';
          }
        }
      });
    };

    /* ================= DATA ================= */
    this.AccountingMapping.forEach((am: any, i: number) => {

      const row = worksheet.addRow([
        am.as_companyName || '-',
        am.accountDescription || '-',
        am.accountNumber || '-',
        am.accountType || '-',
        am.operationalArea || '-',
        am.department || '-',
        am.subType || '-',
        am.subTypeDetail || '-',
        ['S', 'C', 'I', 'X', 'E'].includes(this.selectedheadertab[0]) ? (am.Fin_Summary || '-') : '',
        am.MTD || 0,
        am.YTD || 0
      ]);

      /* ZEBRA ROWS */
      row.eachCell(cell => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: i % 2 === 0 ? 'FFFFFF' : 'F9FBFF' }
        };
      });

      formatRow(row);
    });

    /* ================= BORDERS ================= */
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber >= headerRowIndex) { // ✅ only apply after header
        row.eachCell(cell => {
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
        });
      }
    });

    /* ================= FREEZE ================= */
    worksheet.views = [
      {
        state: 'frozen',
        xSplit: 1,
        ySplit: headerRowIndex // ✅ FIXED
      }
    ];

    /* ================= COLUMN WIDTH ================= */
    worksheet.columns = [
      { width: 35 },
      { width: 35 },
      { width: 15 },
      { width: 20 },
      { width: 20 },
      { width: 18 },
      { width: 15 },
      { width: 20 },
      { width: 25 },
      { width: 20 },
      { width: 18 },
      { width: 18 }
    ];

    /* ================= DOWNLOAD ================= */
    workbook.xlsx.writeBuffer().then(data => {
      FileSaver.saveAs(new Blob([data]), 'AccountingMapping.xlsx');
    });
  }
  private createPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });

    /* ================= TITLE ================= */
    doc.setFontSize(14);
    doc.text('Accounting Mapping', 14, 12);

    const startY = 18;

    /* ================= HEADERS (SAME AS EXCEL) ================= */
    const head = [[
      'Store',
      'Acct Desc',
      'Acct #',
      'Acct Type',
      'Opera Area',
      'Dept',
      'Sub Type',
      'Sub Type Detail',
      'Fin Category',
      'MTD',
      'YTD'
    ]];

    /* ================= BODY ================= */
    const body: any[] = [];

    this.AccountingMapping.forEach((am: any, i: number) => {

      body.push([
        am.as_companyName || '-',
        am.accountDescription || '-',
        am.accountNumber || '-',
        am.accountType || '-',
        am.operationalArea || '-',
        am.department || '-',
        am.subType || '-',
        am.subTypeDetail || '-',
        ['S', 'C', 'I', 'X', 'E'].includes(this.selectedheadertab[0])
          ? (am.Fin_Summary || '-')
          : '',
        this.formatCurrency(am.MTD || 0),
        this.formatCurrency(am.YTD || 0)
      ]);

    });

    /* ================= TABLE ================= */
    autoTable(doc, {
      startY,
      head,
      body,

      theme: 'grid',

      styles: {
        fontSize: 9,
        cellPadding: 2,
        halign: 'left',
        lineWidth: 0.1,
        lineColor: [200, 200, 200],
        textColor: [20, 20, 20],
        valign: 'middle'
      },

      headStyles: {
        fontSize: 9,
        halign: 'center',
        fillColor: [217, 231, 255], // same as excel
        textColor: [0, 0, 0],
        fontStyle: 'bold'
      },

      /* ================= COLUMN ALIGN ================= */
      columnStyles: {
        0: { halign: 'left' },
        1: { halign: 'left' },
        2: { halign: 'center' },
        3: { halign: 'center' },
        4: { halign: 'center' },
        5: { halign: 'center' },
        6: { halign: 'center' },
        7: { halign: 'left' },
        8: { halign: 'left' },
        9: { halign: 'right' },
        10: { halign: 'right' }
      },

      /* ================= ROW STYLING ================= */
      didParseCell: (data: any) => {

        if (data.section === 'body') {

          /* ZEBRA STRIPES */
          if (data.row.index % 2 !== 0) {
            data.cell.styles.fillColor = [249, 251, 255];
          }

          /* NEGATIVE VALUES RED */
          const val = data.cell.raw;
          if (typeof val === 'string') {
            const num = parseFloat(val.replace(/[$,]/g, ''));
            if (!isNaN(num) && num < 0) {
              data.cell.styles.textColor = [255, 0, 0];
            }
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
      doc.save('Account Mapping.pdf'); // ✅ only here
    } catch (e) {
      console.error(e);
      this.toast.show('Error generating PDF', 'danger', 'Error');
    } finally {
      this.shared.spinner.hide();
    }
  }
  printPDF() {
    try {
      const doc = this.createPDF();
      const pdfBlob = doc.output('blob');
      const pdfUrl = URL.createObjectURL(pdfBlob);

      const isEdge = /Edg/.test(navigator.userAgent);

      // ✅ EDGE / POPUP FLOW
      if (isEdge) {
        const win = window.open(pdfUrl);

        if (win) {
          win.onload = () => {
            setTimeout(() => {
              win.focus();

              // ✅ Handle after print properly
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

      // ✅ IFRAME FLOW (BEST FOR CHROME)
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

              // ✅ Handle print completion
              cw.onafterprint = () => {
                document.body.removeChild(iframe);
                URL.revokeObjectURL(pdfUrl);
              };

              cw.print();
            } catch (err) {
              console.warn('Iframe print failed, fallback triggered');

              // ✅ Fallback to new tab
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
      const pdfFile = this.blobToFile(pdfBlob, 'Financial Summary.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Financial Summary');
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
  private formatNumber(val: any): string {
    if (val == null || val === '' || val === 0) return '-';
    return Math.round(Number(val)).toLocaleString('en-US');
  }

  private formatCurrency(val: any): string {
    if (val == null || val === '' || val === 0) return '-';
    return '$ ' + Math.round(Number(val)).toLocaleString('en-US');
  }

  private formatPercent(val: any): string {
    if (val == null || val === '') return '-';

    const num = Number(val);

    // If whole number → no decimals
    if (Number.isInteger(num)) {
      return num + '%';
    }

    // If decimal → show max 2 decimals (no trailing zeros)
    return parseFloat(num.toFixed(2)) + '%';
  }


}
