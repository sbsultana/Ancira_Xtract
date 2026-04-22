import { Component, ElementRef, HostListener, Injector, ViewChild } from '@angular/core';
import { Api } from '../../../../Core/Providers/Api/api';
import { common } from '../../../../common';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { FormsModule } from '@angular/forms';
import { NgbActiveModal, NgbModal, NgbModule } from '@ng-bootstrap/ng-bootstrap';
import { NgxSpinnerService } from 'ngx-spinner';
import { CurrencyPipe, DatePipe, formatDate } from '@angular/common';
import { Title } from '@angular/platform-browser';
import { Subscription } from 'rxjs';
import { environment } from '../../../../../environments/environment';
import { Workbook } from 'exceljs';
import FileSaver from 'file-saver';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
import numeral from 'numeral';
import { BsDatepickerConfig, BsDatepickerDirective, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Router } from '@angular/router';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastContainer } from '../../../../Layout/toast-container/toast-container';
declare var bootstrap: any;

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, ToastContainer, Stores, NgbModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  Current_Date: any;
  subscription!: Subscription;
  subscriptionExcel!: Subscription;
  subscriptionReport!: Subscription;
  FSData: any = [];
  date: any;
  ExpenseTrendByStoreKeys: any = [];
  AllDatakeys: any = [];
  ExpenseTrendByStore_Excel: any;
  XpenseTrendByStoreKeys: any = [];
  ExpenseTrendByStore: any;
  ExpenseTrendByStore_ExcelMonth: any;
  XpenseTrendByStoreKeysMonth: any = [];
  ExpenseTrendByStoreKeysMonth: any = [];
  AllDatakeysMonth: any = [];
  ExpenseTrendByStoreMonth: any;
  SubFilter: any;
  SelectedTab: any = [];
  Month!: any;
  // stores: any;
  selectedstorevalues: any = [];
  selectedstorename: any;
  selectedFilters: string[] = [];
  selectedLabel: string = '( All )';
  StoreName: any = 'All Stores';
  Filter: any = ['New', 'Used', 'Service', 'Parts', 'Detail'];
  storeIds: any = 2;
  groups: any = 1;
  PresentDayDate: string;

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }


  stores: any = [];
  selectDate: Date = new Date();
  currentMonth: any = '';
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  constructor(
    private datepipe: DatePipe,
    public apiSrvc: Api,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private spinner: NgxSpinnerService,
    private title: Title,
    private comm: common,
    private router: Router,
    private toast: ToastService,
    private injector: Injector,
    public shared: Sharedservice,
  ) {
    const lastMonth = new Date();
    let today = new Date();
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')
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
    if (today.getDate() < 5) {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth() - 1));
    } else {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth()));
    }
    this.title.setTitle(this.comm.titleName + '-Net Adds / Deducts');
    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Net Adds / Deducts',
        path1: '',
        path2: '',
        path3: '',
        Month: this.date,
        stores: this.storeIds.toString(),
        store: this.storeIds,
        filter: this.Filter,
        groups: 1,
        count: 0,
      };
      this.apiSrvc.SetHeaderData({
        obj: data,
      });
    }
    const format = 'ddMMyyyy';
    const locale = 'en-US';
    const myDate = new Date();
    const formattedDate = formatDate(myDate, format, locale);
    this.PresentDayDate = formattedDate;
    this.selectDate = this.date
    this.currentMonth = this.selectDate;
    this.selectedFilters = [...this.filters];
    this.GetDataByMonths(this.currentMonth, this.selectedFilters);
  }
  roleId: any;
  ngOnInit(): void {
    this.roleId = localStorage.getItem('roleId');
    console.log('role Id', this.roleId);
    this.getStores();

  }

  StoreNamesHeadings: any = [];
  MonthsHeadings: any = [];

  Scrollpercent: any = 0;
  scrollPositionStoring: number = 0;
  scrollCurrentPosition: number = 0;

  @ViewChild('scrollcent') scrollcent!: ElementRef;

  updateVerticalScroll(event: any): void {
    this.scrollCurrentPosition = event.target.scrollTop;
  }

  // Filters list
  filters: string[] = ['New', 'Used', 'Service', 'Parts', 'Detail'];
  activePopover: number | null = null;
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month',
    maxDate: new Date()
  };
  monthPicker!: BsDatepickerDirective;
  openMonthPicker() {
    if (this.monthPicker) {
      this.monthPicker.show();
    }
  }

  togglePopover(index: number) {
    this.activePopover = this.activePopover === index ? null : index;
  }


  isSelected(filter: string): boolean {
    return this.selectedFilters.includes(filter);
  }

  toggleFilter(filter: string) {
    const index = this.selectedFilters.indexOf(filter);
    if (index >= 0) {
      this.selectedFilters.splice(index, 1);
    } else {
      this.selectedFilters.push(filter);
    }
    this.updateLabel();
  }

  updateLabel() {
    if (!this.selectedFilters || this.selectedFilters.length === 0) {
      this.selectedLabel = '( Select )';
    } else if (this.selectedFilters.length === this.filters.length) {
      this.selectedLabel = '( All )';
    } else if (this.selectedFilters.length === 1) {
      this.selectedLabel = `( ${this.selectedFilters[0]} )`;
    } else {
      this.selectedLabel = `( Selected ${this.selectedFilters.length} )`;
    }
  }



  applyDateChange() {
    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      return;
    }
    if (!this.selectedFilters || this.selectedFilters.length === 0) {
      this.toast.show('Please Select Atleast One Department', 'warning', 'Warning');
      return;
    }
    else {
      this.currentMonth = this.formatMonth(this.selectDate);
      this.GetDataByMonths(this.currentMonth, this.selectedFilters);
      this.activePopover = null;
      this.isLoading = true;
    }
  }
  formatMonth(date: Date): string {
    const year = date.getFullYear();
    const month = ('0' + (date.getMonth() + 1)).slice(-2);
    return `${year}-${month}`;
  }
  isLoading = true;
  NoData = false;
  EndDate: any;
  dates: any;
  storeName: any;
  ExpenseTrendKeys: any = [];
  IncomeSummaryData: any;
  IncomeSummaryDataKeys: any;
  ShowHideGP: any = 'Hide';
  ShowHideBudget: any = 'Hide';
  GetDataByMonths(date: any, filters: any) {
    this.spinner.show();
    this.dates = [];
    this.ExpenseTrendByStoreMonth = [];
    this.IncomeSummaryDataKeys = [];
    this.IncomeSummaryData = [];
    const currentDate = new Date(date);
    const currentMonth = currentDate.getMonth();
    const currentYear = currentDate.getFullYear();
    for (let i = 0; i < 12; i++) {
      const lastMonthDate = new Date(currentYear, currentMonth - i, 1);
      this.dates.push(this.formatDate(lastMonthDate));
    }
    const DateToday = this.datepipe.transform(
      new Date(date),
      'yyyy-MM-dd'
    );
    const obj = {
      "ReportType": "S",
      "As_ids": this.storeIds.toString(),
      "Date": DateToday,
      "Storenames": "",
      "Account": ""
    };
    console.log(obj);
    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetNetAddsDeducts', obj)
      .subscribe(
        (x: any) => {
          const currentTitle = document.title;
          if (x.status == 200 && x.response) {
            this.IncomeSummaryData = x.response;
            console.log('IncomeSummaryData', this.IncomeSummaryData);
            const IncomeSummaryKeys = Object.keys(x.response[0] || {}).slice(3);
            const lastTwoValues = IncomeSummaryKeys.slice(-2);
            const remainingValues = IncomeSummaryKeys.slice(0, -2);
            this.IncomeSummaryDataKeys = lastTwoValues.concat(remainingValues);

            this.IncomeSummaryDataKeys = this.IncomeSummaryDataKeys.filter((dealership: any) =>
              dealership !== 'TOTAL' &&
              dealership !== 'AVG' &&
              dealership !== 'SEQ' &&
              dealership !== 'AS_OF' &&   // ✅ remove AS_OF column
              (this.ShowHideGP === 'Show' || (this.ShowHideGP === 'Hide' && !dealership.endsWith('_GP'))) &&
              (this.ShowHideBudget === 'Show' ||
                (this.ShowHideBudget === 'Hide' &&
                  !dealership.startsWith('BDGT_') &&
                  !dealership.startsWith('VAR')))
            );

            if (this.IncomeSummaryData.length > 0) {
              this.IncomeSummaryData.forEach((x: any) => {
                (x.SubData = []), (x.data2sign = '+');
              });

              console.log('Keys ', this.IncomeSummaryDataKeys);
              this.NoData = this.IncomeSummaryDataKeys.length === 0;
            } else {
              this.NoData = true;
            }
          } else {
            this.NoData = true;
          }

          this.spinner.hide();
        },
        (error) => {
          console.error(error);
          this.NoData = true;
          this.spinner.hide();
        }
      );
  }

  formatKey(key: string): string {
    if (key.startsWith('BDGT_')) return 'BUDGET';
    if (key.startsWith('VAR')) return 'VARIANCE';
    return key;
  }
  isKeyVisible(key: string): boolean {
    return this.IncomeSummaryDataKeys.includes(key);
  }
  isKeyVisibles(key: string): boolean {
    return key !== '-1';

  }
  isNumber(value: any): boolean {
    return (
      value !== null && value !== undefined && value !== 0 && !isNaN(value)
    );
  }
  isPercentage(value: any): boolean {
    return typeof value === 'string' && value.includes('%');
  }
  isSingleStoreId(): boolean {
    return this.storeIds && this.storeIds.toString().includes(',');
  }

  save() {
    console.log('Save clicked!');
    // you can put static save logic here
  }

  initializeModalDraggable() {
    const modal = document.getElementById('payplan');

    if (modal) {
      modal.addEventListener('shown.bs.modal', () => {
        this.makeModalDraggable();
      });
    }
  }

  // private isDragging = false;
  // private offsetX = 0;
  // private offsetY = 0;
  private modal!: HTMLElement | null;
  private modalHeader!: HTMLElement | null;

  private isDragging = false;
  private offsetX = 0;
  private offsetY = 0;

  makeModalDraggable() {
    const modal = document.querySelector('.recivables-draggable-modal') as HTMLElement;
    const header = modal?.querySelector('.modal-header') as HTMLElement;

    if (!modal || !header) return;

    modal.style.width = '1200px';

    header.onmousedown = (event: MouseEvent) => {
      this.isDragging = true;

      this.offsetX = event.clientX - modal.getBoundingClientRect().left;
      this.offsetY = event.clientY - modal.getBoundingClientRect().top;

      document.onmousemove = (e) => {
        if (!this.isDragging) return;

        modal.style.left = (e.clientX - this.offsetX) + 'px';
        modal.style.top = (e.clientY - this.offsetY) + 'px';
      };

      document.onmouseup = () => {
        this.isDragging = false;
        document.onmousemove = null;
      };
    };
  }

  onMouseMove = (event: MouseEvent) => {
    if (!this.isDragging || !this.modal) return;

    this.modal.style.left = `${event.clientX - this.offsetX}px`;
    this.modal.style.top = `${event.clientY - this.offsetY}px`;
  };

  stopDragging = () => {
    this.isDragging = false;
    document.removeEventListener('mousemove', this.onMouseMove);
    document.removeEventListener('mouseup', this.stopDragging);
  };


  disabledSubTypes = [
    'Total Report'
  ];

  isDisabled(item: any, key: string): boolean {
    return (
      item[key] == 0 ||
      item[key] == null ||
      item[key] === '' ||
      this.disabledSubTypes.includes(item['LINE HEAD'])
    );
  }

  formatDate(date: Date): string {
    const options: Intl.DateTimeFormatOptions = {
      year: 'numeric',
      month: 'short',
    };
    return date.toLocaleDateString(undefined, options);
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  SFstate: any;
  StoreCodes: any;
  block: any = '';
  Report: any = '';
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Net Adds / Deducts') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.subscriptionReport = this.apiSrvc
      .GetReportOpening()
      .subscribe((res) => {
        console.log(res);
        if (this.subscriptionReport != undefined) {
          if (res.obj.Module == 'Net Adds / Deducts') {
            document.getElementById('report')?.click();
          }
        }
      });
    this.apiSrvc.GetReports().subscribe((data) => {
      if (data.obj.Reference == 'Net Adds / Deducts') {
        console.log(data);
        if (data.obj.header == undefined) {
          this.date = data.obj.month;
          this.Month = data.obj.month;
          this.storeIds = data.obj.storeValues;
          this.StoreName = data.obj.Sname;
          this.Filter = data.obj.filter;
          this.SubFilter = data.obj.subfilters;
          this.StoreCodes = data.obj.storecode;
          this.groups = data.obj.groups;
          this.index = '';
          this.Scrollpercent = 0;
        } else {
          if (data.obj.header == 'Yes') {
            this.storeIds = data.obj.storeValues;
          }
        }
        this.GetDataByMonths(this.currentMonth, this.selectedFilters);
        if (this.storeIds != '') {
          this.goToFirstPage();
        } else {
          this.NoData = true;
          this.filteredETdetailsData = [];
        }
        const headerdata = {
          title: 'Net Adds / Deducts',
          path1: '',
          path2: '',
          path3: '',
          Month: new Date(this.Month),
          filter: this.Filter,
          stores: this.storeIds,
          storecode: this.StoreCodes,
          Sname: this.StoreName,
          groups: this.groups,
        };
        this.apiSrvc.SetHeaderData({
          obj: headerdata,
        });
      }
    });
    this.subscriptionExcel = this.apiSrvc
      .getExportToExcelAllReports()
      .subscribe((res) => {
        this.SFstate = res.obj.state;
        if (this.subscriptionExcel != undefined) {
          if (res.obj.title == 'Net Adds / Deducts') {
            if (res.obj.state == true) {
              if (this.ExpenseTrend == true) {
                this.exportAsXLSX();
              } else if (this.ExpenseTrendDetails == true) {
                // this.exportToExcel();
              }
            }
          }
        }
      });

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Net Adds / Deducts') return;
      if (obj.state) {
        this.exportAsXLSX();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Net Adds / Deducts') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Net Adds / Deducts') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Net Adds / Deducts') return;
      if (obj.stateEmailPdf) {
        this.sendEmailData(obj.Email, obj.notes, obj.from);
      }
    });
    this.makeModalDraggable()
  }
  ngOnDestroy(): void {
    this.excel?.unsubscribe();
    this.print?.unsubscribe();
    this.pdf?.unsubscribe();
    this.email?.unsubscribe();
  }
  reportOpen(temp: any) {
    this.ngbmodalActive = this.ngbmodal.open(temp, {
      size: 'xl',
      backdrop: 'static',
    });
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
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
  }
  closeReport() {
    this.Report = '';
  }
  getStores() {
    this.selectedstorevalues = [];
    this.stores = JSON.parse(localStorage.getItem('Stores')!);
  }
  ETdetailsData: any = [];
  currentPage: number = 1;
  itemsPerPage: number = 100;
  maxPageButtonsToShow: number = 3;
  clickedPage: number | null = null;
  filteredETdetailsData: any[] = [];
  selectedDate: any;
  DateType: any;
  DetailsSearchName: any;
  Lable: any;
  searchText: string = '';
  ExpenseTrend: boolean = true;
  ExpenseTrendDetails: boolean = false;
  SubtypeDetailLable: any;
  FinSummaryLable: any;
  MonthDate: any;
  spinnerLoader!: boolean;
  Dept: any;
  openMonthDetails(
    Object: any,
    storename: any,
    ref: any,
    item: any,
    DateMethod: any
  ) {
    console.log('Object', Object);
    console.log('storename', storename);
    console.log('ref', ref);
    console.log('item', item);
    console.log('DateMethod', DateMethod);
    const modalEl = document.getElementById('netaddbydeducts');
    if (modalEl) {
      const modal = new bootstrap.Modal(modalEl);
      modal.show();
    }
    this.NoData = false;
    // this.spinner.show();
    this.ExpenseTrend = false;
    this.ExpenseTrendDetails = true;
    this.ETdetailsData = [];
    this.spinnerLoader = true;
    this.DateType = DateMethod;
    this.Lable = ref;
    console.log(storename);
    const DateToday = this.datepipe.transform(
      new Date(this.selectDate),
      'yyyy-MM-dd'
    );
    this.SubtypeDetailLable = Object['LINE HEAD'];
    this.StoreName = storename;
    this.Dept = Object.Category
    let match = ref.match(/\(([^)]+)\)/);
    let refname = match ? match[1] : null;
    const Obj = {
      "ReportType": "D",
      "As_ids": "",
      "Date": DateToday,
      "Storenames": storename,
      "Account": refname
    };
    console.log(Obj);
    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetNetAddsDeducts', Obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.ETdetailsData = res.response;
          this.filterData();
          console.log('ET Details', this.ETdetailsData);
          // this.spinner.hide();
          this.spinnerLoader = false;
          this.NoData = true;
        }
      });
  }
  clearArray() {
    this.ETdetailsData = [];
    this.filteredETdetailsData = [];
  }
  expandedIndex: number | null = null;
  FSSubDetailsMap: { [index: number]: any[] } = {};

  GetSubDetails(AcctNo: any, StoreName: any, index: number) {
    if (this.expandedIndex === index) {
      this.expandedIndex = null;
      return;
    }
    this.expandedIndex = index;

    this.spinner.show();
    const Obj = {
      as_ids: StoreName,
      // dept: this.selectedFilters.toString(),
      dept: this.selectedFilters.toString(),
      subtype: '',
      subtypedetail: this.SubtypeDetailLable,
      FinSummary: this.FinSummaryLable,
      date: this.selectedDate,
      accountnumber: AcctNo
    };

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetExpenseTrendDetailsV1', Obj)
      .subscribe((res) => {
        this.spinner.hide();
        if (res.status === 200) {
          this.FSSubDetailsMap[index] = res.response;
        }
      });
  }

  backtoWR() {
    this.ExpenseTrend = true;
    this.ExpenseTrendDetails = false;
    this.filteredETdetailsData = [];
    this.ETdetailsData = [];
    this.expandedIndex = null;
    if (this.storeIds != '') {
      this.goToFirstPage();
    }
    setTimeout(() => {
      if (this.scrollcent && this.scrollcent.nativeElement) {
        this.scrollcent.nativeElement.scrollTop = this.scrollPositionStoring;
      }
    });
  }

  get postingAmountTotal(): number {
    return this.filteredETdetailsData.reduce((total, item) => {
      return total + (item.postingamount || 0);
    }, 0);
  }
  getPostingSubAmountTotal(index: number): number {
    const subRows = this.FSSubDetailsMap[index];
    if (!subRows || !Array.isArray(subRows)) {
      return 0;
    }

    return subRows.reduce((total, item) => {
      return total + (item.PostingAmount || 0);
    }, 0);
  }
  filterData() {
    const text = this.searchText.trim().toLowerCase();

    if (!text) {
      this.filteredETdetailsData = [...this.ETdetailsData];
    } else {
      this.filteredETdetailsData = this.ETdetailsData.filter((item: any) =>
        [
          item.StoreName,
          item.AccountNumber,
          item.AccountDescription,
          item.PostingAmount
        ].some(val =>
          val?.toString().toLowerCase().includes(text)
        )
      );
    }

    this.currentPage = 1;
  }
  get paginatedItems() {
    const start = (this.currentPage - 1) * this.itemsPerPage;
    return this.filteredETdetailsData.slice(start, start + this.itemsPerPage);
  }

  getMaxPageNumber(): number {
    return Math.max(1,
      Math.ceil(this.filteredETdetailsData.length / this.itemsPerPage)
    );
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
    return (this.currentPage - 1) * this.itemsPerPage + 1;
  }

  getEndRecordIndex(): number {
    return Math.min(
      this.getStartRecordIndex() + this.itemsPerPage - 1,
      this.filteredETdetailsData.length
    );
  }
  onPageSizeChange() {
    this.currentPage = 1;
  }

  openStoresDetails(Object: any, StoreName: any, date: any, item: any) {
    this.index = '';
    console.log(Object);
    console.log(item);
    const myDate = this.datepipe.transform(new Date(date), 'MMM-yyyy');
    let index = this.stores.filter(
      (store: any) => store.DEALER_NAME == StoreName
    );
    this.selectedstorevalues = index[0].AS_ID;
    this.selectedstorename = index[0].DEALER_NAME;
    const DetailsSF = this.ngbmodal.open({
      size: 'xl',
      backdrop: 'static',
    });
    DetailsSF.componentInstance.ETdetails = {
      TYPE: item.LABLEVAL,
      NAME: item.LABLE,
      STORES: this.selectedstorevalues,
      LatestDate: myDate,
      STORENAME: this.selectedstorename,
    };
  }
  ValueFormat: any;
  openGraph(monthname: any, dates: any, Obj: any, SummaryType: any) {
    const modalRef = this.ngbmodal.open({
      size: 'xl',
      backdrop: 'static',
      injector: Injector.create({
        providers: [
          { provide: CurrencyPipe, useClass: CurrencyPipe }
        ],
        parent: this.injector
      })
    });
    const ValueFormat =
      Obj.DISPLAY_LABLE.includes('%') ? 'Percentage' : 'Currency';

    modalRef.componentInstance.ETgraphdetails = {
      ITEM: Obj,
      DATES: dates,
      NAME: Obj.DISPLAY_LABLE,
      ValueFormat,
      STORES: this.storeIds,
      SUMMARYTYPE: SummaryType,
    };
  }

  openStoreGraph(
    Object: any,
    StoreName: any,
    date: any,
    item: any,
    SummaryType: any
  ) {
    console.log(Object, StoreName, date, item, SummaryType);
    const DetailsSF = this.ngbmodal.open({
      size: 'xl',
      backdrop: 'static',
    });
    var DisplayName = item.DISPLAY_LABLE;
    DetailsSF.componentInstance.ETgraphdetails = {
      ITEM: item,
      TYPE: item.DISPLAY_LABLE,
      NAME: SummaryType,
      SUMMARYTYPE: SummaryType,
    };
  }

  index = '';
  commentobj: any = {};
  close(data: any) {
    console.log(data);
    this.index = '';
  }
  addcmt(data: any) {
    if (data == 'A') {
      this.index = '';
      const DetailsSF = this.ngbmodal.open({
        size: 'xl',
        backdrop: 'static',
      });
      // myObject['skillItem2'] = 15;
      this.commentobj['state'] = 0;
      (DetailsSF.componentInstance.SFComments = this.commentobj),
        DetailsSF.result.then(
          (data) => {
            console.log(data);
          },
          (reason) => {
            console.log(reason);

            if (reason == 'O') {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
            } else {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
              if (this.Filter == 'VariableTrendsvsBudget') {
                this.GetDataByMonths(this.currentMonth, this.selectedFilters);
              }
            }
            // // on dismiss

            // const Data = {
            //   state: true,
            // };
            // this.apiSrvc.setBackgroundstate({ obj: Data });
            // this.GetData();
          }
        );
    }
    if (data == 'AD') {
      if (this.Filter == 'VariableTrendsvsBudget') {
        this.GetDataByMonths(this.currentMonth, this.selectedFilters);
      }
    }
  }
  generatePDF() {
    if (this.storeIds.length >= 16) {
      this.toast.show('Please select only up to 15 stores', 'warning', 'Warning');
      return;
    }

    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save(`Net Adds / Deducts.pdf`);
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
  private createPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a2' });

    let startY = 12;

    /* ================= FILTER SECTION (LIKE EXCEL) ================= */
    // const report = this.getReportFilters();

    doc.setFontSize(14);
    // doc.text(report.title, 14, startY);

    startY += 6;

    doc.setFontSize(10);

    // report.filters.forEach((f: any) => {
    //   doc.text(`${f.label}: ${f.value}`, 14, startY);
    //   startY += 5;
    // });

    // startY += 3;

    /* ================= HEADERS ================= */

    const Header = ['TOTAL', 'AVERAGE', ...this.IncomeSummaryDataKeys];

    // Row 1 (BLUE)
    const headerRow1 = ['', ...Header];

    // Row 2 (LIGHT BLUE)
    const headerRow2 = [
      'Net Adds / Deducts',
      ...new Array(Header.length).fill('')
    ];

    /* ================= DATA ================= */

    const body: any[] = [];

    const IncomeSummaryData = this.IncomeSummaryData.map((x: any) => ({ ...x }));

    IncomeSummaryData.forEach((d: any) => {

      let row = [
        d["LINE HEAD"] || '-',
        d.TOTAL ?? '-',
        d.AVG ?? '-'
      ];

      this.IncomeSummaryDataKeys.forEach((e: any) => {
        row.push(d[e] === '' || d[e] === null ? '-' : Number(d[e]));
      });

      body.push(row);
    });

    /* ================= TABLE ================= */

    autoTable(doc, {
      startY,
      head: [headerRow1, headerRow2],
      body,
      theme: 'grid',

      styles: {
        fontSize: 8,
        cellPadding: 2,
        textColor: [30, 30, 30]
      },
      /* ✅ COLUMN WIDTH CONTROL */
      columnStyles: {
        0: { cellWidth: 50 }, // 1st column (Label)
        1: { cellWidth: 30 }, // 2nd column
        // remaining columns → auto (no need to define)
      },
      didParseCell: (data: any) => {

        const rowData = IncomeSummaryData[data.row.index];

        /* ================= HEADER STYLING ================= */

        if (data.section === 'head') {

          // Row 1 → BLUE
          if (data.row.index === 0) {
            data.cell.styles.fillColor = [5, 84, 239];
            data.cell.styles.textColor = [255, 255, 255];
            data.cell.styles.fontStyle = 'bold';
            data.cell.styles.halign = 'center';
          }

          // Row 2 → LIGHT BLUE
          if (data.row.index === 1) {
            data.cell.styles.fillColor = [217, 231, 255];
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.fontStyle = 'bold';

            if (data.column.index === 0) {
              data.cell.styles.halign = 'left';
            } else {
              data.cell.styles.halign = 'center';
            }
          }

          return;
        }

        /* ================= ALIGNMENT ================= */

        if (data.column.index === 0) {
          data.cell.styles.halign = 'left';
        } else {
          data.cell.styles.halign = 'right';
        }

        /* ================= EMPTY ================= */

        if (!data.cell.raw || data.cell.raw === '') {
          data.cell.text = ['-'];
          data.cell.styles.halign = 'center';
          return;
        }

        /* ================= NUMBER FORMAT ================= */

        if (data.column.index >= 1 && typeof data.cell.raw === 'number') {

          const val = Math.round(data.cell.raw);

          if (val === 0) {
            data.cell.text = ['-'];
            data.cell.styles.halign = 'center';
            return;
          }

          data.cell.text = [`$ ${val.toLocaleString()}`];

          if (val < 0) {
            data.cell.styles.textColor = [255, 0, 0];
          }
        }

        /* ================= ALTERNATE ROW ================= */

        if (data.row.index % 2 === 0) {
          data.cell.styles.fillColor = [245, 247, 250];
        }

        /* ================= REPORT TOTAL ================= */

        const isReportTotal =
          (rowData?.["LINE HEAD"] || '')
            .toString()
            .toLowerCase()
            .trim() === 'report total';

        if (isReportTotal) {
          data.cell.styles.fillColor = [141, 180, 255];
          data.cell.styles.fontStyle = 'bold';
        }

        /* ================= DISPLAY HEAD ================= */

        if (rowData?.DISPLAYHEAD_FLAG === 1) {
          data.cell.styles.fillColor = [148, 182, 209];
          data.cell.styles.textColor = [255, 255, 255];
          data.cell.styles.fontStyle = 'bold';
        }

        /* ================= HEAD TOTAL ================= */

        if (rowData?.ISHEAD_TOTAL === 'Y') {
          data.cell.styles.fillColor = [141, 180, 255];
          data.cell.styles.fontStyle = 'bold';
        }

      },


    });

    return doc;
  }


  getSelectedStoreNames(): string {
    if (!this.storeIds || this.storeIds.length === 0) return '';

    const ids = this.storeIds.toString().split(',');

    const selectedStores = this.stores.filter((s: any) =>
      ids.includes(s.ID.toString())
    );

    return selectedStores.map((s: any) => s.storename).join(', ');
  }
  getReportFilters(): { title: string; filters: any[] } {
    return {
      title: 'Net Adds / Deducts',
      filters: [
        {
          label: 'Store',
          value: this.getSelectedStoreNames() || 'All Stores'
        },
        {
          label: 'Group',
          value: this.groupName || ''
        },
        {
          label: 'Month',
          value: this.datepipe.transform(this.currentMonth, 'MMMM yyyy')
        },
      ]
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
  ExcelStoreNames: any = [];
  exportAsXLSX(): void {

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Fixed Income Expense');

    /* ================= FILTER SECTION ================= */
    const filterRowCount = this.addExcelFiltersSection(worksheet);

    /* ================= COMMON FORMAT ================= */
    const formatRow = (row: any) => {
      row.eachCell((cell: any, colNumber: number) => {

        cell.font = { name: 'Calibri', size: 11 };

        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };

        if (colNumber === 1) {
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        } else {
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        }

        if (!cell.value || cell.value === '-') {
          cell.value = '-';
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          return;
        }

        const num = Number(cell.value);

        if (!isNaN(num)) {

          if (num === 0) {
            cell.value = '-';
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            return;
          }

          if (colNumber >= 2) {
            cell.numFmt = '"$" * #,##0;[Red]"$" * -#,##0';
          }

          if (num < 0) {
            cell.font = {
              ...cell.font,
              color: { argb: 'FFFF0000' }
            };
          }
        }
      });

      if (row.number % 2 === 0) {
        row.eachCell((cell: any) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF5F7FA' }
          };
        });
      }
    };

    /* ================= HEADER ================= */

    const Header = ['TOTAL', 'AVERAGE', ...this.IncomeSummaryDataKeys];

    /* ===== ROW 1 → APPLY TITLE STYLE ===== */
    const headerRow1 = worksheet.addRow(['', ...Header]);

    headerRow1.eachCell((cell: any) => {
      cell.font = {
        bold: true,
        name: 'Calibri',
        size: 11,
        color: { argb: 'FFFFFFFF' } // WHITE TEXT
      };

      cell.alignment = {
        horizontal: 'center',
        vertical: 'middle'
      };

      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0554EF' } // 🔵 TITLE BLUE
      };

      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });


    /* ===== ROW 2 → APPLY HEADER STYLE ===== */
    const headerRow2 = worksheet.addRow([
      'Net Adds / Deducts',
      ...new Array(Header.length).fill('')
    ]);

    headerRow2.eachCell((cell: any, colNumber: number) => {

      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "D9E7FF" } // LIGHT BLUE
      };

      cell.font = {
        bold: true,
        name: 'Calibri',
        size: 11,
        color: { argb: 'FF000000' }
      };

      // Alignment
      if (colNumber === 1) {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      } else {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      }

      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    /* ================= DATA ================= */

    const IncomeSummaryData = this.IncomeSummaryData.map((x: any) => ({ ...x }));

    IncomeSummaryData.forEach((d: any) => {

      let Obj = [
        d["LINE HEAD"] || '-',
        d.TOTAL ?? '-',
        d.AVG ?? '-'
      ];

      this.IncomeSummaryDataKeys.forEach((e: any) => {
        Obj.push(d[e] === '' || d[e] === null ? '-' : Number(d[e]));
      });

      const row = worksheet.addRow(Obj);

      formatRow(row);
      // ===== REPORT TOTAL ROW STYLE =====
      const isReportTotal =
        (d["LINE HEAD"] || '')
          .toString()
          .toLowerCase()
          .trim() === 'report total';

      if (isReportTotal) {
        row.eachCell((cell: any) => {
          cell.font = {
            ...cell.font,
            bold: true,
            name: 'Calibri',
            size: 11
          };

          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '8DB4FF' } // ✅ Blue
          };
        });
      }
      if (d.DISPLAYHEAD_FLAG === 1) {
        row.eachCell((cell: any) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '94b6d1' }
          };
          cell.font = {
            bold: true,
            name: 'Calibri',
            size: 11,
            color: { argb: 'FFFFFFFF' }
          };
        });
      }

      if (d.ISHEAD_TOTAL === 'Y') {
        row.eachCell((cell: any) => {
          cell.font = { bold: true, name: 'Calibri', size: 11 };
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '8DB4FF' }
          };
        });
      }
    });

    /* ================= COLUMN WIDTH ================= */
    worksheet.getColumn(1).width = 30;

    for (let i = 2; i <= Header.length + 1; i++) {
      worksheet.getColumn(i).width = 25;
    }

    /* ================= FREEZE ================= */
    worksheet.views = [{
      state: "frozen",
      ySplit: 6
    }];

    /* ================= DOWNLOAD ================= */
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, "Net Adds / Deducts.xlsx");
    });
  }

  Favreports: any = [];

  isDesc: boolean = false;
  column: string = 'CategoryName';

  sort(property: string, data: any[], state?: any) {
    if (state === undefined) {
      this.isDesc = this.column === property ? !this.isDesc : false;
    }
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    data.sort((a, b) => {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }


  GetPrintData() {
    window.print();
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
        this.shared.spinner.hide();
        return;
      }
      const pdfFile = this.blobToFile(pdfBlob, 'Net Adds / Deducts.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Net Adds / Deducts');
      formData.append('file', pdfFile);
      formData.append('notes', notes);
      formData.append('from', from);

      this.apiSrvc.postmethod(this.comm.routeEndpoint + 'mail', formData)
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

  generatePDFDetails() {
    this.spinner.show();
    const printContents = document.getElementById(
      'ExpenseTrendDetailsDownload'
    )!.innerHTML;
    const iframe = document.createElement('iframe');

    // Make the iframe invisible
    iframe.style.position = 'absolute';
    iframe.style.width = '0px';
    iframe.style.height = '0px';
    iframe.style.border = 'none';

    document.body.appendChild(iframe);

    const doc = iframe.contentWindow?.document;
    if (!doc) {
      console.error('Failed to create iframe document');
      return;
    }

    doc.open();
    doc.write(`
    <html>
    <head>
    <title>Net Adds / Deducts</title>
    <style>
        @media print {
        body {-webkit-print-color-adjust: exact;}
        }
      @font-face {
       font-family: 'GothamBookRegular';
       src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
            url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
     }
     @font-face {
       font-family: 'Roboto';
       src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
            url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
     }
     @font-face {
       font-family: 'RobotoBold';
       src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
            url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
     }
            .comment {
  width: fit-content;
  background-color: #d7d7d7;
  border-radius: 0 15px 0 15px;
  padding: .2rem .5rem;
  margin-left: 2rem;
}
.bdr{
  border:1px solid white !important
 }
.notes { 
  width: fit-content;  
  padding: 0.2rem 0.5rem;
  margin-right: 3rem;
  float: inline-end;
}
   .title {
  font-family: "FaktPro-Bold";
  font-size: larger;
  padding-left: 1rem;
  text-align: left !important;
}
     .performance-scorecard-details {
          display: flex;
          flex-direction: column;
          height: auto;
          /* Adjust based on your needs */
          width: 100%;
        }
        .performance-scorecard-details .table > :not(:first-child) {
          border-top: 0px solid #ffa51a;
        }
        .performance-scorecard-details .table {
          text-align: center;
          text-transform: capitalize;
          border: transparent;
          width: 100%;
        }
        .performance-scorecard-details .table th, .performance-scorecard-details .table td {
          white-space: nowrap;
          vertical-align: top;
        }
        .performance-scorecard-details .table th:first-child, .performance-scorecard-details .table td:first-child {
          left: 0;
          z-index: 1;
        }
        .performance-scorecard-details .table tr:nth-child(odd) td:first-child, .performance-scorecard-details .table tr:nth-child(odd) td:nth-child(2) {
          background-color: #e9ecef;
        }
        .performance-scorecard-details .table tr:nth-child(even) td:first-child, .performance-scorecard-details .table tr:nth-child(even) td:nth-child(2) {
          background-color: #fff;
        }
        .performance-scorecard-details .table tr:nth-child(odd) {
          background-color: #e9ecef;
        }
        .performance-scorecard-details .table tr:nth-child(even) {
          background-color: #fff;
        }
        .performance-scorecard-details .table .spacer {
          background-color: #cfd6de !important;
          border-left: 1px solid #cfd6de !important;
          border-bottom: 1px solid #cfd6de !important;
          border-top: 1px solid #cfd6de !important;
        }
        .performance-scorecard-details .table .hidden {
          display: none !important;
        }
        .performance-scorecard-details .table .bdr-rt {
          border-right: 1px solid #abd0ec;
        }
        .performance-scorecard-details .table thead {
          position: sticky;
          top: 0;
          z-index: 99;
          font-family: 'FaktPro-Bold';
          font-size: 0.8rem;
        }
        .performance-scorecard-details .table thead th {
          padding: 5px 10px;
          margin: 0px;
          border-right: 1px solid #abd0ec;
        }
        .performance-scorecard-details .table thead .bdr-btm {
          border-bottom: #005fa3;
        }
        .performance-scorecard-details .table thead tr:nth-child(1) {
          background-color: #fff !important;
          color: #000;
          text-transform: uppercase;
          border-bottom: #cfd6de;
        }
        .performance-scorecard-details .table thead tr:nth-child(1) {
          background-color: #fff !important;
          color: #000;
          text-transform: uppercase;
          border-bottom: #cfd6de;
          box-shadow: inset 0 1px 0 0 #cfd6de;
        }
        .performance-scorecard-details .table thead tr:nth-child(1) th {
          box-shadow: inset 0 -2px 0 #337ab7;
        }
        .performance-scorecard-details .table thead tr:nth-child(3) {
          background-color: #fff !important;
          color: #000;
          text-transform: uppercase;
          border-bottom: #cfd6de;
          box-shadow: inset 0 1px 0 0 #cfd6de;
        }
        .performance-scorecard-details .table thead tr:nth-child(3) th :nth-child(1) {
          background-color: #337ab7 !important;
          color: #fff;
        }
        .performance-scorecard-details .table tbody {
          font-family: 'FaktPro-Normal';
          font-size: 0.9rem;
        }
        .performance-scorecard-details .table tbody td {
          padding: 2px 10px;
          margin: 0px;
          border: 1px solid #cfd6de;
        }
        .performance-scorecard-details .table tbody tr {
          border-bottom: 1px solid #37a6f8;
          border-left: 1px solid #37a6f8;
        }
        .performance-scorecard-details .table tbody td:first-child {
          text-align: start;
          box-shadow: inset -1px 0 0 0 #cfd6de;
        }
        .performance-scorecard-details .table tbody .sub-title {
          font-size: 0.8rem !important;
        }
        .performance-scorecard-details .table tbody .sub-subtitle {
          font-size: 0.7rem !important;
        }
        .performance-scorecard-details .table tbody .alignright {
          text-align: right;
          padding-right: 1rem;
        }
        .performance-scorecard-details .table tbody .alignleft {
          text-align: left;
          padding-left: 1rem;
        }
        .performance-scorecard-details .table tbody .text-bold {
          font-family: 'FaktPro-Bold';
        }
        .performance-scorecard-details .table tbody .darkred-bg {
          background-color: #282828 !important;
          color: #fff;
        }
        .performance-scorecard-details .table tbody .lightblue-bg {
          background-color: #646e7a !important;
          color: #fff;
        }
        .performance-scorecard-details .table tbody .gold-bg {
          background-color: #ffa51a;
          color: #fff;
        }
    </style>
  </head>
  <body id='content'>
        ${printContents}
        </body>
          </html>`);
    doc.close();

    const div = doc.getElementById('content');
    const options = {
      logging: true,
      allowTaint: false,
      useCORS: true,
    };
    if (!div) {
      console.error('Element not found');
      return;
    }
    html2canvas(div, options)
      .then((canvas) => {
        let imgWidth = 285;
        let pageHeight = 227.2;
        let imgHeight = (canvas.height * imgWidth) / canvas.width;
        let heightLeft = imgHeight;
        const contentDataURL = canvas.toDataURL('image/png');
        let pdfData = new jsPDF('l', 'mm', 'a4', true);
        let position = 5;

        function addExtraDataToPage(pdf: any, extraData: any, positionY: any) {
          pdf.text(extraData, 10, positionY);
        }

        function addPageAndImage(pdf: any, contentDataURL: any, position: any) {
          pdf.addPage();
          pdf.addImage(
            contentDataURL,
            'PNG',
            5,
            position,
            imgWidth,
            imgHeight,
            undefined,
            'FAST'
          );
        }

        pdfData.addImage(
          contentDataURL,
          'PNG',
          5,
          position,
          imgWidth,
          imgHeight,
          undefined,
          'FAST'
        );
        addExtraDataToPage(pdfData, '', position + imgHeight + 10);
        heightLeft -= pageHeight;

        while (heightLeft >= 0) {
          position = heightLeft - imgHeight;
          addPageAndImage(pdfData, contentDataURL, position);
          addExtraDataToPage(pdfData, '', position + imgHeight + 10);
          heightLeft -= pageHeight;
        }

        return pdfData;
      })
      .then((doc) => {
        doc.save('Net Adds / Deducts Details.pdf');
        // popupWin.close();
        this.spinner.hide();
      });
  }

  sendEmailDataDetails(Email: any, notes: any, from: any) {
    this.spinner.show();
    const printContents = document.getElementById(
      'ExpenseTrendDetailsDownload'
    )!.innerHTML;
    const iframe = document.createElement('iframe');

    // Make the iframe invisible
    iframe.style.position = 'absolute';
    iframe.style.width = '0px';
    iframe.style.height = '0px';
    iframe.style.border = 'none';

    document.body.appendChild(iframe);

    const doc = iframe.contentWindow?.document;
    if (!doc) {
      console.error('Failed to create iframe document');
      return;
    }
    doc.open();
    doc.write(`
        <html>
            <head>
            <title>Net Adds / Deducts</title>
                 <style>
                 @font-face {
                  font-family: 'GothamBookRegular';
                  src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                @font-face {
                  font-family: 'Roboto';
                  src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                @font-face {
                  font-family: 'RobotoBold';
                  src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                                            .comment {
                        width: fit-content;
                        background-color: #d7d7d7;
                        border-radius: 0 15px 0 15px;
                        padding: .2rem .5rem;
                        margin-left: 2rem;
                      }
                      .bdr{
                        border:1px solid white !important
                      }
                                            .notes { 
                        width: fit-content;  
                        padding: 0.2rem 0.5rem;
                        margin-right: 3rem;
                        float: inline-end;
                      }
                              .performance-scorecard-details {
                        display: flex;
                        flex-direction: column;
                        height: auto;
                        /* Adjust based on your needs */
                        width: 100%;
                      }
                      .performance-scorecard-details .table > :not(:first-child) {
                        border-top: 0px solid #ffa51a;
                      }
                      .performance-scorecard-details .table {
                        text-align: center;
                        text-transform: capitalize;
                        border: transparent;
                        width: 100%;
                      }
                      .performance-scorecard-details .table th, .performance-scorecard-details .table td {
                        white-space: nowrap;
                        vertical-align: top;
                      }
                      .performance-scorecard-details .table th:first-child, .performance-scorecard-details .table td:first-child {
                        left: 0;
                        z-index: 1;
                      }
                      .performance-scorecard-details .table tr:nth-child(odd) td:first-child, .performance-scorecard-details .table tr:nth-child(odd) td:nth-child(2) {
                        background-color: #e9ecef;
                      }
                      .performance-scorecard-details .table tr:nth-child(even) td:first-child, .performance-scorecard-details .table tr:nth-child(even) td:nth-child(2) {
                        background-color: #fff;
                      }
                      .performance-scorecard-details .table tr:nth-child(odd) {
                        background-color: #e9ecef;
                      }
                      .performance-scorecard-details .table tr:nth-child(even) {
                        background-color: #fff;
                      }
                      .performance-scorecard-details .table .spacer {
                        background-color: #cfd6de !important;
                        border-left: 1px solid #cfd6de !important;
                        border-bottom: 1px solid #cfd6de !important;
                        border-top: 1px solid #cfd6de !important;
                      }
                      .performance-scorecard-details .table .hidden {
                        display: none !important;
                      }
                      .performance-scorecard-details .table .bdr-rt {
                        border-right: 1px solid #abd0ec;
                      }
                      .performance-scorecard-details .table thead {
                        position: sticky;
                        top: 0;
                        z-index: 99;
                        font-family: 'FaktPro-Bold';
                        font-size: 0.8rem;
                      }
                      .performance-scorecard-details .table thead th {
                        padding: 5px 10px;
                        margin: 0px;
                        border-right: 1px solid #abd0ec;
                      }
                      .performance-scorecard-details .table thead .bdr-btm {
                        border-bottom: #005fa3;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(1) {
                        background-color: #fff !important;
                        color: #000;
                        text-transform: uppercase;
                        border-bottom: #cfd6de;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(1) {
                        background-color: #fff !important;
                        color: #000;
                        text-transform: uppercase;
                        border-bottom: #cfd6de;
                        box-shadow: inset 0 1px 0 0 #cfd6de;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(1) th {
                        box-shadow: inset 0 -2px 0 #337ab7;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(3) {
                        background-color: #fff !important;
                        color: #000;
                        text-transform: uppercase;
                        border-bottom: #cfd6de;
                        box-shadow: inset 0 1px 0 0 #cfd6de;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(3) th :nth-child(1) {
                        background-color: #337ab7 !important;
                        color: #fff;
                      }
                      .performance-scorecard-details .table tbody {
                        font-family: 'FaktPro-Normal';
                        font-size: 0.9rem;
                      }
                      .performance-scorecard-details .table tbody td {
                        padding: 2px 10px;
                        margin: 0px;
                        border: 1px solid #cfd6de;
                      }
                      .performance-scorecard-details .table tbody tr {
                        border-bottom: 1px solid #37a6f8;
                        border-left: 1px solid #37a6f8;
                      }
                      .performance-scorecard-details .table tbody td:first-child {
                        text-align: start;
                        box-shadow: inset -1px 0 0 0 #cfd6de;
                      }
                      .performance-scorecard-details .table tbody .sub-title {
                        font-size: 0.8rem !important;
                      }
                      .performance-scorecard-details .table tbody .sub-subtitle {
                        font-size: 0.7rem !important;
                      }
                      .performance-scorecard-details .table tbody .alignright {
                        text-align: right;
                        padding-right: 1rem;
                      }
                      .performance-scorecard-details .table tbody .alignleft {
                        text-align: left;
                        padding-left: 1rem;
                      }
                      .performance-scorecard-details .table tbody .text-bold {
                        font-family: 'FaktPro-Bold';
                      }
                      .performance-scorecard-details .table tbody .darkred-bg {
                        background-color: #282828 !important;
                        color: #fff;
                      }
                      .performance-scorecard-details .table tbody .lightblue-bg {
                        background-color: #646e7a !important;
                        color: #fff;
                      }
                      .performance-scorecard-details .table tbody .gold-bg {
                        background-color: #ffa51a;
                        color: #fff;
                      }
                 </style>
            </head>
            <body id='content'>
                ${printContents}
            </body>
        </html>
    `);

    doc.close();

    const div = doc.getElementById('content');
    if (!div) {
      console.error('Element not found');
      return;
    }

    const options = {
      logging: true,
      allowTaint: false,
      useCORS: true,
      scale: 1, // Adjust scale to fit the page better
    };

    html2canvas(div, options)
      .then((canvas) => {
        let imgWidth = 285;
        let pageHeight = 204;
        let imgHeight = (canvas.height * imgWidth) / canvas.width;
        let heightLeft = imgHeight;
        const contentDataURL = canvas.toDataURL('image/png');
        let pdfData = new jsPDF('l', 'mm', 'a4', true);
        let position = 5;
        pdfData.addImage(
          contentDataURL,
          'PNG',
          5,
          position,
          imgWidth,
          imgHeight,
          undefined,
          'FAST'
        );
        heightLeft -= pageHeight;
        while (heightLeft >= 0) {
          position = heightLeft - imgHeight;
          pdfData.addPage();
          pdfData.addImage(
            contentDataURL,
            'PNG',
            5,
            position,
            imgWidth,
            imgHeight,
            undefined,
            'FAST'
          );
          heightLeft -= pageHeight;
        }

        const pdfBlob = pdfData.output('blob');
        const pdfFile = this.blobToFile(pdfBlob, 'Net Adds / Deducts Details.pdf');
        const formData = new FormData();
        formData.append('to_email', Email);
        formData.append('subject', 'Net Adds / Deducts Details');
        formData.append('file', pdfFile);
        formData.append('notes', notes);
        formData.append('from', from);
        this.apiSrvc
          .postmethod(this.comm.routeEndpoint + 'mail', formData)
          .subscribe(
            (res: any) => {
              console.log('Response:', res);
              if (res.status === 200) {

                this.toast.show(res.response, 'success', 'Success');
              } else {

                this.toast.show('Invalid Details.', 'danger', 'Error');
              }
            },
            (error) => {
              console.error('Error:', error);
            }
          );
      })
      .catch((error) => {
        console.error('html2canvas error:', error);
      })
      .finally(() => {
        this.spinner.hide();
        // popupWin.close();
      });
  }
}
