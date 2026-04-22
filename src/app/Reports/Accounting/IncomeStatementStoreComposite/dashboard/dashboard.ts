import { Component, HostListener, Injector } from '@angular/core';
import { Api } from '../../../../Core/Providers/Api/api';
import { common } from '../../../../common';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { FormsModule } from '@angular/forms';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { NgxSpinnerService } from 'ngx-spinner';
import { CurrencyPipe, DatePipe, formatDate } from '@angular/common';
import { Title } from '@angular/platform-browser';
import { Subscription } from 'rxjs';
import { environment } from '../../../../../environments/environment';
import * as ExcelJS from 'exceljs';
import FileSaver from 'file-saver';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
import numeral from 'numeral';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { IncomestatementstoreDetails } from '../incomestatementstore-details/incomestatementstore-details';
import { ToastContainer } from '../../../../Layout/toast-container/toast-container';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import autoTable from 'jspdf-autotable';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, ToastContainer, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  Current_Date: any;
  gridshow: boolean = false;
  monthgridshow: boolean = false;
  FSData: any = [];
  StoreNamesHeadings: any = [];
  MonthsHeadings: any = [];
  NoData: boolean = false;
  ExpenseTrendByStoreKeys: any = [];
  AllDatakeys: any = [];
  ExpenseTrendByStore_Excel: any;
  XpenseTrendByStoreKeys: string[] = [];
  ExpenseTrendByStore: any;
  ExpenseTrendByStore_ExcelMonth: any;
  XpenseTrendByStoreKeysMonth: string[] = [];
  ExpenseTrendByStoreKeysMonth: any;
  AllDatakeysMonth: any;
  ExpenseTrendByStoreMonth: any;
  Filter: any = 'StoreSummary';
  StoreName: any;
  SubFilter: any;
  SelectedTab: any = [];
  SubSelectedTab1: any = [];
  Month: any;
  selectedstorevalues: any;
  selectedstorename: any;
  StoreValues: any = '0';
  popup: any = [{ type: 'Popup' }];
  groups: any = 0;
  gridvisibility: any;
  newdate: Date = new Date();
  ShowHideBudget: any = 'Hide';
  date: any = new Date();
  header: any = [{ type: 'Bar', storeIds: this.StoreValues, Month: this.date, groups: this.groups, ShowHideBudget: this.ShowHideBudget }]
  selectedDate: Date = new Date();
  currentMonth!: Date;

  activePopover: number = -1;
  storeIds: any = '0';
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  stores: any = [];
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
    private toast: ToastService,
    private comm: common,
    private injector: Injector,
    public shared: Sharedservice,
  ) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
    }
    console.log('store displayname', this.storedisplayname);

    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }
    const lastMonth = new Date();
    let today = new Date();
    if (today.getDate() < 7) {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth() - 1));
    } else {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth()));
    }

    this.Month =
      ('0' + (this.date.getMonth() + 1)).slice(-2) +
      '-' +
      ('0' + this.date.getDate()).slice(-2) +
      '-' +
      this.date.getFullYear();
    this.title.setTitle(this.comm.titleName + '-Income Statement Store');
    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Income Statement Store',
        path1: '',
        path2: '',
        path3: '',
        Month: this.date,
        stores: this.StoreValues.toString(),
        store: this.storeIds,
        groups: this.groups,
        count: 0,
      };
      this.apiSrvc.SetHeaderData({
        obj: data,
      });
      let date = ('0' + (this.date.getMonth() + 1)).slice(-2) +
        '-01' +
        '-' +
        this.date.getFullYear();

      this.header = [{
        type: 'Bar', storeIds: this.StoreValues, Month: this.date, groups: this.groups,
        ShowHideBudget: this.ShowHideBudget,
      }]
      this.selectedDate = this.date
      this.currentMonth = this.selectedDate;
      this.GetData(this.currentMonth);
    } else {
    }
    window.addEventListener('scroll', function () {
      const maxHeight = document.body.scrollHeight - window.innerHeight;
    });
  }
  ngOnInit(): void {
    this.getStores();

  }
  Scrollpercent: any = 0;
  updateVerticalScroll(event: any): void {
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );

  }
  sanitizeValue(value: any): string {
    if (value == null) return '-';
    return value.toString().replace(/-/g, '');
  }

  formatKey(key: string): string {
    const upperKey = key.toUpperCase();

    if (upperKey === 'BDGT_TOTAL') {
      return 'BUDGET TOTAL';
    } else if (upperKey === 'VARBDGT') {
      return 'BUDGET VARIANCE';
    } else if (key.toLowerCase().startsWith('bdgt_')) {
      return 'BUDGET';
    } else if (key.toLowerCase().startsWith('var')) {
      return 'VARIANCE';
    }

    return key;
  }


  excelformatKey(key: string): string {
    const upperKey = key.toUpperCase();

    if (upperKey === 'BDGT_TOTAL') {
      return 'BUDGET TOTAL';
    } else if (upperKey === 'VARBDGT') {
      return 'BUDGET VARIANCE';
    } else if (key.toLowerCase().startsWith('bdgt_')) {
      return 'BUDGET';
    } else if (key.toLowerCase().startsWith('var')) {
      return 'VARIANCE';
    }

    return key;
  }
  isKeyVisible(key: string): boolean {
    return key !== '-1';
  }

  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month',
    maxDate: new Date()
  };
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }
  storesWidth: string = '100%';

  applyDateChange() {
    if (!this.storeIds || this.storeIds.length === 0) {

      this.toast.show(
        'Please Select Atleast One Store',
        'warning',
        'Warning'
      );
      return;
    }
    else {
      this.currentMonth = this.selectedDate;

      // set width based on storeIds length
      this.storesWidth = this.storeIds?.toString().length > 3 ? '100%' : '60%';

      this.GetData(this.currentMonth);
    }
  }

  GetData(date: any) {
    ////console.log(this.Month)
    this.spinner.show();
    this.ExpenseTrendByStoreKeys = [];
    this.ExpenseTrendByStore = [];
    this.AllDatakeys = [];
    const DateToday = this.datepipe.transform(
      new Date(date),
      'dd-MMM-yyyy'
    );
    const obj = {
      SalesDate: this.shared.datePipe.transform(new Date(date),'yyyy-MM')+'-10',
      Stores: this.storeIds.toString(),
      UserID: 0,
    };
    let startFrom = new Date().getTime();
    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetIncomeStatementStoreComposite', obj)
      .subscribe(
        (x) => {
          const currentTitle = document.title;
          if (x.status == 200) {
            if (x.response != undefined) {
              if (x.response.length > 0) {
                let resTime = (new Date().getTime() - startFrom) / 1000;
                this.gridshow = true;
                this.XpenseTrendByStoreKeys = Object.keys(x.response[0]);
                let ETByStoreKeys_sorted = this.XpenseTrendByStoreKeys.filter((x) => {
                  if (x === 'SNo' || x === 'NgClass') return false;

                  if (this.ShowHideBudget === 'Show') {
                    return true; // include all keys
                  }
                  const isDynamicBudget = /^BDGT_\d+$/.test(x);
                  const isDynamicVariance = /^VAR\d+$/.test(x);
                  const isBudgetTotal = x === 'BDGT_TOTAL';
                  const isBudgetVariance = x === 'VARBDGT';
                  return !isDynamicBudget && !isDynamicVariance && !isBudgetTotal && !isBudgetVariance;
                });
                console.log('Filtered keys:', ETByStoreKeys_sorted);
                console.log('Show/Hide Budget setting:', this.ShowHideBudget);
                console.log('ETByStoreKeys_sorted', ETByStoreKeys_sorted)


                let AllStore_Label = ETByStoreKeys_sorted.find((x: any) => {
                  if (x.toUpperCase() == 'All Stores'.toUpperCase()) return x;
                });
                console.log('AllStore_Label', AllStore_Label)
                const lable = ETByStoreKeys_sorted.findIndex(
                  (i) => i == 'LABLE'
                );
                ETByStoreKeys_sorted.splice(lable, 1);
                ETByStoreKeys_sorted = ETByStoreKeys_sorted.splice(1);
                this.StoreNamesHeadings = ETByStoreKeys_sorted;
                const index = ETByStoreKeys_sorted.findIndex(
                  (i) => i == 'Report Total'
                );
                if (index >= 0) {
                  ETByStoreKeys_sorted.splice(index, 1);
                  ETByStoreKeys_sorted.splice(0, 0, 'Report Total');
                }

                if (this.ShowHideBudget === 'Show') {
                  // Insert 'BDGT_TOTAL' and 'VARBDGT' after 'Report Total'
                  const reportTotalIndex = ETByStoreKeys_sorted.indexOf('Report Total');
                  if (reportTotalIndex >= 0) {  // Remove BDGT_TOTAL and VARBDGT if already in list
                    ETByStoreKeys_sorted = ETByStoreKeys_sorted.filter(
                      (x) => x !== 'BDGT_TOTAL' && x !== 'VARBDGT'
                    );

                    // Insert them in order
                    ETByStoreKeys_sorted.splice(reportTotalIndex + 1, 0, 'BDGT_TOTAL', 'VARBDGT');
                  }
                }
                const Cmtindex = ETByStoreKeys_sorted.findIndex(
                  (i) => i == 'CommentsStatus'
                );
                if (Cmtindex >= 0) {
                  ETByStoreKeys_sorted.splice(Cmtindex, 1);
                  // ETByStoreKeys_sorted.splice(0, 0, 'CommentsStatus');
                }
                for (var i = 0; i < ETByStoreKeys_sorted.length; i++) {
                  this.ExpenseTrendByStoreKeys.push(
                    ETByStoreKeys_sorted[i].toLowerCase()
                  );
                  this.AllDatakeys.push(ETByStoreKeys_sorted[i]);
                }
                this.ExpenseTrendByStoreKeys.push(AllStore_Label);
                this.AllDatakeys.push(AllStore_Label);
                let XpenseTrendByStoreData = x.response;
                let ETByStoreData_sorted = XpenseTrendByStoreData;
                let AllStoreLabel_Data = '';
                this.ExpenseTrendByStore = x.response;
                if (this.ExpenseTrendByStoreKeys.length == 0) {
                  this.NoData = true;
                } else {
                  this.NoData = false;
                }
                this.spinner.hide();
              } else {
                this.spinner.hide();
                this.NoData = true;
              }
            } else {
              this.ExpenseTrendByStore = [];
              this.AllDatakeys = [];
              this.spinner.hide();
              this.NoData = true;
            }
          } else {
            let resTime = (new Date().getTime() - startFrom) / 1000;
            this.toast.show(
              x.status,
              'danger',
              'Error'
            );
            this.spinner.hide();
            this.NoData = true;
          }
        },
        (error) => {
          this.toast.show(
            '502 Bad Gate Way Error',
            'danger',
            'Error'
          );
          this.spinner.hide();
          this.NoData = true;
        }
      );
    this.apiSrvc.GetHeaderData().subscribe((res) => {
    });
  }
  getStores() {
    this.selectedstorevalues = [];
  }
  openDetails(Object: any, storename: any, ref: any, date: any) {
    console.log(Object, storename);
    const storeId =
      this.shared.common.groupsandstores
        .find((g: any) => g.sg_id == this.groupId)
        ?.Stores
        ?.find((s: any) =>
          s.storename.toLowerCase() === storename.toLowerCase()
        )
        ?.ID ?? 0;
    console.log('store Id', storeId);
    if (ref == 'store') {
      const dateMonth =
        date.toString().substr(8, 2) +
        '-' +
        date.toString().substr(4, 3) +
        '-' +
        date.toString().substr(11, 4);
      const DetailsSF = this.ngbmodal.open(IncomestatementstoreDetails, {
        size: 'xl',
        backdrop: 'static',
        injector: Injector.create({
          providers: [
            { provide: CurrencyPipe, useClass: CurrencyPipe }
          ],
          parent: this.injector
        })
      });


      DetailsSF.componentInstance.Fsdetails = {
        TYPE: Object.LABLEVAL,
        NAME: Object.LABLE,
        STORES: storeId,
        LatestDate: dateMonth,
        Group: this.groups,
        STORENAMES: this.getSelectedStoreNames()
      };
    }
  }
  capitalizeWords(value: string): string {
    return value
      .toLowerCase()
      .replace(/\b\w/g, char => char.toUpperCase());
  }

  selBlock: any;
  commentopen(item: any, i: any, slblock: any = '') {
    this.index = '';
    //console.log('Selected Obj :', item);
    //return
    this.selBlock = slblock + i.toString();
    this.index = i.toString();
    this.commentobj = {
      TYPE: item.LABLE,
      NAME: item.LABLE,
      STORES: i,
      STORENAME: item.LABLE,
      Month: '',
      ModuleId: '73',
      ModuleRef: 'ISSC',
      state: 1,
      indexval: i,
    };
    // const DetailsSF = this.ngbmodal.open(CommentsComponent, {
    //   size: 'xl',
    //   backdrop: 'static',
    // });
    // DetailsSF.componentInstance.SFComments = {
    //   TYPE: item.LABLEVAL,
    //   NAME: item.LABLE,
    //   STORES: this.selectedstorevalues,
    //   LatestDate: this.Month,
    //   STORENAME: this.selectedstorename,
    //   Month: this.Month,
    //   ModuleId: '66',
    //   ModuleRef: 'SF',
    // };
    // DetailsSF.result.then(
    //   (data) => {},
    //   (reason) => {

    //     // // on dismiss
    //     // const Data = {
    //     //   state: true,
    //     // };
    //     // this.apiSrvc.setBackgroundstate({ obj: Data });
    //     this.GetData();
    //   }
    // );

  }
  index = '';
  commentobj: any = {};
  addcmt(data: any) {
    // //console.log('Checking Add cmt  : ', data);
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
            // //console.log(data);
          },
          (reason) => {
            // //console.log(reason);
            if (reason == 'O') {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
            } else {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
              if (this.Filter == 'StoreSummary') {
                this.GetData(this.currentMonth);
              }
            }
            // // on dismiss
            // const Data = {
            //   state: true,
            // };
            // this.apiSrvc.setBackgroundstate({ obj: Data });
          }
        );
    }
    if (data == 'AD') {
      if (this.Filter == 'StoreSummary') {
        this.GetData(this.currentMonth);
      }
    }
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  SFstate: any;
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit(): void {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Income Statement Store') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.apiSrvc.GetReportOpening().subscribe((res) => {
      // //console.log(res);
      // if (res.obj.Module == 'Income Statement Store') {
      //   const modalRef = this.ngbmodal.open(StoreReportComponent, {
      //     size: 'xl',
      //     backdrop: 'static',
      //   });
      // }
      if (res.obj.Module == 'Income Statement Store') {
        document.getElementById('report')?.click()
      }
    });
    this.apiSrvc.getExportToExcelAllReports().subscribe((res) => {
      this.SFstate = res.obj.state;
      if (res.obj.title == 'Income Statement Store') {
        if (res.obj.state == true) {
          this.exportAsXLSX();
        }
      }
    });
    this.apiSrvc.GetReports().subscribe((data) => {
      //console.log(data)
      if (data.obj.Reference == 'Income Statement Store') {
        if (data.obj.header == undefined) {
          this.date = data.obj.month;
          this.Month = data.obj.month;
          // this.Month =  ('0' +( new Date(data.obj.month).getMonth())).slice(-2)+'-01'+'-'+new Date (data.obj.month).getFullYear()
          this.StoreValues = data.obj.storeValues;
          this.StoreName = data.obj.Sname;
          this.Filter = data.obj.filters;
          this.SubFilter = data.obj.subfilters;
          this.ShowHideBudget = data.obj.ShowHideBudget

          this.groups = data.obj.groups;
          if (this.StoreValues.toString().indexOf(',') > 0) {
            this.gridvisibility = 'DL';
          } else {
            this.gridvisibility = 'SL';
          }
          // this.DataSelection(this.Filter);
        } else {
          if (data.obj.header == 'Yes') {
            this.StoreValues = data.obj.storeValues;
            //console.log(this.StoreValues);
          }
        }
        const headerdata = {
          title: 'Income Statement Store',
          path1: '',
          path2: '',
          path3: '',
          Month: this.Month,
          filter: this.Filter,
          stores: this.storeIds,
          ShowHideBudget: this.ShowHideBudget,

          groups: this.groups,
        };
        this.apiSrvc.SetHeaderData({
          obj: headerdata,
        });
        // let date=    ('0' + ( this.Month.getMonth() + 1)).slice(-2) +
        // '-01' +
        // '-' +
        // this.Month.getFullYear();
        this.header = [{ type: 'Bar', storeIds: this.StoreValues, Month: this.Month, groups: this.groups, ShowHideBudget: this.ShowHideBudget, }]
        if (this.StoreValues != '') {
          this.GetData(this.currentMonth);
        } else {
          this.NoData = true;
          this.ExpenseTrendByStore = [];
          this.ExpenseTrendByStoreKeys = []
        }
      }
    });
    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Income Statement Store') return;
      if (obj.state) {
        this.exportAsXLSX();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Income Statement Store') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Income Statement Store') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Income Statement Store') return;
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
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
  }

  reportOpen(temp: any) {
    this.ngbmodalActive = this.ngbmodal.open(temp, {
      size: 'xl',
      backdrop: 'static',
    });
  }
  LogCount = 1;

  close(data: any) {
    this.index = '';
  }

  isUnitRow(label: string): boolean {
    return ['New Units', 'Pre-Owned Units', 'Unit Retail Sales', 'Fleet Units', 'Variable Expenses%',
      'Personnel Expenses%', 'Semi-Fixed Expenses%', 'Fixed Expenses%', 'Other Expenses%'].includes(label);
  }
  isNonClickable(value: any, itemKey: string, label: string): boolean {
    return value === null || value === undefined || value === '' ||
      itemKey === 'Report Total' || this.isUnitRow(label);
  }
  isBoldRow(label: string): boolean {
    const boldLabels = [
      'Unit Retail Sales', 'Pure Gross', 'Variable Gross', 'Total Fixed',
      'Total Store Gross', 'Total Expenses', 'Operating Profit',
      'Net Adds/Deducts', 'Net Profit'
    ];
    return boldLabels.includes(label);
  }

  getCellClasses(item: any, itemKey: string) {
    const isNegative = !this.inTheGreen(item[itemKey]);
    const isBold = this.isBoldRow(item.LABLE);
    const isNonClickable = this.isNonClickableKey(itemKey);

    return {
      negative: isNegative && ['Operating Profit', 'Net Profit'].includes(item.LABLE),
      setPerens: isNegative,
      bold: isBold,
      notbold: !isBold,
      nopointerevents: isNonClickable,
      pointerevents: !isNonClickable
    };
  }

  isNonClickableKey(key: string): boolean {
    return key.startsWith('BDGT_') || key.startsWith('VAR') || key === 'Report Total';
  }
  generatePDF() {
    if (this.storeIds.length >= 16) {

      this.toast.show(
        'Please select only up to 15 stores',
        'warning',
        'Warning'
      );
      return;
    }
    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save('Income Statement Store.pdf'); // ✅ only here
    } catch (e) {
      console.error(e);
      this.toast.show('Error generating PDF', 'danger', 'Error');
    } finally {
      this.shared.spinner.hide();
    }
  }
  printPDF() {
    if (this.storeIds.length >= 16) {
      this.toast.show('Please select only up to 15 stores', 'warning', 'Warning');
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

  private createPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a2' });

    /* ================= TITLE ================= */
    doc.setFontSize(14);
    doc.text('Income Statement Store', 14, 12);

    let startY = 20;

    /* ================= SAFETY ================= */
    if (!this.StoreNamesHeadings || this.StoreNamesHeadings.length === 0) {
      console.error('StoreNamesHeadings is empty');
      return doc;
    }

    if (!this.ExpenseTrendByStore || this.ExpenseTrendByStore.length === 0) {
      console.error('No data for PDF');
      return doc;
    }

    const monthDate = this.datepipe.transform(this.Month, 'MMMM yyyy');

    /* ================= HEADER ================= */
    const head = [
      [
        monthDate,
        ...this.StoreNamesHeadings.map((h: string) => this.excelformatKey(h))
      ]
    ];

    /* ================= BODY ================= */
    const body: any[] = [];

    this.ExpenseTrendByStore.forEach((d: any) => {

      let row: any[] = [d.LABLE];

      this.StoreNamesHeadings.forEach((key: string) => {
        let val = d[key];
        let num = Number(val);

        if (val === '' || val == null || num === 0) {
          row.push('-');
        } else {
          row.push(num);
        }
      });

      body.push(row);

      /* ===== SUB ROWS (same as Excel) ===== */
      if (Array.isArray(d.SubData)) {
        d.SubData.forEach((sub: any) => {

          let subRow: any[] = [sub.LABLE || ''];

          this.StoreNamesHeadings.forEach((key: string) => {
            let val = sub[key];
            let num = Number(val);

            if (val === '' || val == null || num === 0) {
              subRow.push('-');
            } else {
              subRow.push(num);
            }
          });

          body.push(subRow);
        });
      }

    });

    /* ================= TABLE ================= */
    autoTable(doc, {
      startY: startY,
      head: head,
      body: body,
      theme: 'grid',

      styles: {
        fontSize: 10,
        cellPadding: 2,
        halign: 'right',
        valign: 'middle',
        lineWidth: 0.1,
        lineColor: [200, 200, 200],
        textColor: [20, 20, 20], // 👈 slightly richer black
      },
      /* ✅ COLUMN WIDTH CONTROL */
      columnStyles: {
        0: { cellWidth: 50 }, // 1st column (Label)
        1: { cellWidth: 30 }, // 2nd column
        // remaining columns → auto (no need to define)
      },
      didParseCell: (data: any) => {

        /* ===== HEADER ===== */
        if (data.section === 'head') {
          data.cell.styles.fillColor = [69, 132, 255];
          data.cell.styles.textColor = [255, 255, 255];
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.halign = 'center';
          return;
        }

        const rowIndex = data.row.index;
        const rowData = this.ExpenseTrendByStore[rowIndex];

        if (!rowData) return;

        const label = rowData.LABLE;

        /* ===== ALIGN ===== */
        data.cell.styles.halign =
          data.column.index === 0 ? 'left' : 'right';

        /* ===== EMPTY ===== */
        if (
          data.cell.raw === null ||
          data.cell.raw === undefined ||
          data.cell.raw === '' ||
          data.cell.raw === 0
        ) {
          data.cell.text = ['-'];
          data.cell.styles.halign = 'center';
          return;
        }

        /* ===== FORMAT ===== */
        if (data.column.index > 0 && data.cell.raw !== '-') {

          const val = Number(data.cell.raw);

          if (!isNaN(val)) {

            const isPercent = label.includes('%');
            const isNumber = label.includes('Units');
            const isSpecial = this.isSpecialRow(label);

            if (isPercent) {
              const rounded = Math.round(val); // ✅ % fix
              data.cell.text = [`${rounded}%`];
              data.cell.styles.halign = 'center';
            }
            else if (isNumber) {
              data.cell.text = [val.toLocaleString()];
            }
            else {
              data.cell.text = [`$ ${Math.abs(val).toLocaleString()}`];

              if (val < 0) {
                data.cell.styles.textColor = isSpecial
                  ? [255, 0, 0]
                  : [220, 53, 69];
              }
            }
          }
        }

        /* ===== BOLD ===== */
        if (this.isBoldRow(label)) {
          data.cell.styles.fontStyle = 'bold';
        }

        /* ===== ALT ROW COLOR ===== */
        if (rowIndex % 2 === 0) {
          data.cell.styles.fillColor = [245, 247, 250];
        }
      }
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
      title: 'Income Statement Store',
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
        }
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
  isSpecialRow(name: string): boolean {
    return [
      'Total Selling Gross (New)',
      'Total Selling Gross (Used)',
      'Net Income Before Taxes',
      'Total Operating Department Profit'
    ].includes(name);
  }
  exportAsXLSX(): void {

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Income Statement');

    /* ================= 1. FILTERS ================= */
    const filterRowCount = this.addExcelFiltersSection(worksheet);

    /* ================= 2. HEADER ================= */
    const headerStartRow = filterRowCount + 1;
    const monthDate = this.datepipe.transform(this.Month, 'MMMM yyyy');

    let storeHeadings = [...this.StoreNamesHeadings];

    // ✅ SINGLE HEADER ROW (No Category row)
    const headerRow = worksheet.addRow([
      monthDate,
      ...storeHeadings.map((h: any) => this.excelformatKey(h))
    ]);

    // Merge first cell for title spacing
    worksheet.mergeCells(`A${headerStartRow}:A${headerStartRow}`);

    /* ================= HEADER STYLE ================= */
    worksheet.getRow(headerStartRow).eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0554EF' }
      };
      cell.font = {
        bold: true,
        color: { argb: 'FFFFFF' },
        name: 'Calibri'
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    /* ================= FORMAT FUNCTION ================= */
    const formatRow = (
      row: any,
      valueTypes: any[] = [],
      isSpecial = false,
      isTitleRow = false
    ) => {

      row.eachCell((cell: any, colNumber: number) => {

        const isBold = isTitleRow || isSpecial;

        // ✅ ADD BORDER TO ALL CELLS
        cell.border = {
          top: { style: 'thin', color: { argb: 'D3D3D3' } },
          bottom: { style: 'thin', color: { argb: 'D3D3D3' } },
          left: { style: 'thin', color: { argb: 'D3D3D3' } },
          right: { style: 'thin', color: { argb: 'D3D3D3' } }
        };

        // First column
        if (colNumber === 1) {
          cell.alignment = {
            horizontal: 'left',
            vertical: 'middle',
            indent: isTitleRow ? 2 : 3
          };
          cell.font = {
            bold: isBold,
            name: 'Calibri'
          };
          return;
        }

        const valueType = valueTypes[colNumber - 2];

        if (!cell.value || cell.value === '-') {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          return;
        }

        let num = parseFloat(cell.value);

        if (!isNaN(num)) {

          if (valueType === '$') {
            cell.numFmt = '"$" * #,##0; "$" * -#,##0';
            cell.value = num;
          }

          if (valueType === '#') {
            cell.numFmt = '#,##0';
            cell.value = num;
          }

          if (valueType === '%') {
            cell.numFmt = '0.0%';
            // cell.value = num / 100;
          }

          if (isSpecial && num < 0) {
            cell.font = {
              color: { argb: 'FF0000' },
              bold: true
            };
          } else {
            cell.font = {
              bold: isBold,
              name: 'Calibri'
            };
          }
        }

        cell.alignment = valueType === '%'
          ? { horizontal: 'center', vertical: 'middle' }
          : { horizontal: 'right', vertical: 'middle' };
      });
    };
    /* ================= DATA ================= */
    const data = this.ExpenseTrendByStore.map((x: any) => ({ ...x }));

    data.forEach((d: any) => {

      const isTitleRow = this.isBoldRow(d.LABLE);
      const isSpecial = this.isSpecialRow ? this.isSpecialRow(d.LABLE) : false;

      let rowData: any[] = [d.LABLE];

      storeHeadings.forEach((key: any) => {
        let val = d[key];
        rowData.push(val === '' || val == null ? '-' : val);
      });

      const row = worksheet.addRow(rowData);

      const valueType = d.LABLE.includes('%')
        ? '%'
        : d.LABLE.includes('Units')
          ? '#'
          : '$';

      formatRow(
        row,
        Array(storeHeadings.length).fill(valueType),
        isSpecial,
        isTitleRow
      );

      /* ===== OPTIONAL SUB ROW SUPPORT ===== */
      if (Array.isArray(d.SubData)) {
        d.SubData.forEach((sub: any) => {

          let subRowData: any[] = [sub.LABLE || ''];

          storeHeadings.forEach((key: any) => {
            let val = sub[key];
            subRowData.push(val === '' || val == null ? '-' : val);
          });

          const subRow = worksheet.addRow(subRowData);
          subRow.outlineLevel = 1;

          formatRow(
            subRow,
            Array(storeHeadings.length).fill(valueType)
          );
        });
      }

      /* ===== ALT ROW COLOR ===== */
      if (row.number % 2 === 0) {
        row.eachCell(cell => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'F5F7FA' }
          };
        });
      }
    });

    /* ================= FINAL SETTINGS ================= */

    worksheet.views = [
      {
        state: 'frozen',
        xSplit: 1,
        ySplit: headerStartRow
      }
    ];

    worksheet.columns = [
      { width: 40 },
      ...storeHeadings.map(() => ({ width: 28 }))
    ];

    /* ================= EXPORT ================= */
    const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');

    workbook.xlsx.writeBuffer().then((data: BlobPart) => {
      FileSaver.saveAs(
        new Blob([data]),
        'IncomeStatement_' + DATE_EXTENSION + '.xlsx'
      );
    });
  }


  generatePDFBlob(): Blob | null {

    // // ✅ Allow max 15 stores
    // if (this.storeIds?.length > 15) {
    //   this.toast.show('Please select only up to 15 stores', 'warning', 'Warning');
    //   return null;
    // }

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
    if (this.storeIds.length >= 16) {
      this.toast.show('Please select only up to 15 stores', 'warning', 'Warning');
      return;
    }
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
      const pdfFile = this.blobToFile(pdfBlob, 'Income Statement Store.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Income Statement Store');
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
  GetPrintData() {
    window.print();
  }
}