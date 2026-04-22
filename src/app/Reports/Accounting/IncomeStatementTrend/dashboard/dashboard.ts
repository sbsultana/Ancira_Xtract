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
import autoTable from 'jspdf-autotable';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
import numeral from 'numeral';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { ToastContainer } from '../../../../Layout/toast-container/toast-container';
import { IncomestatementtrendDetails } from '../incomestatementtrend-details/incomestatementtrend-details';
import { IncomestatementtrendGraph } from '../incomestatementtrend-graph/incomestatementtrend-graph';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, ToastContainer, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
  providers: [DatePipe]
})
export class Dashboard {
  [x: string]: any;
  Current_Date: any;

  gridshow: boolean = false;
  monthgridshow: boolean = false;
  FSData: any = [];
  StoreNamesHeadings: any = [];
  MonthsHeadings: any = [];

  NoData: boolean = false;
  fromnewdate = new Date();
  date: any = new Date(
    this.fromnewdate.setFullYear(this.fromnewdate.getFullYear())
  );
  ExpenseTrendByStoreKeys: any[] = [];
  AllDatakeys: any[] = [];
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
  SelectedTab: any[] = [];
  SubSelectedTab1: any[] = [];
  Month: any;
  // stores: any;
  selectedstorevalues: any;

  fromdate: any;
  todate: any;
  tonewdate: any = new Date();
  solutionurl: any = environment.apiUrl;
  StoreValues: any = [

  ];
  groups: any = 1;
  selectedstorename: any = 'All Stores';
  header: any = [
    {
      type: 'Bar',
      storeIds: this.StoreValues,
      fromdate: '01' + '-' + '01' + '-' + this.date.getFullYear(),
      todate:
        ('0' + (this.tonewdate.getMonth() + 1)).slice(-2) +
        '-' +
        '01' +
        '-' +
        this.tonewdate.getFullYear(),
      groups: this.groups
    },
  ];
  popup: any = [{ type: 'Popup' }];
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;

  StartDate: Date = new Date(new Date().setFullYear(new Date().getFullYear() - 1));
  EndDate: Date = new Date();
  StartMonth!: Date;
  EndMonth!: Date;
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month',
    maxDate: new Date()
  };
  activePopover: number = -1;
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = '0';
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
    private cp: CurrencyPipe,
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
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }
    this.tonewdate = new Date();

    let today = new Date()
    // this.fromdate = '01' + '-' + '01' + '-' + this.date.getFullYear();

    let enddate = new Date(today.setDate(today.getDate() - 1));
    console.log(enddate)
    if (enddate.getMonth() == 0) {
      this.fromdate = '01-01-' + (enddate.getFullYear() - 1);
      this.todate = '01-01-' + (enddate.getFullYear());
    } else {
      console.log(enddate.getMonth())
      this.fromdate =
        ('0' + (enddate.getMonth() + 1)).slice(-2) +
        '-01' +
        '-' +
        (enddate.getFullYear() - 1);
      console.log(enddate.getMonth(), this.fromdate)


      let today = new Date();
      if (today.getDate() < 7) {
        this.todate =
          ('0' + (enddate.getMonth())).slice(-2) +
          '-' +
          ('0' + enddate.getDate()).slice(-2) +
          '-' +
          enddate.getFullYear();
        console.log(enddate.getMonth(), this.todate)
      } else {
        this.todate =
          ('0' + (enddate.getMonth() + 1)).slice(-2) +
          '-' +
          ('0' + enddate.getDate()).slice(-2) +
          '-' +
          enddate.getFullYear();
      }
      this.EndDate = new Date(this.todate)
    }

    let newdate = new Date();

    this.title.setTitle(this.comm.titleName + '-Income Statement Trend');

    const queryString = window.location.search;
    const urlParams = new URLSearchParams(queryString);

    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Income Statement Trend',
        path1: '',
        path2: '',
        path3: '',
        Month: this.date,
        filter: this.Filter,
        stores: this.storeIds,
        groups: this.groups,
        selectedstorename: this.selectedstorename,
        fromdate:
          this.fromdate,
        todate:
          this.todate,
        count: 0,
      };

      this.apiSrvc.SetHeaderData({
        obj: data,
      });
      this.header = [
        {
          type: 'Bar',
          storeIds: this.StoreValues,
          fromdate:
            this.fromdate,
          todate:
            this.todate,
          groups: this.groups
        },
      ];
      this.StartMonth = this.StartDate;
      this.EndMonth = this.EndDate;
      this.GetDataByMonths(this.StartMonth, this.EndMonth);
    } else {
      this.getFavReports();
    }
  }

  ngOnInit(): void {

    this.getStores();
  }

  Scrollpercent: any = 0;
  updateVerticalScroll(event: any): void {
    const scrollDemo = <HTMLInputElement>document.querySelector('#scrollcent');
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );
  }

  applyDateChange() {
    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show(
        'Please Select Atleast One Store',
        'warning',
        'Warning'
      );
      return;
    }

    if (!this.StartDate || !this.EndDate) {
      this.toast.show(
        'Please Enter Valid Date',
        'warning',
        'Warning'
      );
      return;
    }

    const start = new Date(this.StartDate);
    const end = new Date(this.EndDate);

    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      this.toast.show(
        'Please Enter Valid Date',
        'warning',
        'Warning'
      );
      return;
    }

    if (start > end) {
      this.toast.show(
        'Please Enter Valid Date',
        'warning',
        'Warning'
      );
      return;
    }
    const monthDiff =
      (end.getFullYear() - start.getFullYear()) * 12 +
      (end.getMonth() - start.getMonth());

    if (monthDiff < 3) {
      this.toast.show(
        'Please Select Atleast 3 Months Range',
        'warning',
        'Warning'
      );
      return;
    }

    this.StartMonth = this.StartDate;
    this.EndMonth = this.EndDate;

    this.GetDataByMonths(this.StartMonth, this.EndMonth);
  }


  FromDate: any;
  ToDate: any;
  GetDataByMonths(StartMonth: any, EndMonth: any) {
    this.spinner.show();
    this.ExpenseTrendByStoreKeysMonth = [];
    this.AllDatakeysMonth = [];
    this.ExpenseTrendByStoreMonth = [];
    this.FromDate = this.datepipe.transform(StartMonth, 'dd-MMM-yyyy');
    this.ToDate = this.datepipe.transform(EndMonth, 'dd-MMM-yyyy');
    const obj = {
      Stores: this.storeIds.toString(),
      SalesDATE: '10-'+this.datepipe.transform(StartMonth, 'MMM-yyyy'),
      EndDate: '10-'+this.datepipe.transform(EndMonth, 'MMM-yyyy'),
      UserID: 0,
    };

    let startFrom = new Date().getTime();
    this.apiSrvc.postmethod(this.comm.routeEndpoint + 'GetIncomeStatementTrend', obj).subscribe(
      (x) => {
        const currentTitle = document.title;
        if (x.status == 200) {
          if (x.response != undefined) {
            if (x.response.length > 0) {
              this.spinner.hide();

              let resTime = (new Date().getTime() - startFrom) / 1000;
              this.monthgridshow = true;

              this.XpenseTrendByStoreKeysMonth = Object.keys(x.response[0]);
              let ETByStoreKeys_sortedMonth =
                this.XpenseTrendByStoreKeysMonth.filter((x) => {
                  return x != 'SNo' && x != 'NgClass';
                });
              let AllStore_LabelMonth = ETByStoreKeys_sortedMonth.find(
                (x: any) => {
                  if (x.toUpperCase() == 'All Stores'.toUpperCase()) return x;
                }
              );
              ETByStoreKeys_sortedMonth = ETByStoreKeys_sortedMonth.splice(2);
              this.MonthsHeadings = ETByStoreKeys_sortedMonth;
              const Cmtindex = ETByStoreKeys_sortedMonth.findIndex(
                (i) => i == 'CommentsStatus'
              );
              if (Cmtindex >= 0) {
                ETByStoreKeys_sortedMonth.splice(Cmtindex, 1);
              }
              for (var i = 0; i < ETByStoreKeys_sortedMonth.length; i++) {
                this.ExpenseTrendByStoreKeysMonth.push(
                  ETByStoreKeys_sortedMonth[i]
                );
                this.AllDatakeysMonth.push(ETByStoreKeys_sortedMonth[i]);
              }
              this.ExpenseTrendByStoreKeysMonth.push(AllStore_LabelMonth);
              this.AllDatakeysMonth.push(AllStore_LabelMonth);
              let XpenseTrendByStoreDataMonth = x.response;
              let ETByStoreData_sorted = XpenseTrendByStoreDataMonth;
              let AllStoreLabel_Data = '';
              this.ExpenseTrendByStoreMonth = x.response;

              if (this.ExpenseTrendByStoreKeysMonth.length == 0) {
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
            this.ExpenseTrendByStoreMonth = [];
            this.AllDatakeysMonth = [];
            this.spinner.hide();
            this.NoData = true;
          }
        } else {
          this.toast.show(
            x.status,
            'danger',
            'Error'
          );

          this.spinner.hide();
          this.NoData = true;
          let resTime = (new Date().getTime() - startFrom) / 1000;
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
  }

  isUnitRow(label: string): boolean {
    return ['New Units', 'Pre-Owned Units', 'Unit Retail Sales', 'Fleet Units', 'Variable Expenses%',
      'Personnel Expenses%', 'Semi-Fixed Expenses%', 'Fixed Expenses%', 'Other Expenses%'].includes(label);
  }
  isNonClickable(value: any, itemKey: string, label: string): boolean {
    return value === null || value === undefined || value === '' ||
      itemKey === 'YTD' || itemKey === 'LYTD' || this.isUnitRow(label);
  }
  isBoldRow(label: string): boolean {
    return [
      'Unit Retail Sales',
      'Pure Gross',
      'Variable Gross',
      'Total Fixed',
      'Total Store Gross',
      'Total Expenses',
      'Operating Profit',
      'Net Adds/Deducts',
      'Net Profit'
    ].includes(label);
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
      if (this.shared.common.pageName == 'Income Statement Trend') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.reportOpenSub = this.apiSrvc.GetReportOpening().subscribe((res) => {
      // //console.log(res);
      if (this.reportOpenSub != undefined) {
        if (res.obj.Module == 'Income Statement Trend') {
          document.getElementById('report')?.click();
        }
      }
    });
    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.excel != undefined) {
        this.SFstate = res.obj.state;
        if (res.obj.title == 'Income Statement Trend') {
          if (res.obj.state == true) {
            this.exportAsXLSX();
          }
        }
      }
    });
    this.reportGetting = this.apiSrvc.GetReports().subscribe((data) => {
      ////console.log(data)
      if (this.reportGetting != undefined) {
        if (data.obj.Reference == 'Income Statement Trend') {
          if (data.obj.header == undefined) {
            this.date = data.obj.month;
            this.Month = data.obj.month;
            // this.Month =  ('0' +( new Date(data.obj.month).getMonth())).slice(-2)+'-01'+'-'+new Date (data.obj.month).getFullYear()
            this.StoreValues = data.obj.storeValues;
            this.selectedstorename = data.obj.selectedstorename;
            this.StoreName = data.obj.Sname;
            this.Filter = data.obj.filters;
            this.SubFilter = data.obj.subfilters;
            this.groups = data.obj.groups;
            let newdate = new Date();
            (this.fromdate = this.datepipe.transform(
              data.obj.FromDate,

              'dd-MMM-yyyy'
            )),
              (this.todate = this.datepipe.transform(
                data.obj.ToDate,

                'dd-MMM-yyyy'
              ));
          } else {
            if (data.obj.header == 'Yes') {
              this.StoreValues = data.obj.storeValues;
              //console.log(this.StoreValues);
              if (this.StoreValues != '') {
                let storename = this.comm.groupsandstores.filter(
                  (v: any) => v.sg_id == this.groups
                )[0].Stores;
                this.selectedstorename = storename.filter(
                  (v: any) => v.ID == this.StoreValues
                )[0].storename;
              }

              //console.log(this.selectedstorename);
            }
          }

          // this.DataSelection(this.Filter);
          const headerdata = {
            title: 'Income Statement Trend',
            path1: '',
            path2: '',
            path3: '',
            Month: this.Month,
            filter: this.Filter,
            stores: this.StoreValues,
            fromdate: this.fromdate,
            todate: this.todate,
            groups: this.groups,
            selectedstorename: this.selectedstorename,
          };
          this.apiSrvc.SetHeaderData({
            obj: headerdata,
          });
          this.header = [
            {
              type: 'Bar',
              storeIds: this.StoreValues,
              fromdate: this.fromdate,
              todate: this.todate,
              groups: this.groups
            },
          ];
          // if (this.Filter == 'StoreSummary') {
          // this.GetData();
          // } else {
          this.GetDataByMonths(this.StartMonth, this.EndMonth);
          // }
        }
      }
    });
    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Income Statement Trend') return;
      if (obj.state) {
        this.exportAsXLSX();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Income Statement Trend') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Income Statement Trend') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Income Statement Trend') return;
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
  getStores() {
    // this.stores = environment.stores;
    this.selectedstorevalues = [];
  }

  openDetails(Object: any, monthname: any, ref: any, item: any) {
    console.log(Object, monthname, ref, item, this.selectedstorename);
    const date = '01-' + monthname;
    const myDate = this.datepipe.transform(new Date(date), 'dd-MMM-yyyy');
    const DetailsSF = this.ngbmodal.open(IncomestatementtrendDetails, {
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
      STORES: this.storeIds.toString(),
      LatestDate: myDate,
      STORENAMES: this.getSelectedStoreNames(),
      Group: this.groups,
    };
  }

  openGraph(Object: any, monthname: any, ref: any, item: any) {
    const DetailsSF = this.ngbmodal.open(IncomestatementtrendGraph, {
      size: 'xl',
      backdrop: 'static',
      injector: Injector.create({
        providers: [
          { provide: CurrencyPipe, useClass: CurrencyPipe }
        ],
        parent: this.injector
      })
    });
    DetailsSF.componentInstance.INSTGRAPHdetails = {
      ITEM: item,
      TYPE: item.LABLEVAL,
      NAME: item.LABLE,
      STORES: this.storeIds.toString(),
      STORENAMES: this.getSelectedStoreNames(),
    };
  }

  Favreports: any = [];
  getFavReports() {
    const obj = {
      Id: localStorage.getItem('Fav_id'),
      expression: '',
    };
    this.apiSrvc.postmethod('favouritereports/get', obj).subscribe((res) => {
      if (res.status == 200) {
        if (res.response.length > 0) {
          this.Favreports = res.response;
          let dates = this.Favreports[0].Fav_Report_Value.split(',');
          let startdate = new Date(dates[0]);
          let enddate = new Date(dates[1]);
          // //console.log(startdate, enddate);
          (this.fromdate = this.datepipe.transform(startdate, 'dd-MMM-yyyy')),
            (this.todate = this.datepipe.transform(enddate, 'dd-MMM-yyyy'));
          this.StoreValues = res.response[1].Fav_Report_Value;
          // this.DataSelection(this.Filter);
          this.GetDataByMonths(this.StartMonth, this.EndMonth);
          localStorage.setItem('Fav', 'N');
          const data = {
            title: 'Income Statement Trend',
            path1: '',
            path2: '',
            path3: '',
            Month: this.date,
            filter: this.Filter,
            stores: this.StoreValues,
            fromdate: this.datepipe.transform(startdate, 'MM-dd-yyyy'),
            todate:
              ('0' + (enddate.getMonth() + 1)).slice(-2) +
              '-' +
              '01' +
              '-' +
              enddate.getFullYear(),
          };
          // // //console.log(data)
          this.apiSrvc.SetHeaderData({ obj: data });
        }
      }
    });
  }
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


  generatePDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save('Income Statement Trend.pdf'); // ✅ only here
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

    /* ================= TITLE ================= */
    doc.setFontSize(14);
    doc.text('Income Statement Trend', 14, 12);

    /* ❌ REMOVE FILTERS COMPLETELY */
    let startY: number = 20; // ⬅️ directly start table below title

    /* ================= SAFETY CHECK ================= */
    if (!this.MonthsHeadings || this.MonthsHeadings.length === 0) {
      console.error('MonthsHeadings is empty');
      return doc;
    }

    if (!this.ExpenseTrendByStoreMonth || this.ExpenseTrendByStoreMonth.length === 0) {
      console.error('No data for PDF');
      return doc;
    }
    const StartDate: string = this.datepipe.transform(this.StartDate, 'MMM-yy') || '';
    const EndDate: string = this.datepipe.transform(this.EndDate, 'MMM-yy') || '';

    const firstHeader: string = `${EndDate} to ${StartDate}`;
    /* ================= HEADERS ================= */
    const head: any[][] = [
      [
        firstHeader,
        ...this.MonthsHeadings.map((m: string) => m)
      ]
    ];

    /* ================= BODY ================= */
    const body: any[][] = [];

    this.ExpenseTrendByStoreMonth.forEach((d: any) => {
      const row: any[] = [d.LABLE];

      this.MonthsHeadings.forEach((m: string) => {
        const val = d[m];
        const num = Number(val);

        if (val === '' || val === null || num === 0) {
          row.push('-');
        } else {
          row.push(num);
        }
      });

      body.push(row);
    });

    /* ================= TABLE ================= */
    autoTable(doc, {
      startY: startY,
      head: head,
      body: body,
      theme: 'grid',

      styles: {
        fontSize: 8,
        cellPadding: 2,
        halign: 'right',
        valign: 'middle',
        lineWidth: 0.1,
        lineColor: [200, 200, 200],
        textColor: [20, 20, 20], // 👈 slightly richer black
      },

      didParseCell: (data: any) => {

        /* ================= HEADER STYLE ================= */
        if (data.section === 'head') {
          data.cell.styles.fillColor = [69, 132, 255];
          data.cell.styles.textColor = [255, 255, 255];
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.halign = 'center';
          return;
        }
        // if (data.column.index == 1 && data.column.index == 2){
        //   data.cell.styles.cellWidth = 30;
        // }

        /* ================= BODY ================= */
        const rowIndex: number = data.row.index;
        const rowData: any = this.ExpenseTrendByStoreMonth[rowIndex];

        if (!rowData) return;

        const label: string = rowData.LABLE;

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

          const val: number = Number(data.cell.raw);

          if (!isNaN(val)) {

            const isPercent = label.includes('%');
            const isNumber = label.includes('Units');
            const isSpecial = this.isSpecialRow(label);

            if (isPercent) {
              const rounded = Math.round(val); // ✅ FIX HERE
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
      title: 'Income Statement Trend',
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
          value:
            this.datepipe.transform(this.StartMonth, 'MMM yyyy') +
            ' - ' +
            this.datepipe.transform(this.EndMonth, 'MMM yyyy')
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

  ExcelStoreNames: any = [];
  exportAsXLSX(): void {

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Income Statement Trend');

    /* ================= FILTER SECTION ================= */
    const filterRowCount = this.addExcelFiltersSection(worksheet);

    /* ================= HEADER ================= */
    const StartDate = this.datepipe.transform(this.StartDate, 'MMM-yy');
    const EndDate = this.datepipe.transform(this.EndDate, 'MMM-yy');

    const Header = [`${EndDate} to ${StartDate}`, ...this.MonthsHeadings];

    const headerRow = worksheet.addRow(Header);

    headerRow.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '4584FF' }
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
    const formatRow = (row: any, type: string, label: string) => {
      const isSpecial = this.isSpecialRow(label);

      row.eachCell((cell: any, colNumber: number) => {
        // ✅ ADD BORDER TO ALL CELLS
        cell.border = {
          top: { style: 'thin', color: { argb: 'D3D3D3' } },
          bottom: { style: 'thin', color: { argb: 'D3D3D3' } },
          left: { style: 'thin', color: { argb: 'D3D3D3' } },
          right: { style: 'thin', color: { argb: 'D3D3D3' } }
        };
        if (colNumber === 1) {
          cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 2 };
          cell.font = { name: 'Calibri' };
          return;
        }

        if (!cell.value || cell.value === '-') {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          return;
        }

        let num = Number(cell.value);

        if (!isNaN(num)) {
          if (type === 'currency') {
            cell.numFmt = '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)';
          }
          if (type === 'number') cell.numFmt = '#,##0';
          if (type === 'percent') {
            cell.value = num / 100;
            cell.numFmt = '0.0%';
          }
          // ✅ APPLY RED COLOR ONLY FOR SPECIAL ROWS & NEGATIVE VALUES
          if (isSpecial && num < 0) {
            cell.font = {
              ...cell.font,
              color: { argb: 'FFFF0000' }
            };
          }
        }

        cell.alignment = type === 'percent'
          ? { horizontal: 'center', vertical: 'middle' }
          : { horizontal: 'right', vertical: 'middle' };
      });
    };

    /* ================= DATA ================= */
    const data = this.ExpenseTrendByStoreMonth.map((x: any) => ({ ...x }));

    data.forEach((d: any) => {

      let rowData = [d.LABLE];

      this.MonthsHeadings.forEach((m: any) => {
        let val = d[m];
        let num = Number(val);

        if (val === '' || val == null || num === 0) {
          rowData.push('-');
        } else {
          rowData.push(num);
        }
      });
      const row = worksheet.addRow(rowData);

      /* ===== TYPE DETECTION ===== */
      let type = 'currency';

      if (
        d.LABLE.includes('Units')
      ) type = 'number';

      if (
        d.LABLE.includes('%')
      ) type = 'percent';

      /* ===== APPLY FORMAT ===== */
      formatRow(row, type, d.LABLE);
      /* ===== ALT ROW COLOR ===== */
      if (row.number % 2 === 0) {
        row.eachCell(cell => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'F5F5F5' }
          };
        });
      }

      /* ===== BOLD ROWS ===== */
      if (this.isBoldRow(d.LABLE)) {
        row.eachCell((cell: any) => {
          cell.font = {
            ...cell.font,
            bold: true
          };
        });
      }


    });

    /* ================= FREEZE ================= */
    worksheet.views = [
      { state: 'frozen', xSplit: 1, ySplit: filterRowCount }
    ];

    /* ================= COLUMN WIDTH ================= */
    worksheet.columns = [
      { width: 35 },
      ...this.MonthsHeadings.map(() => ({ width: 15 }))
    ];

    /* ================= EXPORT ================= */
    const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');

    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });

      FileSaver.saveAs(
        blob,
        `Income Statement Trend_${DATE_EXTENSION}.xlsx`
      );
    });
  }
  //Special rows
  isSpecialRow(name: string): boolean {
    return [
      'Operating Profit',
      'Net Profit'
    ].includes(name?.trim());
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
      const pdfFile = this.blobToFile(pdfBlob, 'Income Statement Trend.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Income Statement Trend');
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
}

