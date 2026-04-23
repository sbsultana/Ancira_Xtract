import { Component, ElementRef, Injector, ViewChild, HostListener } from '@angular/core';
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
import { Workbook, Row } from 'exceljs';
import FileSaver from 'file-saver';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
import numeral from 'numeral';
import pdfMake from 'pdfmake/build/pdfmake';
import pdfFonts from 'pdfmake/build/vfs_fonts';
(pdfMake as any)['vfs'] = (pdfFonts as any)['vfs'];

import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { FinancialsummaryDetails } from '../financialsummary-details/financialsummary-details';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { PdfExportService } from '../../../../Core/Providers/Shared/export.service';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  Report: any = '';
  FromDate: any;
  ToDate: any;

  reportOpenSub!: Subscription;
  reportGetting!: Subscription;

  Current_Date: any;
  LMY_Date: any;
  LM_Date: any;

  FSData: any = [];
  Month: any = '';
  showAdditionalRows: boolean = false;
  hiddenRows: boolean[] = [];

  EBITDAdata: any = [];
  ETBudgetData: any = [];
  ETDealerData: any = [];

  SelectedTab: any = ['FinancialSummary'];
  Filter: any = 'FinancialSummary';
  SubFilter: any = '';
  StoreName: any = 'All Stores';
  StoreValues: any = [];
  groups: any = 1;

  NoData: boolean = false;
  date: any = new Date();
  roleId: any;
  PresentDayDate: string;
  header: any = [
    {
      type: 'Bar',
      StoreValues: this.StoreValues,
      Month: this.Month,
      groups: this.groups,
    },
  ];
  activePopover: number = -1;
  storeIds: any = '0';
  popup: any = [{ type: 'Popup' }];
  pdfStyleService: any;
  selectedDate: Date = new Date();
  currentMonth!: Date;
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

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  constructor(
    public apiSrvc: Api,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private spinner: NgxSpinnerService,
    private datepipe: DatePipe,
    private title: Title,
    private comm: common,
    private toast: ToastService,
    private injector: Injector,
    public shared: Sharedservice,
    private Export: PdfExportService
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
    this.title.setTitle(this.comm.titleName + '-Financial Summary');
    const lastMonth = new Date();
    let today = new Date();
    if (today.getDate() < 5) {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth() - 1));
    } else {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth()));
    }

    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Financial Summary',
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
      this.header = [
        {
          type: 'Bar',
          StoreValues: this.StoreValues,
          Month: this.date,
          groups: this.groups,
        },
      ];
      this.Month =
        this.date.toString().substr(8, 2) +
        '-' +
        this.date.toString().substr(4, 3) +
        '-' +
        this.date.toString().substr(11, 4);
      this.selectedDate = this.date;
      this.currentMonth = this.selectedDate;
      this.GetData(this.currentMonth);
    }

    const format = 'ddMMyyyy';
    const locale = 'en-US';
    const myDate = new Date();
    const formattedDate = formatDate(myDate, format, locale);
    this.PresentDayDate = formattedDate;
  }

  ngOnInit(): void {
    this.roleId = localStorage.getItem('roleId');
   
  }



  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month',
    maxDate: new Date()
  };

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
      this.GetData(this.currentMonth);
    }
  }

  FsData: any = [];

  Scrollpercent: any = 0;
  updateVerticalScroll(event: any): void {
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );
  }


  GetData(date: Date) {
    this.Month =
      date.toString().substr(8, 2) +
      '-' +
      date.toString().substr(4, 3) +
      '-' +
      date.toString().substr(11, 4);

    this.Current_Date =
      date.toString().substr(8, 2) +
      '-' +
      date.toString().substr(4, 3) +
      '-' +
      date.toString().substr(11, 4);
    this.LMY_Date =
      date.toString().substr(8, 2) +
      '-' +
      date.toString().substr(4, 3) +
      '-' +
      (date.getFullYear() - 1);

    // console.log(this.LMY_Date);

    let LM_StartDate = new Date();
    let LMDate =
      date.toString().substr(8, 2) +
      '-' +
      date.toString().substr(4, 3) +
      '-' +
      date.toString().substr(11, 4);

    let sample = new Date(LMDate.replace(/-/g, '/'));
    LM_StartDate = new Date(sample.setMonth(sample.getMonth() - 1));

    this.LM_Date =
      date.toString().substr(8, 2) +
      '-' +
      LM_StartDate.toString().substr(4, 3) +
      '-' +
      LM_StartDate.toString().substr(11, 4);

    this.FSData = [];
    this.spinner.show();
    const DateToday = this.datepipe.transform(
      new Date(date),
      'yyyy-MM-dd'
    );
    let Obj = {
      as_Id: this.storeIds.toString(),
      SalesDate: this.shared.datePipe.transform(new Date(date),'yyyy-MM')+'-10',
      // UserID: 0,
    };
    console.log(Obj);
    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetFinancialSummaryReport', Obj)
      .subscribe(
        (res) => {
          if (res.status == 200) {
            this.FSData = res.response;
            const newId = Math.floor(Math.random() * 1000) + 1;
            const newObj = {
              id: newId,
              LABLE1: 'Variable Operations',
              MainTitle: 'MainHeader',
            };
            this.FSData.unshift(newObj);
            console.log('Array Values', this.FSData);
            const Index = this.FSData.findIndex(
              (obj: any) => obj.LABLE1 === 'Variable Gross'
            );
            if (Index !== -1 && Index < this.FSData.length - 1) {
              const newId = Math.floor(Math.random() * 1000) + 1;
              const newObj = {
                id: newId,
                LABLE1: 'Fixed Operations',
                MainTitle: 'MainHeader',
              };
              this.FSData.splice(Index + 1, 0, newObj);
            }
            const IndexOne = this.FSData.findIndex(
              (obj: any) => obj.LABLE1 === 'Total Store Gross'
            );
            if (IndexOne !== -1 && IndexOne < this.FSData.length - 1) {
              const newId = Math.floor(Math.random() * 1000) + 1;
              const newObj = {
                id: newId,
                LABLE1: 'Expenses',
                MainTitle: 'MainHeader',
              };
              this.FSData.splice(IndexOne + 1, 0, newObj);
            }
            const IndexTwo = this.FSData.findIndex(
              (obj: any) => obj.LABLE1 === 'Operating Profit'
            );
            if (IndexTwo !== -1 && IndexTwo < this.FSData.length - 1) {
              const newId = Math.floor(Math.random() * 1000) + 1;
              const newObj = {
                id: newId,
                LABLE1: 'Net Adjustments/Other Income',
                MainTitle: 'MainHeader',
              };
              this.FSData.splice(IndexTwo + 1, 0, newObj);
            }
            const IndexThree = this.FSData.findIndex(
              (obj: any) => obj.LABLE1 === 'Net to Gross'
            );
            if (IndexThree !== -1 && IndexThree < this.FSData.length - 1) {
              const newId = Math.floor(Math.random() * 1000) + 1;
              const newObj = {
                id: newId,
                LABLE1: 'Memo: Super Gross',
                MainTitle: 'MainHeader',
              };
              this.FSData.splice(IndexThree + 1, 0, newObj);
            }
            this.FSData.forEach((val: any) => {
              if (
                val.LABLE1 == 'Unit Retail Sales' ||
                val.LABLE1 == 'Pure Gross' ||
                val.LABLE1 == 'Variable Gross' ||
                val.LABLE1 == 'Total Store Gross' ||
                val.LABLE1 == 'Selling Gross' ||
                val.LABLE1 == 'Selling Gross%' ||
                val.LABLE1 == 'Total Expenses' ||
                val.LABLE1 == 'Operating Profit' ||
                val.LABLE1 == 'Net Adds/Deducts' ||
                val.LABLE1 == 'Net Profit' ||
                val.LABLE1 == 'Total Store Super Gross'
              ) {
                val.Fs_Titles = 'FontBold';
              } else if (
                val.LABLE1 == 'New' ||
                val.LABLE1 == 'Used' ||
                val.LABLE1 == 'Service' ||
                val.LABLE1 == 'Parts'
              ) {
                val.Fs_Titles = 'PaddingLeft';
              }
            });
            this.FSData.forEach((val: any, index: any) => {
              if (val.LABLE1 == 'Selling Gross%') {
                this.showAdditionalRows = !this.showAdditionalRows;
                for (
                  let i = index + 1;
                  i < index + 5 && i < this.FSData.length;
                  i++
                ) {
                  this.hiddenRows[i] = true;
                }
              }
            });
            this.spinner.hide();
            if (this.FSData.length > 0) {
              this.NoData = false;
            } else {
              this.NoData = true;
            }
          } else {

            this.toast.show('Invalid Details.', 'danger', 'Error');
          }
        },
        (error) => {
          console.log(error);
        }
      );
  }

  toggleRows(index: number): void {
    console.log(index);
    if (this.FSData[index].LABLE1 === 'Selling Gross%') {
      this.showAdditionalRows = !this.showAdditionalRows;
      for (let i = index + 1; i < index + 5 && i < this.FSData.length; i++) {
        this.hiddenRows[i] = !this.hiddenRows[i];
        console.log(this.hiddenRows);
      }
    }
  }
  getStyle(value: any, valueType: string) {
    if (value === null || value === '' || value === 0 || value === '0') {
      return { 'justify-content': 'center' };
    }
    switch (valueType) {
      case '$':
        return { 'justify-content': 'space-between' };
      case '#':
        return { 'justify-content': 'flex-end' };
      case '%':
        return { 'justify-content': 'center' };
      default:
        return { 'justify-content': 'initial' };
    }
  }
  formatValue(value: any, valueType: string): string {
    if (value === null || value === undefined || value === 0 || value === '0') {
      return '-';
    } else if (valueType === '#') {
      return typeof value === 'number' ? numeral(value).format('0,0') : value;
    } else if (valueType === '$') {
      return typeof value === 'number' ? numeral(value).format('0,0') : value;
    } else if (valueType === '%') {
      return typeof value === 'number' ? value.toLocaleString() + '%' : value;
    } else {
      return value;
    }
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  // StoreValues: any = 0;
  FsClick: any;
  block: any = '';
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Financial Summary') {
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
      if (this.reportOpenSub != undefined) {
        if (res.obj.Module == 'Financial Summary') {
          document.getElementById('report')?.click();
        }
      }
    });
    this.reportGetting = this.apiSrvc.GetReports().subscribe((data) => {
      console.log(data);
      if (this.reportGetting != undefined) {
        if (data.obj.Reference == 'Financial Summary') {
          if (data.obj.header == undefined) {
            this.date = data.obj.month;
            this.Month = data.obj.month;
            this.StoreValues = data.obj.storeValues;
            this.StoreName = data.obj.Sname;
            this.groups = data.obj.groups;
          } else {
            if (data.obj.header == 'Yes') {
              this.StoreValues = data.obj.storeValues;
            }
          }
          if (this.StoreValues != '') {
            this.GetData(this.currentMonth);
          } else {
            this.NoData = true;
            this.FSData = [];
          }
          const headerdata = {
            title: 'Financial Summary',
            path1: '',
            path2: '',
            path3: '',
            Month: new Date(this.Month),
            stores: this.StoreValues,
            groups: this.groups,
          };
          this.apiSrvc.SetHeaderData({
            obj: headerdata,
          });
          this.header = [
            {
              type: 'Bar',
              StoreValues: this.StoreValues,
              Month: new Date(this.Month),
              groups: this.groups,
            },
          ];
        }
      }
    });
    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Financial Summary') return;
      if (obj.state) {
        this.exportToExcelFinancialSummary();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Financial Summary') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Financial Summary') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Financial Summary') return;
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
  SubSelectedTab1: any = [];

  openMenu(Object: any, LatestDate: any) {
    const DetailsFs = this.ngbmodal.open(FinancialsummaryDetails, {
      size: 'xl',
      backdrop: 'static',
      injector: Injector.create({
        providers: [
          { provide: CurrencyPipe, useClass: CurrencyPipe }
        ],
        parent: this.injector
      })
    });

    DetailsFs.componentInstance.Fsdetails = {
      TYPE: Object.LABLE1Val,
      NAME: Object.LABLE1,
      STORES: this.storeIds,
      LatestDate: LatestDate,
      Group: this.groups,
    };
  }
  GraphHeadings: any;
  GraphData: any;
  ValueFormat: any;
  openGraph(Object: any, Current_Date: any, LMY_Date: any, LM_Date: any) {
    // console.log(Object, Current_Date, LMY_Date, LM_Date);
    const CurrentMonth = this.datepipe.transform(Current_Date, 'MMMM yyyy');
    const LastYearMonth = this.datepipe.transform(LMY_Date, 'MMMM yyyy');
    const LastMonth = this.datepipe.transform(LM_Date, 'MMMM yyyy');
    const CurrentYear = this.datepipe.transform(Current_Date, 'yyyy');
    const LastYear = this.datepipe.transform(LMY_Date, 'yyyy');
    this.GraphHeadings = [
      CurrentMonth,
      LastYearMonth,
      LastMonth,
      CurrentYear + ' Average',
      LastYear + ' Average',
    ];
    this.GraphData = [
      Object.MTD,
      Object.LMY_MTD,
      Object.LM_MTD,
      Object.YTD,
      Object.LY_YTD,
    ];
    if (
      Object.LABLE1 == 'New Units' ||
      Object.LABLE1 == 'Pre-Owned Units' ||
      Object.LABLE1 == 'Wholesale Units' ||
      Object.LABLE1 == 'Unit Retail Sales'
    ) {
      this.ValueFormat = 'Number';
    } else if (
      Object.LABLE1 == 'Selling Gross%' ||
      Object.LABLE1 == 'New' ||
      Object.LABLE1 == 'Used' ||
      Object.LABLE1 == 'Service' ||
      Object.LABLE1 == 'Parts' ||
      Object.LABLE1 == 'Net to Sales' ||
      Object.LABLE1 == 'Net to Gross'
    ) {
      this.ValueFormat = 'Percentage';
    } else {
      this.ValueFormat = 'Currancy';
    }
    const DetailsSF = this.ngbmodal.open({
      size: 'xl',
      backdrop: 'static',
    });
    DetailsSF.componentInstance.FSgraphdetails = {
      Header: this.GraphHeadings,
      Data: this.GraphData,
      NAME: Object.LABLE1,
      ValueFormat: this.ValueFormat,
      STORES: this.StoreValues,
      LatestDate: this.Month,
      Group: this.groups,
    };
  }
  openfs() {
    const DetailsFs = this.ngbmodal.open({
      size: 'xl',
      backdrop: 'static',
    });
  }

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }



  private createPDF(): jsPDF {
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });

    doc.setFontSize(14);
    doc.text('Financial Summary', 14, 12);

    // const filtersObj = this.getReportFilters();
    doc.setFontSize(10);

    // let startY = 18;
    // filtersObj.filters.forEach((f: any) => {
    //   doc.text(`${f.label}: ${f.value}`, 14, startY);
    //   startY += 5;
    // });
    // startY += 5;

    const currentMonthName = this.datepipe.transform(this.currentMonth, 'MMMM');
    const currentYear = this.datepipe.transform(this.currentMonth, 'yyyy');
    const lmyMonth = this.datepipe.transform(this.LMY_Date, 'MMMM');
    const lmyYear = this.datepipe.transform(this.LMY_Date, 'yyyy');
    const lmMonth = this.datepipe.transform(this.LM_Date, 'MMMM');
    const fullMonth = this.datepipe.transform(this.currentMonth, 'MMMM yyyy');

    const headers: any = [
      [
        { content: 'For the period ending', colSpan: 1 },
        { content: 'MTD', colSpan: 4 },
        { content: 'LY MTD', colSpan: 2 },
        { content: 'LM', colSpan: 2 },
        { content: 'YTD', colSpan: 3 },
        { content: 'LY YTD', colSpan: 2 }
      ],
      [
        { content: fullMonth },
        { content: currentMonthName },
        { content: 'Pace' },
        { content: 'Budget' },
        { content: 'Variance' },
        { content: lmyMonth },
        { content: 'Variance' },
        { content: lmMonth },
        { content: 'Variance' },
        { content: currentYear },
        { content: 'Budget' },
        { content: 'Variance' },
        { content: lmyYear },
        { content: 'Variance' }
      ]
    ];

    const body: any[] = [];

    this.FSData.forEach((data: any) => {
      body.push({
        rowData: data,
        values: [
          data.LABLE1,
          data.MTD,
          data.PACE,
          data.BUDGET,
          data.VARIANCE,
          data.LMY_MTD,
          data.LMY_VARIANCE,
          data.LM_MTD,
          data.LM_VARIANCE,
          data.YTD,
          data.CY_BUDGET,
          data.CY_VARIANCE,
          data.LY_YTD,
          data.LY_VARIANCE
        ]
      });
    });

    autoTable(doc, {
      startY: 18,
      head: headers,
      body: body.map(r => r.values),
      theme: 'grid',

      styles: {
        fontSize: 8,
        cellPadding: 2,
        halign: 'right',
        lineWidth: 0.1,
        lineColor: [200, 200, 200],
        textColor: [20, 20, 20], // 👈 slightly richer black
        valign: 'middle'
      },

      headStyles: {
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        lineWidth: 0.1,
        lineColor: [200, 200, 200]
      },

      didParseCell: (data: any) => {
        const rowObj = body[data.row.index];
        const rowData = rowObj?.rowData;
        const isHeaderRow = rowData?.MainTitle === 'MainHeader';
        const isBold = rowData?.Fs_Titles === 'FontBold';

        if (data.section === 'head') {
          if (data.row.index === 0) {
            data.cell.styles.fillColor = [5, 84, 239];
            data.cell.styles.textColor = [255, 255, 255];
            data.cell.styles.fontStyle = 'bold';
            data.cell.styles.halign = 'center';
          }
          if (data.row.index === 1) {
            data.cell.styles.fillColor = [69, 132, 255];
            data.cell.styles.textColor = [255, 255, 255];
            data.cell.styles.fontStyle = 'bold';
            data.cell.styles.halign = data.column.index === 0 ? 'left' : 'center';
          }
          return;
        }

        if (!rowData) return;
        data.cell.styles.halign =
          data.column.index === 0 ? 'left' : 'right';

        const valEmpty =
          data.cell.raw === null ||
          data.cell.raw === undefined ||
          data.cell.raw === 0 ||
          data.cell.raw === '';

        if (isHeaderRow) {
          if (valEmpty) {
            data.cell.text = [''];
            return;
          }
        } else {
          if (valEmpty) {
            data.cell.text = ['-'];
            data.cell.styles.halign = 'center';
            return;
          }
        }


        if (data.column.index > 0) {
          const valueType = rowData.ValueType;
          let val = parseFloat(data.cell.raw);

          if (!isNaN(val)) {
            const rounded = Math.round(val);

            if (valueType === '$') {
              data.cell.text = [`$ ${rounded.toLocaleString()}`];
            } else if (valueType === '#') {
              data.cell.text = [rounded.toLocaleString()];
            } else if (valueType === '%') {
              data.cell.text = [`${rounded}%`];
            }

            if (valueType === '$' && rounded < 0) {
              data.cell.styles.textColor = [220, 53, 69];
            }

            if (this.isSpecialRow?.(rowData.LABLE1) && rounded < 0 && valueType === '$') {
              data.cell.styles.textColor = [255, 0, 0];
            }
          }
        }

        /* ================= SECTION HEADER ROW ================= */
        if (isHeaderRow) {
          Object.keys(data.row.cells).forEach((key: any) => {
            const cell = data.row.cells[key];
            cell.styles.fillColor = [217, 231, 255];
            cell.styles.fontStyle = 'bold';
          });
        }

        /* ================= BOLD ROWS ================= */
        if (isBold) {
          Object.keys(data.row.cells).forEach((key: any) => {
            data.row.cells[key].styles.fontStyle = 'bold';
          });
        }

        /* ================= ALTERNATE ROW COLOR ================= */
        if (!isHeaderRow && data.row.index % 2 === 0) {
          Object.keys(data.row.cells).forEach((key: any) => {
            data.row.cells[key].styles.fillColor = [245, 247, 250];
          });
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
      title: 'Financial Summary',
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
  createExcelWorkbook(): Workbook {

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Financial Summary');

    /* ================= 1. FILTERS AT TOP ================= */
    const filterRowCount = this.addExcelFiltersSection(worksheet);

    const startRow = filterRowCount + 1;

    /* ================= 2. HEADER ROW 1 ================= */

    worksheet.getCell(`A${startRow}`).value = 'For the period ending';

    worksheet.getCell(`B${startRow}`).value = 'MTD';
    worksheet.mergeCells(`B${startRow}:E${startRow}`);

    worksheet.getCell(`F${startRow}`).value = 'LY MTD';
    worksheet.mergeCells(`F${startRow}:G${startRow}`);

    worksheet.getCell(`H${startRow}`).value = 'LM';
    worksheet.mergeCells(`H${startRow}:I${startRow}`);

    worksheet.getCell(`J${startRow}`).value = 'YTD';
    worksheet.mergeCells(`J${startRow}:L${startRow}`);

    worksheet.getCell(`M${startRow}`).value = 'LY YTD';
    worksheet.mergeCells(`M${startRow}:N${startRow}`);

    worksheet.getRow(startRow).eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0554EF' } };
      cell.font = { bold: true, color: { argb: 'FFFFFF' }, name: 'Calibri' };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    /* ================= 3. HEADER ROW 2 ================= */

    const header2 = [
      this.datepipe.transform(this.currentMonth, 'MMMM yyyy'),

      this.datepipe.transform(this.currentMonth, 'MMMM'),
      'Pace',
      'Budget',
      'Variance',

      this.datepipe.transform(this.LMY_Date, 'MMMM'),
      'Variance',

      this.datepipe.transform(this.LM_Date, 'MMMM'),
      'Variance',

      this.datepipe.transform(this.currentMonth, 'yyyy'),
      'Budget',
      'Variance',

      this.datepipe.transform(this.LMY_Date, 'yyyy'),
      'Variance'
    ];

    const headerRow2 = worksheet.addRow(header2);

    headerRow2.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '4584FF' } };
      cell.font = { bold: true, color: { argb: 'FFFFFF' }, name: 'Calibri' };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    /* ================= VARIANCE COLUMN INDEX ================= */

    const varianceColumns: number[] = [];
    header2.forEach((h: any, i) => {
      if (h && h.toString().toLowerCase().includes('variance')) {
        varianceColumns.push(i + 1);
      }
    });

    /* ================= FORMAT FUNCTION ================= */

    const formatRow = (row: Row, data: any, index: number) => {

      const isHeaderRow = data.MainTitle === 'MainHeader';
      const isBold = data.Fs_Titles === 'FontBold';
      const isPadding = data.Fs_Titles === 'PaddingLeft';
      const isSpecial = this.isSpecialRow(data.LABLE1);
      const valueType = data.ValueType;

      const totalCols = 14;

      let bgColor = '';
      if (isHeaderRow) bgColor = 'D9E7FF';
      else bgColor = index % 2 === 0 ? 'F9FBFF' : 'FFFFFF';

      for (let i = 1; i <= totalCols; i++) {
        row.getCell(i);
      }

      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {

        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: bgColor }
        };

        /* ===== FIRST COLUMN ===== */
        if (colNumber === 1) {
          cell.font = {
            name: 'Calibri',
            bold: isBold || isSpecial || isHeaderRow
          };

          cell.alignment = {
            horizontal: 'left',
            vertical: 'middle',
            indent: isHeaderRow ? 1 : (isPadding ? 3 : 2)
          };
          return;
        }

        if (isHeaderRow) {
          cell.value = '';
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.font = { name: 'Calibri', bold: true };
          return;
        }

        if (cell.value === null || cell.value === undefined || cell.value === '' || cell.value === 0) {
          cell.value = '-';
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          return;
        }

        const num = Number(cell.value);
        let fontStyle: any = { name: 'Calibri', bold: isBold };

        if (!isNaN(num)) {

          if (valueType === '$') {
            cell.numFmt = '"$" * #,##0; "$" * -#,##0';
          } else if (valueType === '#') {
            cell.numFmt = '#,##0';
          } else if (valueType === '%') {
            cell.numFmt = '0%';
            cell.value = num / 100;
          }

          if (isSpecial && varianceColumns.includes(colNumber) && num < 0) {
            fontStyle.color = { argb: 'FF0000' };
          }
        }

        cell.font = fontStyle;

        cell.alignment = valueType === '%'
          ? { horizontal: 'center', vertical: 'middle' }
          : { horizontal: 'right', vertical: 'middle' };
      });
    };

    /* ================= DATA ================= */

    this.FSData.forEach((data: any, i: number) => {

      const row = worksheet.addRow([
        data.LABLE1,
        data.MTD,
        data.PACE,
        data.BUDGET,
        data.VARIANCE,
        data.LMY_MTD,
        data.LMY_VARIANCE,
        data.LM_MTD,
        data.LM_VARIANCE,
        data.YTD,
        data.CY_BUDGET,
        data.CY_VARIANCE,
        data.LY_YTD,
        data.LY_VARIANCE
      ]);

      formatRow(row, data, i);
    });

    /* ================= BORDERS ================= */

    worksheet.eachRow(row => {
      row.eachCell(cell => {
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    /* ================= WIDTH ================= */

    worksheet.columns = [
      { width: 35 },
      ...Array(13).fill({ width: 18 })
    ];

    /* ================= FREEZE ================= */

    worksheet.views = [{
      state: 'frozen',
      xSplit: 1,
      ySplit: startRow + 1,
      showGridLines: false
    }];

    /* ================= DOWNLOAD ================= */

    return workbook;
  }
  exportToExcelFinancialSummary() {
    const workbook = this.createExcelWorkbook();
    workbook.xlsx.writeBuffer().then(data => {
      FileSaver.saveAs(
        new Blob([data]),
        'FinancialSummary.xlsx'
      );
    });
  }
  generatePDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save('Financial Summary.pdf');
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

  //Special rows
  isSpecialRow(name: string): boolean {
    return [
      'Total Selling Gross (New)',
      'Total Selling Gross (Used)',
      'Net Income Before Taxes',
      'Total Operating Department Profit',
      'Operating Profit',
      'Net Profit'
    ].includes(name);
  }


  GetPrintData() {
    window.print();
  }


  @ViewChild('printSection') printSection!: ElementRef;


}


