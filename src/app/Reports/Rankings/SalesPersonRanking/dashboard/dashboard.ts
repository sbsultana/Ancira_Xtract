import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { NgbModal } from '@ng-bootstrap/ng-bootstrap';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { SalesgrossDetails } from '../../../Sales/SalesGross/salesgross-details/salesgross-details';
import { Subscription } from 'rxjs';


@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  IndividualSalesPersonsData: any = [];

  FromDate: any = '';
  ToDate: any = '';
  DupFromDate: any = '';
  DupToDate: any = ''
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  DupDateType: any = 'MTD';
  displaytime: any = '';

  NoData: boolean = false;

  storeIds!: any;
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;


  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'Y',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname,
  
  };

  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    Types: [
      { 'code': 'MTD', 'name': 'MTD' },
      { 'code': 'QTD', 'name': 'QTD' },
      { 'code': 'YTD', 'name': 'YTD' },
      { 'code': 'PYTD', 'name': 'PYTD' },
      { 'code': 'LY', 'name': 'Last Year' },
      { 'code': 'LM', 'name': 'Last Month' },
      { 'code': 'PM', 'name': 'Same Month PY' },
    ]
  }



  storeorgrp: any = 'G';
  saleType: any = 'Retail,Lease,Misc,Special Order';
  retailorlease: any = this.saleType.split(',');
  columnName: any = 'Rank';
  columnState: any = 'asc';
  dealStatus: any = ['Booked', 'Finalized', 'Delivered'];
  GrossType: any = ['Front Gross'];
  DupGrossType: any = [];

  constructor(public shared: Sharedservice, public setdates: Setdates, private comm: common, private toast: ToastService, private ngbmodal: NgbModal
  ) {
    if (typeof window !== 'undefined') {
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
      if (this.shared.common.groupsandstores.length > 0) {
        this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
        this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
        this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
        this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
        this.getStoresandGroupsValues()
      }

      if (localStorage.getItem('stime') != null) {
        let stime = localStorage.getItem('stime');
        if (stime != null && stime != '') {
          this.initializeDates(stime)
        }
      } else {
        this.initializeDates('MTD')
      }
      this.shared.setTitle(this.shared.common.titleName + '-Salesperson Rankings');
      this.setHeaderData()
      this.GetData('Rank', 'asc');

    }
  }

  ngOnInit(): void { }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.DateType = type;
    this.setDates(this.DateType)


  }
  setHeaderData() {
    const data = {
      title: 'Salesperson Rankings',
      stores: this.storeIds,
      datetype: this.DateType,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groupId,
      storeorgroup: this.storeorgrp,
      saleType: this.saleType.toString(),
      dealStatus: this.dealStatus,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }

  openDetails(data: any) {
    const DetailsSalesPeron = this.shared.ngbmodal.open(SalesgrossDetails, { size: 'xxl', backdrop: 'static', windowClass: 'SalesDetails' });
    DetailsSalesPeron.componentInstance.Salesdetails = [
      {
        StartDate: this.FromDate,
        EndDate: this.ToDate,
        dealtype: "New,Used",
        saletype: this.saleType,
        dealstatus: "Delivered,Booked,Finalized",
        var1: 'store',
        var2: 'salesperson',
        var3: '',
        var1Value: data.StoreName,
        var2Value: data.SalesPerson,
        var3Value: '',
        userName: data.SPTrimName,
        FinanceManager: "0",
        SalesManager: "0",
        SalesPerson: "0"
      },
    ];
    DetailsSalesPeron.result.then((data: any) => { }, (reason: any) => { });
  }


  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;

  }


  tabClick(col_Name: any) {
    if (this.columnName !== col_Name) {
      this.columnName = col_Name;
      this.columnState = 'asc';
    }
    else {
      this.columnState = this.columnState === 'asc' ? 'desc' : 'asc';
    }
    this.GetData(this.columnName, this.columnState);
  }
  GetData(sortdata?: any, sortstate?: any) {
    console.log(sortdata, sortstate, this.storeIds);
    this.DupGrossType = [...this.GrossType]
    this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    this.DupDateType = this.DateType

    this.IndividualSalesPersonsData = [];
    this.shared.spinner.show();
    const obj = {
      UserID: 0,
      StartDate: this.FromDate,
      EndDate: this.ToDate,
      StoreID: [...this.storeIds],
      Exp: sortdata,
      OrderType: sortstate,
      RankBy: this.storeorgrp,
      DealType: this.saleType,
      // DealStatus: this.dealStatus.toString(),
      // IncludeBackGross: this.GrossType.indexOf('Back Gross') >= 0 ? 'Y' : 'N'
    };
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetSalesPersonsRankings', obj)
      .subscribe(
        (res) => {
          if (res.status == 200) {
            if (res.response != undefined) {
              if (res.response.length > 0) {
                this.IndividualSalesPersonsData = res.response;
                this.shared.spinner.hide();
                this.NoData = false;
              } else {
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            this.shared.spinner.hide();
            this.NoData = true;
          }
        },
        (error) => {
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  SPRstate: any;
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Salesperson Rankings') {
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
      if (!obj || obj.title !== 'Salesperson Rankings') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Salesperson Rankings') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Salesperson Rankings') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Salesperson Rankings') return;
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
      'type': 'M', 'others': 'Y', 
    };

  }

  updatedDates(data: any) {
    // console.log(data);
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
  }
  setDates(type: any) {


    this.displaytime = '(' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ')';
    this.maxDate = new Date();
    this.minDate = new Date();
    this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    this.maxDate.setDate(this.maxDate.getDate());
    this.Dates.FromDate = this.FromDate;
    this.Dates.ToDate = this.ToDate;
    this.Dates.MinDate = this.minDate;
    this.Dates.MaxDate = this.maxDate;
    this.Dates.DateType = this.DateType;
    this.Dates.DisplayTime = this.displaytime;
  }


  activePopover: number = -1;
  storeorgroup: any = ['G'];
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

  multipleorsingle(block: string, val: string) {
    if (block === 'RL') {
      this.toggleSelection(this.retailorlease, val);
    }
    if (block === 'DS') {
      this.toggleSelection(this.dealStatus, val);
    }
    if (block === 'GT') {
      this.toggleSelection(this.GrossType, val);
    }
  }

  private toggleSelection(arr: any[], val: string) {
    const idx = arr.indexOf(val);
    if (idx >= 0) {
      arr.splice(idx, 1);
    } else {
      arr.push(val);
    }
  }

  storeorgroups(_block: any, val: string) {
    this.storeorgroup = [val];
  }




  close() {
    this.shared.ngbmodal.dismissAll();
  }


  viewreport() {
    this.activePopover = -1;

    if ((!this.storeIds || this.storeIds.length === 0)) {

      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      return;
    }
    else if (this.retailorlease.length == 0) {

      this.toast.show('Please select any one Deal Type', 'warning', 'Warning');
    }
    else if (this.dealStatus.length == 0) {

      this.toast.show('Please Select Atleast One Deal Status', 'warning', 'Warning');
    }
    else if (this.GrossType && this.GrossType.length == 0) {
      this.toast.show('Please Select Atleast One Gross Type', 'warning', 'Warning');
    }
    else {
      this.setHeaderData()
      this.GetData(this.columnName, this.columnState);


    }
  }

  datetype() {
    if (this.DateType == 'PM') {
      return 'SP';
    }
    else if (this.DateType == 'C') {
      return 'C'
    }
    return this.DateType;
  }
  ExcelStoreNames: any = []
  exportToExcel(): void {

    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Salesperson Rankings');
  
    /* ================= TITLE ================= */
    const title = worksheet.addRow(['Salesperson Rankings']);
    title.font = { size: 14, bold: true };
    worksheet.mergeCells('A1:R1');
  
    worksheet.addRow([]);
  
    /* ================= FILTERS ================= */
  
    const formattedFromDate = this.shared.datePipe.transform(this.FromDate, 'dd-MMM-yyyy');
    const formattedToDate = this.shared.datePipe.transform(this.ToDate, 'dd-MMM-yyyy');
  
    let storeValue = '';
  
    if (!this.storeIds || this.storeIds.length === 0 || this.storeIds.length === this.stores.length) {
      storeValue = this.stores.map((s: any) => s.storename).join(', ');
    } else {
      storeValue = this.stores
        .filter((s: any) => this.storeIds.includes(s.ID))
        .map((s: any) => s.storename)
        .join(', ');
    }
  
    const filters = [
      { name: 'Store:', values: storeValue },
      { name: 'Time Frame:', values: `${formattedFromDate} to ${formattedToDate}` },
      { name: 'Rank By:', values: this.storeorgroup == 'S' ? 'Store' : 'Group' },
      { name: 'Deal Type:', values: this.retailorlease || 'All' },
      { name: 'Deal Status:', values: this.dealStatus?.toString() || 'All' },
      { name: 'Gross Type:', values: this.GrossType?.toString() || 'All' }
    ];
  
    filters.forEach(f => {
      const row = worksheet.addRow([f.name, f.values]);
      row.getCell(1).font = { bold: true};
      worksheet.mergeCells(`B${row.number}:F${row.number}`);
    });
  
    worksheet.addRow([]);
  
    /* ================= DATE HEADER ================= */
  
    let dateHeader: any = '';
    if (this.datetype() === 'C') {
      const from = new Date(this.FromDate);
      const to = new Date(this.ToDate);
  
      const format = (d: Date) =>
        ('0' + (d.getMonth() + 1)).slice(-2) + '.' +
        ('0' + d.getDate()).slice(-2) + '.' +
        d.getFullYear();
  
      dateHeader = `${format(from)} - ${format(to)}`;
    } else {
      dateHeader = this.datetype();
    }
  
    /* ================= HEADERS ================= */
  
    let firstHeader: any[] = [];
    let secondHeader: any[] = [];
    let bindingHeaders: any[] = [];
  
    const startRow = worksheet.rowCount + 1;
  
    if (this.DupGrossType.length > 1) {
  
      firstHeader = [
        dateHeader, '', '',
        'Unit Count', '', '', '', '',
        'Front Gross', '', '',
        'Back Gross', '', '',
        'Total Gross', '', ''
      ];
  
      secondHeader = [
        'Rank', 'Salesperson', 'Store Name',
        'New', 'Used', 'Total', 'Pace', '90 Day Avg',
        'Gross', 'Pace', 'PVR',
        'Gross', 'Pace', 'PVR',
        'Gross', 'Pace', 'PVR'
      ];
  
      bindingHeaders = [
        'Rank', 'SalesPerson', 'StoreName',
        'MTD_NEW', 'MTD_USED', 'MTD_Total', 'Pace', 'UnitDayAvg',
        'Gross', 'GrossPace', 'PVR',
        'BackGross', 'BackGross_Pace', 'BackGross_PVR',
        'TotalGross', 'TotalGross_Pace', 'TotalGross_PVR'
      ];
  
      worksheet.mergeCells(`A${startRow}:C${startRow}`);
      worksheet.mergeCells(`D${startRow}:H${startRow}`);
      worksheet.mergeCells(`I${startRow}:K${startRow}`);
      worksheet.mergeCells(`L${startRow}:N${startRow}`);
      worksheet.mergeCells(`O${startRow}:Q${startRow}`);
  
    } else if (this.DupGrossType.includes('Front Gross')) {
  
      firstHeader = [
        dateHeader, '', '',
        'Unit Count', '', '', '', '',
        'Front Gross', '', ''
      ];
  
      secondHeader = [
        'Rank', 'Salesperson', 'Store Name',
        'New', 'Used', 'Total', 'Pace', '90 Day Avg',
        'Total', 'Pace', 'PVR'
      ];
  
      bindingHeaders = [
        'Rank', 'SalesPerson', 'StoreName',
        'MTD_NEW', 'MTD_USED', 'MTD_Total', 'Pace', 'UnitDayAvg',
        'Gross', 'GrossPace', 'PVR'
      ];
  
      worksheet.mergeCells(`A${startRow}:C${startRow}`);
      worksheet.mergeCells(`D${startRow}:H${startRow}`);
      worksheet.mergeCells(`I${startRow}:K${startRow}`);
  
    } else {
  
      firstHeader = [
        dateHeader, '', '',
        'Unit Count', '', '', '', '',
        'Back Gross', '', ''
      ];
  
      secondHeader = [
        'Rank', 'Salesperson', 'Store Name',
        'New', 'Used', 'Total', 'Pace', '90 Day Avg',
        'Gross', 'Pace', 'PVR'
      ];
  
      bindingHeaders = [
        'Rank', 'SalesPerson', 'StoreName',
        'MTD_NEW', 'MTD_USED', 'MTD_Total', 'Pace', 'UnitDayAvg',
        'BackGross', 'BackGross_Pace', 'BackGross_PVR'
      ];
  
      worksheet.mergeCells(`A${startRow}:C${startRow}`);
      worksheet.mergeCells(`D${startRow}:H${startRow}`);
      worksheet.mergeCells(`I${startRow}:K${startRow}`);
    }
  
    const headerRow1 = worksheet.addRow(firstHeader);
    const headerRow2 = worksheet.addRow(secondHeader);
  
    /* ================= HEADER STYLE ================= */
  
    // FIRST HEADER
    headerRow1.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0554EF' }
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
  
    // SECOND HEADER
    headerRow2.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9E7FF' }
      };
      cell.font = { bold: true, color: { argb: 'FF000000' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
  
    /* ================= DATA ================= */
  
    const currencyFields = [
      'Gross', 'GrossPace', 'PVR',
      'BackGross', 'BackGross_Pace', 'BackGross_PVR',
      'TotalGross', 'TotalGross_Pace', 'TotalGross_PVR'
    ];
  
    this.IndividualSalesPersonsData.forEach((item: any, i: number) => {
  
      const row = worksheet.addRow(
        bindingHeaders.map(k => item[k] ?? '-')
      );
  
      row.eachCell((cell, col) => {
  
        // ALIGNMENT
        if (col === 1) {
          cell.alignment = { horizontal: 'center' }; // Rank
        } else if (col === 2 || col === 3) {
          cell.alignment = { horizontal: 'left' };
        } else {
          cell.alignment = { horizontal: 'right' };
        }
  
        // CURRENCY
        if (currencyFields.includes(bindingHeaders[col - 1]) && typeof cell.value === 'number') {
          cell.numFmt = '"$"#,##0';
        }
  
        // BORDERS
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFDADADA' } },
          left: { style: 'thin', color: { argb: 'FFDADADA' } },
          bottom: { style: 'thin', color: { argb: 'FFDADADA' } },
          right: { style: 'thin', color: { argb: 'FFDADADA' } }
        };
  
        // ALTERNATE ROW COLOR
        if (i % 2 === 1) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF5F5F5' }
          };
        }
      });
    });
  
    /* ================= COLUMN WIDTH ================= */
  
    worksheet.columns.forEach((col, i) => {
      if (i === 1 || i === 2) {
        col.width = 30;
      } else {
        col.width = 18;
      }
    });
  
    /* ================= FREEZE HEADER ================= */
    // worksheet.views = [{ state: 'frozen', ySplit: headerRow2.number }];
  
    /* ================= EXPORT ================= */
    workbook.xlsx.writeBuffer().then(() => {
      this.shared.exportToExcel(workbook, 'Salesperson Rankings');
    });
  
  }

  private createPDF(): jsPDF {

    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a3'
    });

    /* ================= TITLE ================= */
    doc.setFontSize(14);
    doc.text('Salesperson Rankings', 14, 12);

    const startY = 18;

    /* ================= DATE HEADER ================= */
    const fromMonth = this.shared.datePipe.transform(this.DupFromDate, 'MMMM');
    const fromDay = this.shared.datePipe.transform(this.DupFromDate, 'dd');
    const toDay = this.shared.datePipe.transform(this.DupToDate, 'dd');
    const year = this.shared.datePipe.transform(this.DupFromDate, 'yyyy');

    const dateText =
      this.DupDateType === ''
        ? `${fromMonth} (${fromDay}-${toDay}, ${year})`
        : (
          this.DupDateType === 'MTD' ? 'MTD' :
            this.DupDateType === 'YTD' ? 'YTD' :
              this.DupDateType === 'QTD' ? 'QTD' :
                this.DupDateType === 'LM' ? 'LAST MONTH' :
                  this.DupDateType === 'LY' ? 'LAST YEAR' :
                    `${this.DupFromDate} - ${this.DupToDate}`
        );

    /* ================= HEAD ================= */
    const head: any[] = [];

    /* ---- Header Row 1 (Group headers) ---- */
    const groupRow: any[] = [
      { content: dateText, colSpan: 3, styles: { halign: 'center' } },
      { content: 'Unit Count', colSpan: 5, styles: { halign: 'center' } }
    ];

    if (this.DupGrossType.indexOf('Front Gross') >= 0) {
      groupRow.push({ content: 'Front Gross', colSpan: 4, styles: { halign: 'center' } });
    }

    if (this.DupGrossType.indexOf('Back Gross') >= 0) {
      groupRow.push({ content: 'Back Gross', colSpan: 3, styles: { halign: 'center' } });
    }

    if (this.DupGrossType.length > 1) {
      groupRow.push({ content: 'Total Gross', colSpan: 3, styles: { halign: 'center' } });
    }

    head.push(groupRow);

    /* ---- Header Row 2 (Column headers) ---- */
    const colHeaders: any[] = [
      'Rank',
      'Salesperson',
      'Store Name',
      'New',
      'Used',
      'Total',
      'Pace',
      '90 Day Avg'
    ];

    if (this.DupGrossType.indexOf('Front Gross') >= 0) {
      colHeaders.push('Total', 'Pace', 'PVR', '90 Day Avg');
    }

    if (this.DupGrossType.indexOf('Back Gross') >= 0) {
      colHeaders.push('Total', 'Pace', 'PVR');
    }

    if (this.DupGrossType.length > 1) {
      colHeaders.push('Total', 'Pace', 'PVR');
    }

    head.push(colHeaders);

    /* ================= BODY ================= */
    const body: any[] = [];

    this.IndividualSalesPersonsData.forEach((spdata: any) => {

      const row: any[] = [
        spdata.data1 === 'Reports Total' ? '' : spdata.Rank ?? '-',
        spdata.data1 === 'Reports Total' ? '' : spdata.SalesPerson ?? '-',
        spdata.StoreName ?? '-',

        spdata.MTD_NEW,
        spdata.MTD_USED,
        spdata.MTD_Total,
        spdata.Pace,
        spdata.UnitDayAvg
      ];

      if (this.DupGrossType.indexOf('Front Gross') >= 0) {
        row.push(spdata.Gross, spdata.GrossPace, spdata.PVR, spdata.GrossDayAvg);
      }

      if (this.DupGrossType.indexOf('Back Gross') >= 0) {
        row.push(spdata.BackGross, spdata.BackGross_Pace, spdata.BackGross_PVR);
      }

      if (this.DupGrossType.length > 1) {
        row.push(spdata.TotalGross, spdata.TotalGross_Pace, spdata.TotalGross_PVR);
      }

      body.push(row);
    });

    /* ================= TABLE ================= */
    autoTable(doc, {
      startY,
      head,
      body,
      theme: 'grid',

      styles: {
        fontSize: 8,
        cellPadding: 2,
        valign: 'middle'
      },

      headStyles: {
        fillColor: [5, 84, 239],
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        halign: 'center'
      },

      didParseCell: (data: any) => {

        /* Group header row */
        if (data.section === 'head' && data.row.index === 0) {
          data.cell.styles.fillColor = [255, 255, 255];
          data.cell.styles.textColor = [0, 0, 0];
        }

        if (data.section === 'body') {

          /* Alignment */
          if ([0, 1, 2].includes(data.column.index)) {
            data.cell.styles.halign = 'left';
          } else {
            data.cell.styles.halign = 'right';
          }

          /* Dash for empty */
          if (data.cell.raw === null || data.cell.raw === undefined || data.cell.raw === 0) {
            data.cell.text = ['-'];
            data.cell.styles.halign = 'center';
          }

          /* Currency */
          if (typeof data.cell.raw === 'number' && data.column.index >= 8) {
            const val = Math.round(data.cell.raw);
            data.cell.text = [`$ ${val.toLocaleString()}`];
            if (val < 0) data.cell.styles.textColor = [255, 0, 0];
          }

          /* Alternate rows */
          if (data.row.index % 2 === 0) {
            data.cell.styles.fillColor = [245, 247, 250];
          }

          /* Reports Total / Average */
          const rowData = this.IndividualSalesPersonsData[data.row.index];
          if (rowData?.data1 === 'Reports Total' || rowData?.SalesPerson === 'Average') {
            data.cell.styles.fontStyle = 'bold';
            data.cell.styles.fillColor = [220, 230, 255];
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
      doc.save('Salesperson Rankings.pdf'); // ✅ only here
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
  generatePDFBlob(): Blob | null {
    try {
      const doc = this.createPDF()
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
      const pdfFile = this.blobToFile(pdfBlob, 'Salesperson Rankings.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Salesperson Rankings');
      formData.append('file', pdfFile);
      formData.append('notes', notes);
      formData.append('from', from);

      this.shared.api.postmethod(this.shared.common.routeEndpoint + 'mail', formData)
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
}


