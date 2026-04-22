import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Subscription } from 'rxjs';
import { environment } from '../../../../../environments/environment';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { CurrencyPipe } from '@angular/common';
import { NgbModalRef } from '@ng-bootstrap/ng-bootstrap';
import autoTable from 'jspdf-autotable';
import jsPDF from 'jspdf';
@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores],

  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  ServiceAdvisorData: any = [];
  NoData: boolean = false;

  columnName: any = 'Rank';
  columnState: any = 'asc';
  storeorgroup: any = 'G';


  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = 0;

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'Y',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname, 
  };

  FromDate: any = '';
  ToDate: any = '';
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  DupDateType: any = 'MTD';
  displaytime: any = '';
  DupFromDate: any = '';
  DupToDate: any = ''

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

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService,

  ) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      if (localStorage.getItem('flag') == 'V') {
        this.storeIds = [];
        console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.groupid
        JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store.indexOf(',') > 0 ?
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store.split(',') :
          this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store)
        localStorage.setItem('flag', 'M')
      } else {
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
        this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
      }
    }
    if (localStorage.getItem('stime') != null) {
      let stime = localStorage.getItem('stime');
      if (stime != null && stime != '') {
        this.initializeDates(stime)
      }
    } else {
      this.initializeDates('MTD')
      this.DateType = 'MTD'
    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }

    this.shared.setTitle(this.comm.titleName + '-Parts Counter Person Rankings');


    this.setHeaderData()
    this.GetData('Rank', 'asc');
  }

  ngOnInit(): void {

  }

  setHeaderData() {
    const data = {
      title: 'Parts Counter Person Rankings',
      stores: this.storeIds.toString(),
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groupId,
      storeorgroup: this.storeorgroup,

    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }
  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.DateType = type
    this.setDates(this.DateType)

  }

  tabClick(col_Name: any, Col_state: any) {
    if (this.columnName == col_Name) {
      if (Col_state == 'asc') {
        this.columnState = 'desc';
        this.GetData(this.columnName, this.columnState);
      } else {
        this.columnState = 'asc';
        this.GetData(this.columnName, this.columnState);
      }
    } else {
      if (this.storeorgroup == 'G' && (col_Name != 'Rank' && col_Name != 'ServiceAdvisor' && col_Name != 'StoreName')) {
        this.columnState = 'desc';
        this.columnName = col_Name;
        this.GetData(this.columnName, this.columnState);
      } else {
        this.columnState = 'asc';
        this.columnName = col_Name;
        this.GetData(this.columnName, this.columnState);
      }
    }
    console.log(this.columnName, this.columnState);
  }
  GetData(sortdata?: any, sortstate?: any) {

    this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    this.DupDateType = this.DateType

    this.ServiceAdvisorData = [];
    this.shared.spinner.show();
    const obj = {
      StartDate: this.FromDate,
      EndDate: this.ToDate,
      StoreID: [...this.storeIds],

      Exp: sortdata,
      OrderType: sortstate,
      RankBy: this.storeorgroup,
      UserID: 0,
    };
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetPartsCounterPersonRankings', obj)
      .subscribe(
        (res) => {
          if (res.status == 200) {
            if (res.response != undefined) {
              if (res.response.length > 0) {
                this.ServiceAdvisorData = res.response;
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
          }
          else {
            this.toast.show(res.status, 'danger', 'error');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        },
        (error) => {
          // this.toast.error('502 Bad Gate Way Error', '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }

  openDetails(Item: any) {
    // const DetailsPartsPerson = this.shared.ngbmodal.open(PartsGrossDetails, { size: 'xl', backdrop: 'static', });
    // DetailsPartsPerson.componentInstance.Partsdetails = [
    //   {
    //     StartDate: this.FromDate,
    //     EndDate: this.ToDate,
    //     var1: 'DealerName',
    //     var2: 'AP_CounterPerson',
    //     var3: '',
    //     var1Value: Item.StoreName,
    //     var2Value: Item.CounterPerson,
    //     var3Value: '',
    //     userName: Item.CounterPerson,
    //     type: Item.type,
    //     Labortype: 'C,T,I',
    //     Saletype: 'R,W',
    //     Store: Item.StoreName,
    //     PartsSource: ''
    //   },
    // ];
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  SARstate: any;
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Parts Counter Person Rankings') {
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
      if (!obj || obj.title !== 'Parts Counter Person Rankings') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Parts Counter Person Rankings') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Parts Counter Person Rankings') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Parts Counter Person Rankings') return;
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
      'type': 'M', 'others': 'Y'
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
  updatedDates(data: any) {
    console.log(data);
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

  storeorgroups(block: any, e: any) {
    this.storeorgroup = [];
    this.storeorgroup.push(e)

  }
  viewreport() {
    this.activePopover = -1
    if (this.DateType == '') {
      this.toast.show("Please select any Datetype", 'warning', 'Warning');
    }
    else if (this.storeIds.length == 0 ) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    }
    else {
      this.setHeaderData()
      this.GetData(this.columnName, this.columnState);

    }
  }

  activePopover: number = -1;
  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }
  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds

    storeNames = this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.filter((item: any) =>
      store.some((cat: any) => cat === item.ID.toString())
    );
    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.length) {
      this.ExcelStoreNames = 'All Stores'
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }
    const ServiceAdvisorData = this.ServiceAdvisorData.map(
      (_arrayElement: any) => Object.assign({}, _arrayElement)
    );
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Parts Counter Person Rankings');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 13, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A14', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Parts Counter Person Rankings']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');
    worksheet.addRow('');
    const DateToday = this.shared.datePipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.shared.datePipe.transform(
      new Date(),
      'MMddyyyy'
    );
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };

    const ReportFilter = worksheet.addRow(['Report Controls :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const Timeframe = worksheet.addRow(['Timeframe :']);
    Timeframe.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const timeframe = worksheet.getCell('B6');
    timeframe.value = this.FromDate + ' to ' + this.ToDate;
    timeframe.font = { name: 'Arial', family: 4, size: 9 };
    const Groups = worksheet.getCell('A7');
    Groups.value = 'Group :';
    Groups.font = { name: 'Arial', family: 4, size: 9, bold: true };
    const groups = worksheet.getCell('B7');
    groups.value =
      this.comm.groupsandstores.filter((val: any) => val.sg_id == this.groupId.toString())[0].sg_name;
    // this.groups == 1 ? this.comm.excelName : (this.groups == 2 ? 'Domestic' : (this.groups == 3 ? 'Import': (this.groups == 4 ? 'Warehouse': '-')));
    groups.font = { name: 'Arial', family: 4, size: 9 };
    groups.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.mergeCells('B8', 'K10');
    const Stores = worksheet.getCell('A8');
    Stores.value = 'Stores :'
    const stores = worksheet.getCell('B8');
    stores.value = this.ExcelStoreNames == 0
      ? '-'
      : this.ExcelStoreNames == null
        ? '-'
        : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores.font = { name: 'Arial', family: 4, size: 9 };
    stores.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    Stores.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    worksheet.addRow('');
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 3 && rowNumber < 11) {
        // Apply styles, formatting, or other actions for even rows
        row.getCell(1).alignment = {
          vertical: 'middle', horizontal: 'left', indent: 1
        }
      }
    });
    worksheet.mergeCells('A12', 'C12');
    let MTD = worksheet.getCell('A12');
    MTD.value = 'MTD';
    MTD.alignment = { vertical: 'middle', horizontal: 'center' };
    MTD.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    MTD.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    MTD.border = { right: { style: 'thin' } };

    worksheet.mergeCells('D12', 'I12');
    let UnitCredit = worksheet.getCell('D12');
    UnitCredit.value = '';
    UnitCredit.alignment = { vertical: 'middle', horizontal: 'center' };
    UnitCredit.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    UnitCredit.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    UnitCredit.border = { right: { style: 'thin' } };

    let Headings = [
      'Rank',
      'Counter person',
      'Store Name',
      'Total Sale',
      'Total Gross',
      'GP%',
      'Invoice Count',
      'Parts Count',
      'Parts per Invoice',
    ];
    const headerRow = worksheet.addRow(Headings);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: false,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'center' };
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '788494' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    for (const d of ServiceAdvisorData) {
      const Data = worksheet.addRow([
        d.Rank == '' ? '-' : d.Rank == null ? '-' : d.Rank,
        d.CounterPerson == ''
          ? '-'
          : d.CounterPerson == null
            ? '-'
            : d.CounterPerson,
        d.StoreName == '' ? '-' : d.StoreName == null ? '-' : d.StoreName,
        d.TotalSale == '' ? '-' : d.TotalSale == null ? '-' : d.TotalSale,
        d.TotalGross == '' ? '-' : d.TotalGross == null ? '-' : d.TotalGross,
        d.GP == '' ? '-' : d.GP == null ? '-' : d.GP,
        d.InvoiceCount == '' ? '-' : d.InvoiceCount == null ? '-' : d.InvoiceCount,
        d.AP_Numberofparts == '' ? '-' : d.AP_Numberofparts == null ? '-' : d.AP_Numberofparts,
        d.PartsPerInvoice == '' ? '-' : d.PartsPerInvoice == null ? '-' : d.PartsPerInvoice,
      ]);
      // Data1.outlineLevel = 1; // Grouping level 1
      Data.font = { name: 'Arial', family: 4, size: 8 };
      Data.alignment = { vertical: 'middle', horizontal: 'center' };
      Data.getCell(1).alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'center',
      };
      Data.eachCell((cell, number) => {
        cell.border = { right: { style: 'thin' } };
        if (number > 6 && number < 9) {
          cell.numFmt = '#,##0';
          cell.alignment = { vertical: 'middle', horizontal: 'right', indent: 1 };
        }
        if (number == 9) {
          cell.numFmt = '#,##0.0';
          cell.alignment = { vertical: 'middle', horizontal: 'right', indent: 1 };
        }
        if (number > 3 && number < 6) {
          cell.numFmt = '$#,##0.0';
          cell.alignment = { vertical: 'middle', horizontal: 'right', indent: 1 };
        } if (number == 6) {
          cell.numFmt = '0.0%';
        }
      });
      if (Data.number % 2) {
        Data.eachCell((cell, number) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e5e5e5' },
            bgColor: { argb: 'FF0000FF' },
          };
        });
      }
    }
    worksheet.getColumn(1).width = 15;
    worksheet.getColumn(2).width = 30;
    worksheet.getColumn(3).width = 30;
    worksheet.getColumn(4).width = 15;
    worksheet.getColumn(5).width = 15;
    worksheet.getColumn(6).width = 15;
    worksheet.getColumn(7).width = 15;
    worksheet.getColumn(8).width = 15;
    worksheet.getColumn(9).width = 15;
    worksheet.getColumn(10).width = 15;
    worksheet.getColumn(11).width = 15;
    worksheet.getColumn(12).width = 15;
    worksheet.getColumn(13).width = 15;
    worksheet.getColumn(14).width = 15;
    worksheet.addRow([]);

    this.shared.exportToExcel(workbook, 'Parts Counter Person Rankings_' + DATE_EXTENSION);

  }



  private createPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });

    doc.setFontSize(14);
    doc.text('Parts Counter Person Rankings', 14, 12);

    const periodTitle = this.DupDateType == "MTD" ? "MTD" : this.DupDateType == "YTD" ? "YTD" : this.DupDateType == "QTD" ? "QTD" : this.DupDateType == "LM" ? "LAST MONTH" : this.DupDateType == "LY" ? "LAST YEAR" : `${this.DupFromDate} - ${this.DupToDate}`;
    const headers = [
      [
        { content: periodTitle, colSpan: 3 }, // only for first 3 columns
        { content: '', colSpan: 9 },
      ],
      [
        'Rank', 'Counter Person', 'Store Name',
        'Total Sale', 'Total Gross', 'GP %',
        'Invoice Count', 'Parts Count', 'Parts Per Invoice'
      ]
    ];

    // ✅ Rows Mapping
    const rows = this.ServiceAdvisorData.map((x: any) => [
      x.Rank,
      x.CounterPerson,
      x.StoreName,
      x.TotalSale ?? 0,
      x.TotalGross ?? 0,
      x.GP ?? '-',
      x.InvoiceCount ?? '-',
      x.AP_Numberofparts ?? '-',
      x.PartsPerInvoice ?? '-'
    ]);

    autoTable(doc, {
      head: headers,
      body: rows,
      startY: 18,

      // ✅ NO BORDERS ANYWHERE
      styles: {
        fontSize: 9,
        cellPadding: 2,
        halign: 'center',
        valign: 'middle',
        lineWidth: 0 // 🔥 removes all borders
      },

      // ✅ COLUMN WIDTHS (KEEP ALIGNMENT PERFECT)
      columnStyles: {
        0: { cellWidth: 15 },
        1: { cellWidth: 50, halign: 'left' },
        2: { cellWidth: 50, halign: 'left' },

        3: { cellWidth: 24 },
        4: { cellWidth: 24 },
        5: { cellWidth: 18 },
        6: { cellWidth: 24 },

        7: { cellWidth: 25, halign: 'right' },
        8: { cellWidth: 30, halign: 'right' },
      },



      didParseCell: (data: any) => {

        // 🔵 ONLY TOP GROUP HEADER
        if (data.section === 'head' && data.row.index === 0) {
          data.cell.styles.fillColor = [30, 80, 200];
          data.cell.styles.textColor = 255;
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.halign = 'center';

          if (data.column.index !== 0) {
            data.cell.styles.lineWidth = { left: 0.5 };  // left border
            data.cell.styles.lineColor = [255, 255, 255]; // white divider
          }

          if (data.column.index === 2 || data.column.index === 6) {
            data.cell.styles.lineWidth = { right: 0.5 }; // right border for group end
            data.cell.styles.lineColor = [255, 255, 255];
          }
        }

        // ⚪ SECOND HEADER → PLAIN WHITE
        if (data.section === 'head' && data.row.index === 1) {
          data.cell.styles.fillColor = '#d9e7ff';
          data.cell.styles.textColor = 0;
          data.cell.styles.fontStyle = 'bold';

          // data.cell.styles.backGround = '#d9e7ff'
        }

        // 💰 FORMAT CURRENCY
        if (data.section === 'body') {
          const col = data.column.index;
          const value = data.cell.raw;

          if ([3, 4].includes(col)) {
            if (!isNaN(value)) {
              data.cell.text = [`$${Number(value).toLocaleString()}`];
            }
          }
        }
      }

      // ❌ NO alternateRowStyles
    });

    return doc;
  }

  generatePDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save('Parts Counter Person Rankings_Report.pdf'); // ✅ only here
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
      const pdfFile = this.blobToFile(pdfBlob, 'Parts Counter Person Rankings.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Parts Counter Person Rankings');
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

