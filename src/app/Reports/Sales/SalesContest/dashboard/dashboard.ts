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
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';


@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  SalesContest: any = [];
  NoData: any = false;
  searchQuery: any = ''



  keys: any = [];
  AsofNow: any = ''
  key: any = 'Rank'
  order: any = 'desc'

  reportOpenSub!: Subscription;
  reportGetting!: Subscription;

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

  FromDate: any = '';
  ToDate: any = '';
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';
  DupDateType: any = 'MTD'

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

  filterData: any = [
    { 'name': 'Units', 'id': 1 },
    { 'name': 'Gross', 'id': 2 },
  ]

  filtertype: any = ['Units']

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
    this.initializeDates('MTD')
    this.shared.setTitle(this.comm.titleName + '- Sales Contest');
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

    this.setDates(this.DateType)
    this.GetSalesContest();
    this.setHeaderData()
  }
  count: any = 0;


  ngOnInit(): void {
  }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    console.log(this.FromDate, this.ToDate);

    localStorage.setItem('time', type);
  }

  sort(key: any, order: any) {
    if (this.key == key) {
      if (order == 'asc') {
        this.order = 'desc'
      } else {
        this.order = 'asc'
      }
    } else {
      this.order = 'asc'
    }
    this.key = key
    // this.order = order
    this.count = 1;
    this.GetSalesContest()
    // this.GetInventorySummaryReport()
  }

  setHeaderData() {
    const data = {
      title: 'Sales Contest',
      stores: this.storeIds,
      groups: this.groupId,
      filtertype: this.filtertype,
      fromdate: this.FromDate,
      todate: this.ToDate,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }

  datetype() {
    if (this.DupDateType == 'PM') {
      return 'SP';
    }
    else if (this.DupDateType == 'C') {
      return 'C'
    }
    return this.DupDateType;
  }
  isDesc: boolean = false;
  column: string = '';

  TotalData: any = []
  exceldata: any = []
  GetSalesContest() {
    this.SalesContest = [];
    this.NoData = false;
    this.shared.spinner.show();
    this.dupFiltertype = this.filtertype
    this.DupDateType = this.DateType
    const obj = {
      "Stores": this.storeIds,
      "StartDate": this.FromDate,
      "EndDate": this.ToDate,
      "ContestType": this.filtertype.toString(),
      "Exp": this.key,
      "OrderType": this.order,
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetSalesContestV1', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          this.TotalData = []
          this.shared.spinner.hide();
          if (res.response != undefined) {
            if (res.response.length > 0) {
              let data = res.response
              this.exceldata = res.response
              this.SalesContest = data.filter((e: any) => e.StoreName != 'Totals');
              this.TotalData = data.filter((e: any) => e.StoreName == 'Totals')
              console.log(this.TotalData);

              let key = Object.keys(res.response[0]);
              this.keys = key.splice(3)
              this.AsofNow = res.response[0].ASofTime
              this.NoData = false;
            } else {
              this.NoData = true;
            }
          } else {
            this.NoData = true;
          }
        } else {
          this.toast.show(res.status, 'danger', 'Error');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        this.shared.spinner.hide();
        this.NoData = true;
      }
    );
  }

  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Sales Contest') {
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
      if (!obj || obj.title !== 'Sales Contest') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Sales Contest') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Sales Contest') return;
      if (obj.statePDF) {
        this.downloadPDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Sales Contest') return;
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

  filterselection(e: any) {
    const index = this.filtertype.findIndex((i: any) => i == e.name);
    if (index >= 0) {
      this.filtertype.splice(index, 1);
    } else {
      this.filtertype = []
      this.filtertype.push(e.name);
    }

  }
  dupFiltertype: any = ['Units']

  viewreport() {
    this.activePopover = -1

    if (this.storeIds.length == 0 || this.filtertype.length == 0) {
      if (this.storeIds.length == 0) {
        this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      }
      if (this.filtertype.length == 0) {
        this.toast.show('Please Select Atleast One Type', 'warning', 'Warning');
      }
    } else {
      this.setHeaderData()
      this.GetSalesContest()
    }
    // }

  }
  activePopover: number = -1;
  togglePopover(popoverIndex: number) {
    if (this.activePopover === popoverIndex) {
      // If the same popover is clicked, close it
      this.activePopover = -1;
    } else {
      // Open the selected popover and close others
      this.activePopover = popoverIndex;
    }
  }

  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0
  @ViewChild('scrollcent') scrollcent!: ElementRef;

  updateVerticalScroll(event: any): void {

    this.scrollCurrentposition = event.target.scrollTop
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );
  }


  ExcelStoreNames: any = [];

  exportToExcel() {

    const groupData = this.comm.groupsandstores
      .filter((v: any) => v.sg_id == this.groupId)[0];

    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Sales Contest');

    /* ================= FREEZE ================= */
    // worksheet.views = [{
    //   state: 'frozen',
    //   ySplit: 4,
    //   topLeftCell: 'A5',
    //   showGridLines: false,
    // }];

    /* ================= TITLE ================= */
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Sales Contest']);
    titleRow.font = { size: 12, bold: true };
    worksheet.mergeCells('A2:O2');

    worksheet.addRow('');
    /* ================= FILTERS ================= */

    const filters = [
      { name: 'Stores :', values: this.storedisplayname },
      { name: 'Time Frame :', values: this.FromDate + ' to ' + this.ToDate },
      { name: 'Content Type :', values: this.dupFiltertype || '-' },
    ];
    filters.forEach((val: any) => {

      const row = worksheet.addRow([val.name, val.values]);

      // 🔥 Make filter name bold
      row.getCell(1).font = { bold: true };

      // Optional: better alignment
      row.getCell(1).alignment = { horizontal: 'left' };
      row.getCell(2).alignment = { horizontal: 'left', wrapText: true };

    });

    let startIndex = 3;
  
    filters.forEach((val: any) => {
      startIndex++;
      worksheet.getCell(`A${startIndex}`).value = val.name;
      worksheet.mergeCells(`B${startIndex}:C${startIndex}`);
      worksheet.getCell(`B${startIndex}`).value = val.values;
    });
    worksheet.addRow('');
    /* ================= HEADERS ================= */

    const dateLabel = this.datetype() === 'C'
      ? `${this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy')} - ${this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy')}`
      : this.datetype();

    // 🔷 GROUP HEADER
    const groupHeader = worksheet.addRow([
      'Rank', 'Store',
      'New', '', '', '',
      'Used', '', '', '',
      'Total', '', '', '', ''
    ]);

    worksheet.mergeCells(`A${groupHeader.number}:A${groupHeader.number}`);
    worksheet.mergeCells(`B${groupHeader.number}:B${groupHeader.number}`);
    worksheet.mergeCells(`C${groupHeader.number}:F${groupHeader.number}`);
    worksheet.mergeCells(`G${groupHeader.number}:J${groupHeader.number}`);
    worksheet.mergeCells(`K${groupHeader.number}:O${groupHeader.number}`);

    // 🔷 SUB HEADER
    const subHeader = worksheet.addRow([
      '', '',
      dateLabel, 'Pace', 'Target', 'Variance',
      dateLabel, 'Pace', 'Target', 'Variance',
      dateLabel, 'Pace', 'Target', 'Variance', '% to Target'
    ]);

    /* ================= HEADER STYLING ================= */

    // Dark Blue (Row 1)
    groupHeader.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0554EF' }
      };
      cell.font = { color: { argb: 'FFFFFF' }, bold: true, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    // Light Blue (Row 2)
    subHeader.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '4584FF' }
      };
      cell.font = { color: { argb: 'FFFFFF' }, bold: true, size: 9 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    // 🔥 Header Borders
    [groupHeader, subHeader].forEach((row) => {
      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin', color: { argb: 'CCCCCC' } },
          left: { style: 'thin', color: { argb: 'CCCCCC' } },
          bottom: { style: 'thin', color: { argb: 'CCCCCC' } },
          right: { style: 'thin', color: { argb: 'CCCCCC' } }
        };
      });
    });

    /* ================= DATA ================= */

    this.exceldata.forEach((d: any, index: number) => {

      const row = worksheet.addRow([
        d.Rank ?? '-',
        d.StoreName ?? '-',

        d.New_MTD ?? '-',
        d.New_Pace ?? '-',
        d.New_Target ?? '-',
        d.New_Variance ?? '-',

        d.Used_MTD ?? '-',
        d.Used_Pace ?? '-',
        d.Used_Target ?? '-',
        d.Used_Variance ?? '-',

        d.Total_MTD ?? '-',
        d.Total_Pace ?? '-',
        d.Total_Target ?? '-',
        d.Total_Variance ?? '-',
        d.Total_Percentage != null ? `${d.Total_Percentage}%` : '-'
      ]);

      const isLastRow = index === this.exceldata.length - 1;

      row.font = { name: 'Arial', size: 9 };

      row.eachCell((cell, colNumber) => {

        /* 🔷 ALIGNMENT */
        if (colNumber === 1) {
          cell.alignment = { horizontal: 'center' }; // Rank
        } else if (colNumber === 2) {
          cell.alignment = { horizontal: 'left' };   // Store
        } else {
          cell.alignment = { horizontal: 'right' };  // Numbers
        }

        /* 🔷 GRID BORDERS */
        cell.border = {
          top: { style: 'thin', color: { argb: 'CCCCCC' } },
          left: { style: 'thin', color: { argb: 'CCCCCC' } },
          bottom: { style: 'thin', color: { argb: 'CCCCCC' } },
          right: { style: 'thin', color: { argb: 'CCCCCC' } }
        };

        /* 🔴 NEGATIVE VALUES */
        if (typeof cell.value === 'number' && cell.value < 0) {
          cell.font = { ...cell.font, color: { argb: 'FFDC3545' } };
        }

        /* 🔥 TOTAL ROW STYLE */
        if (isLastRow) {
          cell.font = { ...cell.font, bold: true };
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'DCE6F1' } // light blue
          };
        }

      });

      /* 🔷 ALTERNATE ROW COLOR */
      if (!isLastRow && row.number % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'F5F7FA' }
          };
        });
      }
    });

    /* ================= COLUMN WIDTH ================= */
    worksheet.columns.forEach((column: any) => {
      let maxLength = 10;
      column.eachCell({ includeEmpty: true }, (cell: any) => {
        const len = cell.value ? cell.value.toString().length : 10;
        if (len > maxLength) maxLength = len;
      });
      column.width = maxLength + 2;
    });

    worksheet.getColumn(1).width = 15;
    worksheet.getColumn(2).width = 30;

    /* ================= EXPORT ================= */
    workbook.xlsx.writeBuffer().then(() => {
      this.shared.exportToExcel(workbook, 'Sales Contest');
    });
  }
  private createSalesContestPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });

    // ===== TITLE =====
    doc.setFontSize(14);
    const title = `${'Sales Contest'}`;
    doc.text(title, 14, 12);

    // ===== DATE / AS OF =====
    doc.setFontSize(10);
    doc.text(`${this.AsofNow || ''}`, 14, 18);

    let startY = 22;

    // ===== SAFETY =====
    if (!this.SalesContest || this.SalesContest.length === 0) {
      doc.text('No Data Available', 14, startY);
      return doc;
    }

    // =========================
    // 🔷 HEADERS (GROUP)
    // =========================

    const headGroup = [[
      'Rank',
      'Store',
      { content: 'New', colSpan: 4 },
      { content: 'Used', colSpan: 4 },
      { content: 'Total', colSpan: 5 }
    ]];

    const dateLabel = this.datetype() === 'C'
      ? `${this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy')} - ${this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy')}`
      : this.datetype();

    const headColumns = [[
      '', '',   // 🔥 MUST for Rank & Store

      dateLabel, 'Pace', 'Target', 'Variance',
      dateLabel, 'Pace', 'Target', 'Variance',
      dateLabel, 'Pace', 'Target', 'Variance', '% to Target'
    ]];

    // =========================
    // 🔷 BODY
    // =========================

    const rows: any[] = [];

    (this.SalesContest || []).forEach((d: any, i: number) => {

      const row = [
        d.Rank ?? '-',
        d.StoreName ?? '-',

        d.New_MTD ?? '-',
        d.New_Pace ?? '-',
        d.New_Target ?? '-',
        d.New_Variance ?? '-',

        d.Used_MTD ?? '-',
        d.Used_Pace ?? '-',
        d.Used_Target ?? '-',
        d.Used_Variance ?? '-',

        d.Total_MTD ?? '-',
        d.Total_Pace ?? '-',
        d.Total_Target ?? '-',
        d.Total_Variance ?? '-',
        d.Total_Percentage != null ? `${d.Total_Percentage}%` : '-'
      ];

      rows.push(row);
    });

    // =========================
    // 🔷 TABLE
    // =========================

    autoTable(doc, {
      startY,
      head: [...headGroup as any, ...headColumns as any],
      body: rows,
      theme: 'grid',

      styles: {
        fontSize: 7,
        cellPadding: 1.5,
        halign: 'right',
        lineWidth: 0.3,
        lineColor: [200, 200, 200]
      },
      headStyles: {
        lineWidth: 0.3,
        lineColor: [200, 200, 200]
      },
      columnStyles: {
        0: { halign: 'center' }, // Rank
        1: { halign: 'left' }    // Store
      },

      didParseCell: (data: any) => {

        /* 🔵 HEADER COLORS */
        if (data.section === 'head') {

          // Row 1 → Dark Blue
          if (data.row.index === 0) {
            data.cell.styles.fillColor = [5, 84, 239];
            data.cell.styles.textColor = 255;
            data.cell.styles.halign = 'center';
          }

          // Row 2 → Light Blue (FULL WIDTH)
          if (data.row.index === 1) {
            data.cell.styles.fillColor = [69, 132, 255];
            data.cell.styles.textColor = 255;
            data.cell.styles.halign = 'center';
          }
        }

        /* 🔴 BODY */
        if (data.section === 'body') {

          // % column
          if (data.column.index === 14) {
            data.cell.styles.halign = 'right';
          }

          // negative values
          const val = parseFloat(data.cell.raw);
          if (!isNaN(val) && val < 0) {
            data.cell.styles.textColor = [220, 53, 69];
          }

          // 🔥 LAST ROW BOLD (like Total row in HTML)
          if (data.row.index === this.SalesContest.length - 1) {
            data.cell.styles.fontStyle = 'bold';
          }
        }
      }
    });

    return doc;
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }
  downloadPDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createSalesContestPDF();
      doc.save('Sales Contest.pdf');
    } catch (e) {
      console.error(e);
      this.toast.show('Error generating PDF', 'danger', 'Error');
    } finally {
      this.shared.spinner.hide();
    }
  }
  printPDF() {
    try {
      const doc = this.createSalesContestPDF();
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
      const doc = this.createSalesContestPDF()
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
      const pdfFile = this.blobToFile(pdfBlob, 'Sales Contest.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Sales Contest');
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
