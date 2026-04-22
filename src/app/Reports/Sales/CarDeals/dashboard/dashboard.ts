import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Component, HostListener, } from '@angular/core';

import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { CardealsReports } from '../cardeals-reports/cardeals-reports';
import { Subscription } from 'rxjs';
import { Notes } from '../../../../Layout/notes/notes';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';

import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';




@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, CardealsReports, Notes],
  standalone: true,
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  Viewmore: boolean = false
  spinnerLoader: boolean = false
  notesViewState: boolean = false
  FromDate: any = '';
  ToDate: any = '';
  TotalReport: any = 'T';
  NoData!: boolean;
  CompleteComponentState: boolean = true;
  store: any = '1';
  QISearchName: any = '';
  path1: any = '';
  path2: any = '';
  path3: any = '';
  path1name: any = '';
  path2name: any = '';
  path3name: any = '';
  path1id: any;
  path2id: any;
  path3id: any;
  CurrentDate = new Date();
  groups: any = 0;
  GridView = 'Global';
  dealType: any = ['New', 'Used'];
  saleType: any = ['Retail', 'Lease', 'Misc', 'Special Order']
  dealStatus: any = ['Delivered', 'Capped', 'Finalized'];
  acquisition: any = ['All'];
  count = 0;
  storeName: any = ''
  DateType: any = 'MTD'
  RoleId: any = '';
  routepath: any = ''

  header: any = [{
    type: 'Bar', storeIds: this.store, fromDate: this.FromDate, toDate: this.ToDate, groups: this.groups, datevaluetype: this.DateType
  }]
  pageNumber: any = 0;
  userData: any = {}
  constructor(public shared: Sharedservice, public setdates: Setdates,
    private toast: ToastService,
  ) {

    this.shared.setTitle(this.shared.common.titleName + '-Car Deals')

    // if (typeof window !== 'undefined') {
    // if (localStorage.getItem('userInfo') != null) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      this.groups = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.store = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
      this.RoleId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.roleid
      if (localStorage.getItem('flag') == 'V') {
        this.column = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.CN
        JSON.parse(localStorage.getItem('userInfo')!).WidgetData.DS ? this.dealStatus = ['Finalized'] : '';
        this.userData = JSON.parse(localStorage.getItem('userInfo')!).WidgetData
        localStorage.setItem('flag', 'M')

      }
    }
    if (localStorage.getItem('CarDeals') != undefined && localStorage.getItem('CarDeals') != null) {
      let fromIBData = JSON.parse(localStorage.getItem('CarDealsData')!)[0]
      console.log(fromIBData);
      this.store = fromIBData.storeid;
      this.groups = fromIBData.groups;
      this.dealType = fromIBData.stocktype;
      this.FromDate = fromIBData.fromDate;
      this.ToDate = fromIBData.toDate;
      this.QISearchName = fromIBData.model ? fromIBData.model : '';
      this.acquisition = fromIBData.acquisition ? fromIBData.acquisition : ['All'];
      this.routepath = fromIBData.routepath
      this.DateType = fromIBData.datetype
      localStorage.setItem('time', this.DateType);
    } else {
      this.setDates('MTD')
      this.DateType = 'MTD'
    }

    // }
    // if (localStorage.getItem('Fav') != 'Y') {
    this.setHeaderReportData()
    // }
    // }
  }

  ngOnInit(): void {
    // localStorage.setItem('time', 'MTD');
    // this.reportOpenSub.unsubscribe()
    // this.reportGetting.unsubscribe()
    // this.Pdf.unsubscribe()
    // this.print.unsubscribe()
    // this.email.unsubscribe()
    // this.excel.unsubscribe()
  }
  details: any = [];
  otherblocks: any = ['Retail and Lease']
  obTemporary: any = 'Retail and Lease'
  callLoadingState = 'FL'
  cardealsdata: any = []
  getCardealsFilter() {
    this.details = []
    this.cardealsdata = []
    this.GetData()
  }

  setDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
  }
  GetData() {
    this.shared.spinner.show();
    this.NoData=false
    console.log(this.details, this.cardealsdata);
    this.spinnerLoader = true
    const obj = {
      startdealdate: this.FromDate,
      enddealdate: this.ToDate,
      Stores: this.store,
      dealtype: this.dealType.toString() == 'C' ? 'New,Used' : this.dealType.toString(),
      saletype: this.saleType.toString(),
      dealstatus: this.dealStatus.toString(),
      PageNumber: this.pageNumber,
      PageSize: '1000',
      ExtraType: this.otherblocks.toString(),
      SearchExp: this.QISearchName,
      AcquisitionSource: this.acquisition.toString() == 'All' ? '' : this.acquisition.toString(),
      SalesPersonID: this.RoleId == 121 ? JSON.parse(localStorage.getItem('userInfo')!).user_Info?.EmpID : ''
    };
    this.count++
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetSalesCarDeals', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          this.shared.spinner.hide()
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.details = res.response;
              this.details.some(function (x: any) {
                if (x.Notes != undefined && x.Notes != null) {
                  x.Notes = JSON.parse(x.Notes);
                }
              });
              this.cardealsdata = [
                ...this.cardealsdata,
                ...this.details,
              ];
              this.callLoadingState == 'ANS' ? this.sort(this.column, this.callLoadingState) : JSON.parse(localStorage.getItem('userInfo')!).WidgetData.DS ? (this.sort(this.column, this.userData?.SR)) : '';

              this.NoData = false;
              this.spinnerLoader = false
              console.log(this.details, this.cardealsdata);
            }
            else {
              // this.toast.error('Empty Response','');
              this.details = []
              this.shared.spinner.hide();
              if (this.cardealsdata.length == 0) {
                this.NoData = true;
              }
              this.spinnerLoader = false

            }
          } else {
            // this.toast.error('Empty Response');
            this.shared.spinner.hide();
            if (this.cardealsdata.length == 0) {
              this.NoData = true;
            }
            this.spinnerLoader = false
          }
        } else {
          this.spinnerLoader = false
          this.shared.spinner.hide();
          this.NoData = true;

        }
      },
      (error) => {

        this.shared.spinner.hide();

      }
    );
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }

  // getTotal(frontgross: any, colname: any) {
  //   let total: any = 0
  //   frontgross.some(function (x: any) {
  //     total += parseInt(x[colname])
  //   })
  //   return total
  // }

  getTotal(frontgross: any[], colname: string) {
    return frontgross.reduce((total, x) => {
      const raw = x?.[colname];

      // Skip null/undefined/empty strings
      if (raw === null || raw === undefined || raw === '') return total;

      // Parse as number (handles strings like "123" or "123.45")
      const n = Number(raw);

      // Skip NaN
      if (Number.isNaN(n)) return total;

      // Add the value (use Math.trunc if you only want integer part)
      return total + n;
    }, 0);
  }


  notesView() {
    this.notesViewState = !this.notesViewState
  }
  notesData: any = {}
  Notespopup: any;
  scrollpositionstoring: any;
  addNotes(data: any, ref: any) {
    // this.scrollpositionstoring = this.scrollCurrentposition
    // this.notesData = {
    //   store: data.dealerid,
    //   mainkey: data.dealid,
    //   module: 'CD',
    //   notes: data.Notes,
    //   apiRoute: 'AddGeneralNotes'
    // }

    this.notesData = {
      store: data.dealerid,
      title1: data.dealid,
      title2: '',
      apiRoute: 'AddGeneralNotes',
       moduleCode: 'CD'
    }
    this.Notespopup = this.shared.ngbmodal.open(ref, { size: 'lg', backdrop: 'static' });
  }
  closeNotes(e: any) {
    this.shared.ngbmodal.dismissAll()
    if (e == 'S') {
      this.details = []
      this.cardealsdata = [];
      this.callLoadingState = 'ANS'
      this.GetData()
    }
  }
  isDesc: boolean = false;
  column: string = 'CategoryName';

  tradeDetails: any = []
  spinnerLoaderTrade: boolean = false;
  openTradePopup(item: any) {
    this.tradeDetails = []
    this.spinnerLoaderTrade = true
    const obj = {
      "startdealdate": this.FromDate,
      "enddealdate": this.ToDate,
      "dealid": item.dealid,
      "StoreID": item.dealerid
    }
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetSalesCarDealsTradeDetails', obj).subscribe((res: any) => {
      if (res.status == 200) {
        if (res.response != undefined) {
          if (res.response.length > 0) {
            this.tradeDetails = res.response
            this.spinnerLoaderTrade = false
          } else {
            this.spinnerLoaderTrade = false
          }
        } else {
          this.spinnerLoaderTrade = false
        }
      } else {
        this.spinnerLoaderTrade = false
      }
    })
  }

  toggleView(data: any) {
    if (data.notesView == '+') {
      data.notesView = '-'
    } else {
      data.notesView = '+'
    }
  }
  sort(property: any, state?: any) {
    if (state == undefined) {
      this.isDesc = !this.isDesc;
    }
    console.log(this.isDesc, 'Desc');

    this.callLoadingState = 'FL'
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    this.cardealsdata.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
  subdataindex: any = 0;

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
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit() {

    this.shared.api.GetReports().subscribe((data) => {
      let count = data.obj.count
      if (data.obj.Reference == 'Car Deals') {
        count++
        console.log(data.obj);
        this.cardealsdata = [];
        this.pageNumber = 0
        if (data.obj.header == undefined) {
          this.TotalReport = data.obj.TotalReport;
          this.FromDate = data.obj.FromDate;
          this.ToDate = data.obj.ToDate;
          this.store = data.obj.storeValues;
          this.groups = data.obj.groups

          this.dealType = data.obj.dealType;
          this.saleType = data.obj.saleType;
          this.dealStatus = data.obj.dealStatus;
          this.otherblocks = data.obj.otherblock;
          this.obTemporary = data.obj.otherblock
          this.DateType = data.obj.dateType;
          this.acquisition = data.obj.acquisition;
          this.pageNumber = 0;
          this.QISearchName = data.obj.search;
          this.callLoadingState = 'FL'
          this.setHeaderReportData();

        }
        // this.GetData();
        else {
          if (data.obj.header == 'Yes' && data.obj.Reference == 'Car Deals') {
            count = 0
            this.store = data.obj.storeValues;
            this.pageNumber = 0
            console.log(data.obj);
            // this.GetData();
            this.setHeaderReportData()
          }
        }
      }

    });

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Car Deals') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Car Deals') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Car Deals') return;
      if (obj.statePDF) {
        this.downloadPDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Car Deals') return;
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
  setHeaderReportData() {
    this.pageNumber = 0

    this.details = [];
    this.cardealsdata = []
    this.obTemporary = this.otherblocks[0]
    const headerdata = {
      title: 'Car Deals',
      path1: this.path1name,
      path2: this.path2name,
      path3: this.path3name,
      path1id: this.path1id,
      path2id: this.path2id,
      path3id: this.path3id,
      stores: this.store,
      dealType: this.dealType,
      saleType: this.saleType,
      dealStatus: this.dealStatus,
      ToporBottom: this.TotalReport,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groups,
      otherblock: this.otherblocks,
      search: this.QISearchName,
      count: this.count, datevaluetype: this.DateType, routepath: this.routepath
    };
    this.shared.api.SetHeaderData({
      obj: headerdata,
    });
    this.header = [{
      type: 'Bar', storeIds: this.store, fromDate: this.FromDate, toDate: this.ToDate, ReportTotal: this.TotalReport, search: this.QISearchName, groups: this.groups, datevaluetype: this.DateType, dealType: this.dealType,
      saleType: this.saleType,
      dealStatus: this.dealStatus, as: this.acquisition, otherblock: this.otherblocks, routepath: this.routepath
    }]
    if (this.store != '') {
      this.GetData()
    } else {
      this.NoData = true;
      this.cardealsdata = []
    }
  }
  ExcelStoreNames: any = [];

  exportToExcel(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Car Deals');
  
    /* ================= TITLE ================= */
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Car Deals']);
    titleRow.font = { size: 14, bold: true };
    worksheet.mergeCells('A2:Z2');
    worksheet.addRow('');
    /* ================= STORE ================= */
    // const Stores1 = worksheet.getCell('A3');
    // Stores1.value = 'Stores :';
  
    // worksheet.mergeCells('B3:Z3');
    // const stores1 = worksheet.getCell('B3');
    // stores1.value = this.ExcelStoreNames?.toString().replaceAll(',', ', ');
    // stores1.font = { name: 'Arial', size: 9 };
    // stores1.alignment = { wrapText: true };
  
    /* ================= FILTERS ================= */
    const filters = [
      { name: 'Time Frame :', values: this.FromDate + ' to ' + this.ToDate },
      { name: 'New Used : ', values: this.dealType || '-' },
      { name: 'Deal Type :', values: this.saleType || '-' },
      { name: 'Deal Status :', values: this.dealStatus || '-' },
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
    /* ================= HEADER ================= */
    const headerRowIndex = startIndex + 2;
  
    const headers = [
      'Date', 'Store', 'Deal #', 'Status', 'R/L', 'Stock #',
      'New/Used', 'Year', 'Make', 'Model', 'Trade','Cost',
      'VIN', 'Age', 'F Gross', 'B Gross', 'Total',
      'Mgr', 'F&I Mgr', 'Sp 1', 'Sp 2', 'Buyer', 'Cust #'
    ];
  
    const bindingHeaders = [
      'displaydate', 'store', 'dealid', 'Status', 'RL', 'Stock', 'Type',
      'ad_year', 'ad_make', 'ad_model', 'TradeACV', 'CostPrice' , 'VIN',
      'ad_custAge', 'frontgross', 'backgross', 'totalgross',
      'salesmanager', 'fimanager', 'salesperson1', 'salesperson2',
      'Buyer', 'ad_custid'
    ];
  
    const headerRow = worksheet.addRow(headers);
  
    // headerRow.height = 25;
  
    headerRow.font = {
      name: 'Arial',
      size: 10,
      bold: true,
      color: { argb: 'FFFFFFFF' }
    };
  
    headerRow.alignment = {
      vertical: 'middle',
      horizontal: 'center',
      wrapText: true
    };
  
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF0554EF' } // PDF BLUE
    };
  
    headerRow.eachCell((cell) => {
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
      };
    });
  
    /* ================= DATA ================= */
  
    const currencyFields = ['TradeACV', 'frontgross', 'backgross', 'totalgross','CostPrice'];
  
    this.cardealsdata.forEach((info: any) => {
  
      const rowData = bindingHeaders.map((key) => {
        if (key === 'displaydate') {
          return this.shared.datePipe.transform(info[key], 'MM/dd/yyyy');
        }
        return info[key] === 0 || info[key] == null ? '-' : info[key];
      });
  
      const row = worksheet.addRow(rowData);
  
      row.eachCell((cell, colNumber) => {
  
        const key = bindingHeaders[colNumber - 1];
  
        /* Alignment */
        cell.alignment = {
          vertical: 'middle',
          horizontal: isNaN(Number(cell.value)) ? 'left' : 'center'
        };
  
        /* Font */
        cell.font = { name: 'Arial', size: 9 };
  
        /* Borders */
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
          left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
          bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
          right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
        };
  
        /* Currency Format */
        if (currencyFields.includes(key) && typeof info[key] === 'number') {
          cell.value = info[key];
          cell.numFmt = '"$"#,##0.00';
          cell.alignment = { horizontal: 'right' };
  
          /* Negative values red */
          if (info[key] < 0) {
            cell.font = { color: { argb: 'FFDC3545' } };
          }
        }
      });
  
      /* ================= NOTES ROW ================= */
      if (info.Notes && this.notesViewState) {
        const notesText = info.Notes.map((n: any) => n.GN_Text).join('\n');
  
        const notesRow = worksheet.addRow([`Notes:\n${notesText}`]);
        worksheet.mergeCells(`A${notesRow.number}:V${notesRow.number}`);
  
        const cell = notesRow.getCell(1);
  
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF0F5FF' }
        };
  
        cell.font = { size: 8, bold: true };
        cell.alignment = { wrapText: true };
      }
  
    });
  
    /* ================= COLUMN WIDTH ================= */
    worksheet.columns.forEach((column: any) => {
      let maxLength = 10;
    
      column.eachCell({ includeEmpty: true }, (cell: any) => {
        let cellValue = cell.value;
    
        if (cellValue == null) return;
    
        // Handle rich text / objects
        if (typeof cellValue === 'object') {
          cellValue = cellValue.text || cellValue.richText?.map((t: any) => t.text).join('') || '';
        }
    
        const length = cellValue.toString().length;
    
        if (length > maxLength) {
          maxLength = length;
        }
      });
    
      // Add padding
      column.width = Math.min(maxLength + 2, 50); // max limit 50
    });
    /* ================= FREEZE HEADER ================= */
    worksheet.views = [{ state: 'frozen', ySplit: headerRow.number }];
  
    /* ================= EXPORT ================= */
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'Car Deals');
    });
  }


  private createDealsPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a2' });

    // ✅ Safety check AFTER doc creation
    if (!this.cardealsdata || this.cardealsdata.length === 0) {
      doc.text('No Data Available', 14, 20);
      return doc;
    }

    // ===== TITLE =====
    doc.setFontSize(14);
    const title = `${'Car Deals'} `;
    doc.text(title, 14, 12);

    // ===== DATE =====
    const from = this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy');
    const to = this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy');

    doc.setFontSize(10);
    doc.text(`${from} - ${to}`, 14, 18);

    let startY = 22; // 🔥 adjusted because date added

    // =========================
    // 🔷 TOTALS BLOCK
    // =========================
    const totalFront = this.getTotal(this.cardealsdata, 'frontgross');
    const totalBack = this.getTotal(this.cardealsdata, 'backgross');

    autoTable(doc, {
      startY,
      theme: 'grid',

      styles: {
        fontSize: 9,
        cellPadding: 2,
        valign: 'middle',
        halign: 'center',
        lineWidth: 0.3,
        lineColor: [200, 200, 200]
      },
      head: [[
        'Deal Count',
        'Total Front Gross',
        'Total Back Gross',
        'Total Gross',
        'Front PVR',
        'Back PVR',
        'Total PVR'
      ]],
      body: [[
        this.cardealsdata.length,
        this.formatMoney(totalFront),
        this.formatMoney(totalBack),
        this.formatMoney(totalFront + totalBack),
        this.formatMoney((totalFront / this.cardealsdata.length || 0).toFixed(0)),
        this.formatMoney((totalBack / this.cardealsdata.length || 0).toFixed(0)),
        this.formatMoney(((totalFront + totalBack) / this.cardealsdata.length || 0).toFixed(0)),
      ]],

      headStyles: {
        fillColor: [5, 84, 239],
        textColor: 255,
        fontStyle: 'bold',
        halign: 'center',
        lineWidth: 0.3,
        lineColor: [200, 200, 200]
      },

      columnStyles: {
        0: { halign: 'center' },
        1: { halign: 'right' },
        2: { halign: 'right' },
        3: { halign: 'right' },
        4: { halign: 'right' },
        5: { halign: 'right' },
        6: { halign: 'right' }
      }
    });
    startY = (doc as any).lastAutoTable.finalY + 6;

    // =========================
    // 🔷 MAIN TABLE (Retail/Lease)
    // =========================

    const head = [[
      '#', 'Date', 'Store', 'Deal #', 'Status', 'R/L',
      'Stock #', 'Type', 'Year', 'Make', 'Model',
      'Trade', 'Cost','VIN', 'Age',
      'F Gross', 'B Gross', 'Total',
      'Mgr', 'F&I Mgr', 'Sp1', 'Sp2', 'Buyer', 'Cust #'
    ]];

    const rows: any[] = [];

    this.cardealsdata.forEach((d: any, i: number) => {

      const row = [
        i + 1,
        d.contractdate ? this.formatDateMMDDYYYY(d.contractdate) : '-',
        d.store || '-',
        d.dealid || '-',
        d.Status || '-',
        d.RL || '-',
        d.Stock || '-',
        d.Type || '-',
        d.ad_year || '-',
        d.ad_make || '-',
        d.ad_model || '-',
        this.formatMoney(d.TradeACV),
        this.formatMoney(d.CostPrice),        
        d.VIN || '-',
        d.ad_custAge || '-',
        this.formatMoney(d.frontgross),
        this.formatMoney(d.backgross),
        this.formatMoney(d.totalgross),
        d.salesmanager || '--',
        d.fimanager || '--',
        d.salesperson1 || '--',
        d.salesperson2 || '--',
        d.Buyer || '--',
        d.ad_custid || '-'
      ];

      rows.push(row);

      // 🔥 NOTES ROW (same like HTML)
      if (d.Notes && this.notesViewState) {
        rows.push([{
          content: 'Notes:\n' + d.Notes.map((n: any) => n.GN_Text).join('\n'),
          colSpan: head[0].length,
          styles: {
            fillColor: [240, 245, 255],
            textColor: [0, 0, 0],
            fontSize: 8
          }
        }]);
      }
    });

    autoTable(doc, {
      startY,
      head,
      body: rows,
      theme: 'grid',
      styles: {
        fontSize: 8,
        cellPadding: 1.5,
        overflow: 'linebreak',
        lineWidth: 0.3,
        lineColor: [200, 200, 200]
      },
      headStyles: {
        fillColor: [5, 84, 239],
        textColor: 255,
        lineWidth: 0.3,
        lineColor: [200, 200, 200]
      },
      didParseCell: (data: any) => {

        if ([11,12, 15, 16, 17].includes(data.column.index)) {
          const val = parseFloat(data.cell.raw?.toString().replace(/[$,]/g, ''));

          data.cell.styles.halign = 'right';

          if (val < 0) {
            data.cell.styles.textColor = [220, 53, 69];
          }
        }

        // Notes styling
        if (data.cell.raw?.content?.startsWith('Notes')) {
          data.cell.styles.fontStyle = 'bold';
        }
      }
    });

    // =========================
    // 🔷 CUSTOMER TABLE (SECOND VIEW)
    // =========================

    if (this.obTemporary === 'Customer') {

      startY = (doc as any).lastAutoTable.finalY + 10;

      const head2 = [[
        'Store', 'Date', 'Deal #', 'Status', 'R/L',
        'Stock', 'Type', 'Year', 'Make', 'Model',
        'Age', 'Total', 'Sale Price',
        'Sp1', 'Sp2', 'Cust #', 'Buyer',
        'Buyer Age', 'CoBuyer', 'CB Age',
        'City', 'Zip', 'Mileage'
      ]];

      const rows2 = this.cardealsdata.map((d: any) => [
        d.store || '-',
        d.displaydate || '-',
        d.dealid || '-',
        d.Status || '-',
        d.RL || '-',
        d.Stock || '-',
        d.ad_dealtype || '-',
        d.ad_year || '-',
        d.ad_make || '-',
        d.ad_model || '-',
        d.ad_custAge || '-',
        this.formatMoney(d.Total),
        this.formatMoney(d.CostPrice),
        d.salesperson1 || '--',
        d.salesperson2 || '--',
        d.ad_custid || '-',
        d.Buyer || '--',
        d.Buyerage || '-',
        d.ad_coname1 || '--',
        d.CoBuyerage || '-',
        d.ad_city || '-',
        d.ad_custZip || '-',
        d.ad_trade1mileage || '-'
      ]);

      autoTable(doc, {
        startY,
        head: head2,
        body: rows2,
        theme: 'grid',
        styles: { fontSize: 8 }
      });
    }

    return doc;
  }

  formatMoney(value: any): string {
    if (value === null || value === undefined || value === '' || value === 0) {
      return '-';
    }

    const num = Number(value);
    if (isNaN(num)) return '-';

    return `$${num.toLocaleString('en-US', {
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    })}`;
  }

  formatDateMMDDYYYY(date: any): string {
    if (!date) return '-';
    return this.shared.datePipe.transform(date, 'MM.dd.yyyy') || '-';
  }

  downloadPDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createDealsPDF();   // ✅ IMPORTANT CHANGE
      doc.save(`Car Deals.pdf`);
    } catch (e) {
      console.error(e);
      this.toast.show('Error generating PDF', 'danger', 'Error');
    } finally {
      this.shared.spinner.hide();
    }
  }


  printPDF() {
    try {
      const doc = this.createDealsPDF();
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
      const doc = this.createDealsPDF()
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
      const pdfFile = this.blobToFile(pdfBlob, 'Car Deals.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Car Deals');
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