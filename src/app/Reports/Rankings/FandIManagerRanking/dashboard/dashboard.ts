import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Component, ViewChild, ElementRef, HostListener, SimpleChanges } from '@angular/core';
// import { Subscription } from 'rxjs';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Subscription } from 'rxjs';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { SalesgrossDetails } from '../../../Sales/SalesGross/salesgross-details/salesgross-details';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, Stores, DateRangePicker],
  standalone: true,
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {

  FIManagerData: any = [];
  NoData: boolean = false;

  storeorgrp: any = 'G';
  groups: any = 0;
  storeIds: any = '0';
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;




  neworused: string[] = ['New', 'Used'];
  storeorgroup: any = ['G'];
  retailorlease: any = ['Retail', 'Lease', 'Misc', 'Special Order'];
  financetype: any = ['Finance', 'Cash', 'Lease'];


  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'Y',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname,
   

  };

  FromDate: any = '';
  ToDate: any = '';
  DupFromDate: any = '';
  DupToDate: any = ''
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  DupDateType: any = 'MTD';

  displaytime: any = '';

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


  columnName: any = 'Rank';
  columnState: any = 'asc';


  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common, private toast: ToastService,
  ) {
    this.shared.setTitle(this.shared.common.titleName + 'F&I Manager Ranking');
    if (typeof window !== 'undefined') {
      if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
        if (localStorage.getItem('flag') == 'V') {
          this.storeIds = [];
          this.groupId = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.groupid
          JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store.indexOf(',') > 0 ?
            this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store.split(',') :
            this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store)
          localStorage.setItem('flag', 'M')
        } else {
          this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',');
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
      this.setHeaderData()
      this.GetData('Rank', 'asc');
    }


  }

  ngOnInit(): void {

  }

  setHeaderData() {
    const data = {
      title: 'F&I Manager Rankings',
      stores: this.storeIds,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groups,
      storeorgroup: this.storeorgrp,
      dealType: this.neworused,
      saleType: this.retailorlease,
      financeType: this.financetype,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });

  }


  GetData(sortdata?: any, sortstate?: any) {
    this.FIManagerData = [];
    this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    this.DupDateType = this.DateType
    this.shared.spinner.show();
    const obj = {
      StartDate: this.FromDate,
      EndDate: this.ToDate,
      StoreID: [...this.storeIds],
      Exp: sortdata,
      OrderType: sortstate,
      RankBy: this.storeorgroup,
      UserID: 0,
      SaleType: this.neworused.toString(),
      DealType: this.retailorlease.toString(),
      FinanceType: this.financetype.toString(),
    };
    let startFrom = new Date().getTime();
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetFIManagerRankings', obj).subscribe(
      (res: { message: any; status: number; response: string | any[] | undefined; }) => {
        if (res.status == 200) {
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.FIManagerData = res.response;
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
      (error: any) => {
        this.shared.spinner.hide();
        this.NoData = true;
      }
    );
  }
  openDetails(data: any) {
    const DetailsSalesPeron = this.shared.ngbmodal.open(SalesgrossDetails, { size: 'xxl', backdrop: 'static', windowClass: 'SalesDetails' });
    DetailsSalesPeron.componentInstance.Salesdetails = [
      {
        StartDate: this.FromDate,
        EndDate: this.ToDate,
        dealtype: this.neworused,
        saletype: this.retailorlease,
        dealstatus: "Delivered,Capped,inalized",
        var1: 'store',
        var2: 'fimanager',
        var3: '',
        var1Value: data.StoreName,
        var2Value: data.fimanager,
        var3Value: '',
        userName: data.fimanager,
        FinanceManager: "0",
        SalesManager: "0",
        SalesPerson: "0"
      },
    ];
    DetailsSalesPeron.result.then((data: any) => { }, (reason: any) => { });
  }


  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }


  FIMstate: any;
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
      if (this.storeorgrp == 'G' && (col_Name != 'Rank' && col_Name != 'fimanager' && col_Name != 'StoreName')) {
        this.columnState = 'desc';
        this.columnName = col_Name;
        this.GetData(this.columnName, this.columnState);
      } else {
        this.columnState = 'asc';
        this.columnName = col_Name;
        this.GetData(this.columnName, this.columnState);
      }

    }
    // //console.log(this.columnName, this.columnState);
  }

  excel!: Subscription;
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'F&I Manager Rankings') {
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
      if (!obj || obj.title !== 'F&I Manager Rankings') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'F&I Manager Rankings') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'F&I Manager Rankings') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'F&I Manager Rankings') return;
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

  updatedDates(data: any) {
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

  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;

  }


  private createPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });

    doc.setFontSize(14);
    doc.text('F&I Manager Rankings', 14, 12);

    const periodTitle = this.DupDateType == "MTD" ? "MTD" : this.DupDateType == "YTD" ? "YTD" : this.DupDateType == "QTD" ? "QTD" : this.DupDateType == "LM" ? "LAST MONTH" : this.DupDateType == "LY" ? "LAST YEAR" : `${this.DupFromDate} - ${this.DupToDate}`;
    const headers = [
      [
        { content: periodTitle, colSpan: 3 },
        { content: 'UNIT COUNT', colSpan: 4 },
        { content: 'BACK GROSS', colSpan: 3 }
      ],
      [
        'RANK',
        'FINANCE MANAGER',
        'STORE NAME',

        'NEW',
        'USED',
        'TOTAL',
        'PACE',

        'GROSS',
        'PACE',
        'PVR'
      ]
    ];

    // ✅ DATA
    const rows = this.FIManagerData.map((x: any) => [
      x.Rank,
      x.fimanager,
      x.StoreName,

      x.New ?? 0,
      x.Used ?? 0,
      x.Total ?? 0,
      x.Pace ?? 0,

      x.BackGross ?? 0,
      x.BackGrossPace ?? 0,
      x.BackgrossPVR ?? 0
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
        lineWidth: 0.3,
        lineColor: [200, 200, 200]
      },
      headStyles: {
        lineWidth: 0.3,
        lineColor: [200, 200, 200]
      },

      columnStyles: {
        0: { cellWidth: 'auto' },
        1: { cellWidth: 'auto', halign: 'left' },
        2: { cellWidth: 'auto', halign: 'left' },
        3: { cellWidth: 'auto' },
        4: { cellWidth: 'auto' },
        5: { cellWidth: 'auto' },
        6: { cellWidth: 'auto' },
        7: { cellWidth: 'auto', halign: 'right' },
        8: { cellWidth: 'auto', halign: 'right' },
        9: { cellWidth: 'auto', halign: 'right' }
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

          if ([7, 8, 9].includes(col)) {
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
      doc.save('F&I Manager Rankings.pdf'); // ✅ only here
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
      const pdfFile = this.blobToFile(pdfBlob, 'F&I Manager Rankings.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'F&I Manager Rankings');
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


  activePopover: number = -1;
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



  multipleorsingle(block: string, val: string) {
    if (block === 'NU') {
      this.toggleSelection(this.neworused, val);
    }
    if (block === 'RL') {
      this.toggleSelection(this.retailorlease, val);
    }
    if (block === 'DS') {
      this.toggleSelection(this.financetype, val);
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


  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.DateType = type
    this.setDates(this.DateType);
  }



  close() {
    this.shared.ngbmodal.dismissAll();
  }


  viewreport() {
    this.activePopover = -1;
    if (this.storeIds.length === 0 ) {
      this.toast.show('Please select any store', 'warning', 'Warning');
    }
    else if (this.retailorlease.length == 0) {
      this.toast.show('Please select any one Deal Type', 'warning', 'Warning');
    }
    else if (this.neworused.length == 0) {
      this.toast.show('Please Select Atleast One Sale Type', 'warning', 'Warning');
    }
    else if (this.financetype.length == 0) {
      this.toast.show('Please Select Atleast One Finance Type', 'warning', 'Warning');
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
    const worksheet = workbook.addWorksheet('F&I Manager Rankings');

    /* ================= TITLE ================= */
    const title = worksheet.addRow(['F&I Manager Rankings']);
    title.font = { size: 14, bold: true };
    worksheet.mergeCells('A1:J1');

    worksheet.addRow([]);

    /* ================= FILTERS ================= */

    const formattedFromDate = this.shared.datePipe.transform(this.FromDate, 'dd-MMM-yyyy');
    const formattedToDate = this.shared.datePipe.transform(this.ToDate, 'dd-MMM-yyyy');

    const selectedStoreIds: string[] =
      this.storeIds?.length ? this.storeIds.map((id: any) => id.toString()) : [];

    const storeValue = (this.stores || [])
      .filter((s: any) => selectedStoreIds.includes(s.ID.toString()))
      .map((s: any) => s.storename.trim())
      .join(', ') || selectedStoreIds.join(', ');

    const filters = [
      { name: 'Stores:', values: storeValue },
      { name: 'Time Frame:', values: `${formattedFromDate} to ${formattedToDate}` },
      { name: 'Rank By:', values: this.storeorgroup == 'S' ? 'Store' : 'Group' },
      { name: 'New/Used:', values: this.neworused || 'All' },
      { name: 'Deal Type:', values: this.retailorlease || 'All' },
      { name: 'Finance Type:', values: this.financetype || 'All' },
    ];

    filters.forEach((val: any) => {

      const row = worksheet.addRow([val.name, val.values]);

      // Merge values column
      worksheet.mergeCells(`B${row.number}:J${row.number}`);

      // 🔥 Make filter name bold
      row.getCell(1).font = { bold: true };

      // Optional: better alignment
      row.getCell(1).alignment = { horizontal: 'left' };
      row.getCell(2).alignment = { horizontal: 'left', wrapText: true };

    });

    worksheet.addRow([]);

    /* ================= HEADERS ================= */

    const headerRow1 = worksheet.addRow([
      'MTD', '', '',
      'UNIT COUNT', '', '', '',
      'BACK GROSS', '', ''
    ]);

    const headerRow2 = worksheet.addRow([
      'RANK', 'FINANCE MANAGER', 'STORE NAME',
      'NEW', 'USED', 'TOTAL', 'PACE',
      'GROSS', 'PACE', 'PVR'
    ]);

    // Merge dynamically (NO hardcoding ❌)
    worksheet.mergeCells(`A${headerRow1.number}:C${headerRow1.number}`);
    worksheet.mergeCells(`D${headerRow1.number}:G${headerRow1.number}`);
    worksheet.mergeCells(`H${headerRow1.number}:J${headerRow1.number}`);

    /* ================= HEADER STYLE ================= */

    // Top Header (Dark Blue)
    headerRow1.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1E50C8' }
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    // Second Header (Light Blue)
    headerRow2.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9E7FF' }
      };
      cell.font = { bold: true, color: { argb: 'FF000000' }, size: 10 };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    const applyBorder = (row: any) => {
      row.eachCell((cell: any) => {
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFC8C8C8' } },
          left: { style: 'thin', color: { argb: 'FFC8C8C8' } },
          bottom: { style: 'thin', color: { argb: 'FFC8C8C8' } },
          right: { style: 'thin', color: { argb: 'FFC8C8C8' } },
        };
      });
    };

    applyBorder(headerRow1);
    applyBorder(headerRow2);

    /* ================= DATA ================= */

    this.FIManagerData.forEach((info: any, index: number) => {

      const row = worksheet.addRow([
        info.Rank,
        info.fimanager,
        info.StoreName,
        info.New,
        info.Used,
        info.Total,
        info.Pace,
        info.BackGross,
        info.BackGrossPace,
        info.BackgrossPVR
      ]);

      row.eachCell((cell: any, colNumber: number) => {

        // Grid borders
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          right: { style: 'thin', color: { argb: 'FFE0E0E0' } },
        };

        // Alignment like PDF
        if (colNumber === 2 || colNumber === 3) {
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        } else if ([8, 9, 10].includes(colNumber)) {
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        } else {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }

        // Currency
        if ([8, 9, 10].includes(colNumber) && typeof cell.value === 'number') {
          cell.numFmt = '"$"#,##0';
        }
      });

      // Alternate rows
      if (index % 2 === 0) {
        row.eachCell((cell: any) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF7F9FC' }
          };
        });
      }
    });

    /* ================= COLUMN WIDTH ================= */

    worksheet.columns.forEach(col => col.width = 18);
    worksheet.getColumn(2).width = 30;
    worksheet.getColumn(3).width = 30;

    /* ================= FREEZE HEADER ================= */
    worksheet.views = [{ state: 'frozen', ySplit: headerRow2.number }];

    /* ================= DOWNLOAD ================= */
    workbook.xlsx.writeBuffer().then(() => {
      this.shared.exportToExcel(workbook, 'F&I Manager Rankings');
    });
  }

}

