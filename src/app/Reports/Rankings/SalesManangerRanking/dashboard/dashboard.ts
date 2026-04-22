import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Subscription } from 'rxjs';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores, DateRangePicker],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {

  FIManagerData: any = [];
  NoData: boolean = false;

  storeorgrp: any = 'G';
  storeorgroup: any = ['G'];
  retailorlease: any = ['Retail', 'Lease', 'Misc', 'Special Order'];
  financetype: any = ['Finance', 'Cash', 'Lease'];
  neworused: string[] = ['New', 'Used'];
  GrossType: any = ['Back Gross'];
  DupGrossType: any = [];


  storeIds: any = '0';
  stores: any = [];
  columnName: any = 'Rank';
  columnState: any = 'asc';
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 8;

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
  DupFromDate: any = '';
  DupToDate: any = '';
  DupDateType: any = 'MTD';


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

  constructor(
    public shared: Sharedservice, public setdates: Setdates, private toast: ToastService,
  ) {

    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      if (localStorage.getItem('stime') != null) {
        let stime = localStorage.getItem('stime');
        if (stime != null && stime != '') {
          this.initializeDates(stime)
          this.DateType = stime
        }
      } else {
        this.initializeDates('MTD')
        this.DateType = 'MTD'
      }
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

    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }
    this.shared.setTitle(this.shared.common.titleName + '-Sales Manager Rankings');
    this.setHeaderData()
    this.GetData('Rank', 'asc');

  }

  ngOnInit(): void { }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    this.DateType = type
    localStorage.setItem('time', type);
    this.setDates(this.DateType)
  }

  setHeaderData() {
    const data = {
      title: 'Sales Manager Rankings',
      stores: this.storeIds,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groupId,
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
    this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    this.DupDateType = this.DateType
    this.DupGrossType = [...this.GrossType]

    this.FIManagerData = [];
    this.shared.spinner.show();
    const obj = {
      StartDate: this.FromDate,
      EndDate: this.ToDate,
      StoreID: this.storeIds,
      Exp: sortdata,
      OrderType: sortstate,
      RankBy: this.storeorgroup,
      UserID: 0,
      SaleType: this.neworused,
      DealType: this.retailorlease,
      FinanceType: this.financetype,
      IncludeFrontGross: this.GrossType.indexOf('Front Gross') >= 0 ? 'Y' : 'N'

    };
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetSalesManagerRankingsV1', obj).subscribe(
      (res) => {
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
          // this.toast.error(res.status, '');
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



  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  tabClick(col_Name: any) {

    // First click on a column
    if (this.columnName !== col_Name) {
      this.columnName = col_Name;
      this.columnState = 'asc';   // 👈 start with ASC on first click
    }
    // Clicking same column again → toggle
    else {
      this.columnState = this.columnState === 'asc' ? 'desc' : 'asc';
    }

    this.GetData(this.columnName, this.columnState);
  }

  excel!: Subscription;
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Sales Manager Rankings') {
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
      if (!obj || obj.title !== 'Sales Manager Rankings') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Sales Manager Rankings') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Sales Manager Rankings') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Sales Manager Rankings') return;
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

  private createPDF(): jsPDF {

    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a3'
    });

    /* ================= TITLE ================= */
    doc.setFontSize(14);
    doc.text('Sales Manager Rankings', 14, 12);

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
                    `${this.DupFromDate} to ${this.DupToDate}`
        );

    /* ================= COLUMN COUNT ================= */
    let totalCols = 8;
    if (this.DupGrossType.indexOf('Front Gross') >= 0) totalCols += 3;
    if (this.DupGrossType.indexOf('Back Gross') >= 0) totalCols += 4;
    if (this.DupGrossType.length > 1) totalCols += 3;

    /* ================= HEAD ================= */
    const head: any[] = [];

    /* ---- Header Row 1 (Group Titles) ---- */
    const groupRow: any[] = [
      { content: dateText, colSpan: 3, styles: { halign: 'center' } },
      { content: 'Unit Count', colSpan: 5, styles: { halign: 'center' } }
    ];

    if (this.DupGrossType.indexOf('Front Gross') >= 0) {
      groupRow.push({ content: 'Front Gross', colSpan: 3, styles: { halign: 'center' } });
    }

    if (this.DupGrossType.indexOf('Back Gross') >= 0) {
      groupRow.push({ content: 'Back Gross', colSpan: 4, styles: { halign: 'center' } });
    }

    if (this.DupGrossType.length > 1) {
      groupRow.push({ content: 'Total Gross', colSpan: 3, styles: { halign: 'center' } });
    }

    head.push(groupRow);

    /* ---- Header Row 2 (Column Titles) ---- */
    const colHeaders: any[] = [
      'Rank',
      'Sales Manager',
      'Store Name',
      'New',
      'Used',
      'Total',
      'Pace',
      '90 Day Avg'
    ];

    if (this.DupGrossType.indexOf('Front Gross') >= 0) {
      colHeaders.push('Gross', 'Pace', 'PVR');
    }

    if (this.DupGrossType.indexOf('Back Gross') >= 0) {
      colHeaders.push('Gross', 'Pace', 'PVR', '90 Day Avg');
    }

    if (this.DupGrossType.length > 1) {
      colHeaders.push('Gross', 'Pace', 'PVR');
    }

    head.push(colHeaders);

    /* ================= BODY ================= */
    const body: any[] = [];

    this.FIManagerData.forEach((spdata: any, i: number) => {

      const row: any[] = [
        spdata.data1 === 'Reports Total' ? '' : (spdata.Rank ?? '-'),
        spdata.data1 === 'Reports Total' ? '' : (spdata.SalesManager ?? '-'),
        spdata.StoreName ?? '-',
        spdata.New || '-',
        spdata.Used || '-',
        spdata.Total || '-',
        spdata.Pace || '-',
        spdata.UnitsDayAvg || '-'
      ];

      if (this.DupGrossType.indexOf('Front Gross') >= 0) {
        row.push(spdata.FrontGross, spdata.FrontGrossPace, spdata.FrontGrossPvr);
      }

      if (this.DupGrossType.indexOf('Back Gross') >= 0) {
        row.push(spdata.BackGross, spdata.BackGrossPace, spdata.BackGrossPvr, spdata.GrossDayAvg);
      }

      if (this.DupGrossType.length > 1) {
        row.push(spdata.TotalGross, spdata.TotalGrossPace, spdata.TotalGrossPvr);
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
        fillColor: [5, 84, 239], // Excel blue
        textColor: [255, 255, 255],
        fontStyle: 'bold'
      },

      didParseCell: (data: any) => {

        /* ===== GROUP HEADER (ROW 0) ===== */
        if (data.section === 'head' && data.row.index === 0) {
          data.cell.styles.fillColor = [255, 255, 255];
          data.cell.styles.textColor = [0, 0, 0];
          data.cell.styles.fontStyle = 'bold';
        }

        /* ===== BODY ===== */
        if (data.section === 'body') {

          /* Alignment */
          if ([0, 1, 2].includes(data.column.index)) {
            data.cell.styles.halign = 'left';
          } else {
            data.cell.styles.halign = 'right';
          }

          /* Dash for empty */
          if (data.cell.raw === null || data.cell.raw === 0) {
            data.cell.text = ['-'];
            data.cell.styles.halign = 'center';
          }

          /* Currency formatting */
          if (typeof data.cell.raw === 'number' && data.column.index >= 8) {
            const val = Math.round(data.cell.raw);
            data.cell.text = [`$ ${val.toLocaleString()}`];

            if (val < 0) {
              data.cell.styles.textColor = [255, 0, 0];
            }
          }

          /* Alternate rows */
          if (data.row.index % 2 === 0) {
            data.cell.styles.fillColor = [245, 247, 250];
          }

          /* Reports Total row */
          if (this.FIManagerData[data.row.index]?.data1 === 'Reports Total') {
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
      doc.save('Sales Manager Rankings.pdf');
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
      const pdfFile = this.blobToFile(pdfBlob, 'Sales Manager Rankings.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Sales Manager Rankings');
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


  //----------REPORTS---------------//
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
      'type': 'M', 'others': 'N'
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

  // Deal type & status
  multipleorsingle(block: any, e: any) {
    if (block == 'NU') {
      const index = this.neworused.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.neworused.splice(index, 1);
      } else {
        this.neworused.push(e);
      }
    }

    if (block == 'RL') {
      const index = this.retailorlease.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.retailorlease.splice(index, 1);
      } else {
        this.retailorlease.push(e);
      }
    }
    if (block == 'DS') {
      const index = this.financetype.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.financetype.splice(index, 1);
      } else {
        this.financetype.push(e);
      }
    }
    if (block == 'GT') {
      const index = this.GrossType.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.GrossType.splice(index, 1);
      } else {
        this.GrossType.push(e);
      }
    }

  }

  storeorgroups(block: any, e: any) {
    this.storeorgroup = [];
    this.storeorgroup.push(e)
  }
  viewreport() {
    this.activePopover = -1;

    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      return;
    } else
      if (this.retailorlease.length == 0) {
        this.toast.show('Please select any one Deal Type', 'warning', 'Warning');
      }
      else if (this.neworused.length == 0) {
        this.toast.show('Please Select Atleast One Sale Type', 'warning', 'Warning');
      }
      else if (this.financetype.length == 0) {
        this.toast.show('Please Select Atleast One Finance Type', 'warning', 'Warning');
      }
      else if (this.GrossType && this.GrossType.length == 0) {
        this.toast.show('Please Select Atleast One Gross Type', 'warning', 'Warning');

      }
      else {
        this.setHeaderData()
        this.GetData(this.columnName, this.columnState);


      }
  }
  close() {
    this.shared.ngbmodal.dismissAll();
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
    const worksheet = workbook.addWorksheet('Sales Manager Rankings');

    /* ================= TITLE ================= */
    const title = worksheet.addRow(['Sales Manager Rankings']);
    title.font = { size: 14, bold: true };
    worksheet.mergeCells('A1:R1');

    worksheet.addRow([]);

    /* ================= FILTERS ================= */
    const formattedFromDate = this.shared.datePipe.transform(this.FromDate, 'dd-MMM-yyyy');
    const formattedToDate = this.shared.datePipe.transform(this.ToDate, 'dd-MMM-yyyy');

    const selectedStoreIds = this.storeIds?.map((x: any) => x.toString()) || [];

    const storeValue = (this.stores || [])
      .filter((s: any) => selectedStoreIds.includes(s.ID.toString()))
      .map((s: any) => s.storename)
      .join(', ');

    const filters = [
      { name: 'Store:', values: storeValue },
      { name: 'Time Frame:', values: `${formattedFromDate} to ${formattedToDate}` },
      { name: 'Rank By:', values: this.storeorgroup == 'S' ? 'Store' : 'Group' },
      { name: 'Deal Type:', values: this.retailorlease || 'All' },
      { name: 'Gross Type:', values: this.GrossType?.toString() || 'All' }
    ];

    filters.forEach(f => {
      const row = worksheet.addRow([f.name, f.values]);
      row.getCell(1).font = { bold: true };
      worksheet.mergeCells(`B${row.number}:F${row.number}`);
    });

    worksheet.addRow([]);

    /* ================= DATE HEADER ================= */
    let dateHeader: any = '';
    if (this.datetype() === 'C') {
      const from = new Date(this.FromDate);
      const to = new Date(this.ToDate);

      const fromText =
        ('0' + (from.getMonth() + 1)).slice(-2) + '.' +
        ('0' + from.getDate()).slice(-2) + '.' +
        from.getFullYear();

      const toText =
        ('0' + (to.getMonth() + 1)).slice(-2) + '.' +
        ('0' + to.getDate()).slice(-2) + '.' +
        to.getFullYear();

      dateHeader = `${fromText} - ${toText}`;
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
        'Back Gross', '', '', '',
        'Total Gross', '', ''
      ];

      secondHeader = [
        'Rank', 'Sales Manager', 'Store Name',
        'New', 'Used', 'Total', 'Pace', '90 Day Avg',
        'Gross', 'Pace', 'PVR',
        'Gross', 'Pace', 'PVR', '90 Day Avg',
        'Gross', 'Pace', 'PVR'
      ];

      bindingHeaders = [
        'Rank', 'SalesManager', 'StoreName',
        'New', 'Used', 'Total', 'Pace', 'UnitsDayAvg',
        'FrontGross', 'FrontGrossPace', 'FrontGrossPvr',
        'BackGross', 'BackGrossPace', 'BackGrossPvr', 'GrossDayAvg',
        'TotalGross', 'TotalGrossPace', 'TotalGrossPvr'
      ];

      worksheet.mergeCells(`A${startRow}:C${startRow}`);
      worksheet.mergeCells(`D${startRow}:H${startRow}`);
      worksheet.mergeCells(`I${startRow}:K${startRow}`);
      worksheet.mergeCells(`L${startRow}:O${startRow}`);
      worksheet.mergeCells(`P${startRow}:R${startRow}`);

    } else if (this.DupGrossType.includes('Front Gross')) {

      firstHeader = [
        dateHeader, '', '',
        'Unit Count', '', '', '', '',
        'Front Gross', '', ''
      ];

      secondHeader = [
        'Rank', 'Sales Manager', 'Store Name',
        'New', 'Used', 'Total', 'Pace', '90 Day Avg',
        'Gross', 'Pace', 'PVR'
      ];

      bindingHeaders = [
        'Rank', 'SalesManager', 'StoreName',
        'New', 'Used', 'Total', 'Pace', 'UnitsDayAvg',
        'FrontGross', 'FrontGrossPace', 'FrontGrossPvr'
      ];

      worksheet.mergeCells(`A${startRow}:C${startRow}`);
      worksheet.mergeCells(`D${startRow}:H${startRow}`);
      worksheet.mergeCells(`I${startRow}:K${startRow}`);

    } else {

      firstHeader = [
        dateHeader, '', '',
        'Unit Count', '', '', '', '',
        'Back Gross', '', '', ''
      ];

      secondHeader = [
        'Rank', 'Sales Manager', 'Store Name',
        'New', 'Used', 'Total', 'Pace', '90 Day Avg',
        'Gross', 'Pace', 'PVR', '90 Day Avg'
      ];

      bindingHeaders = [
        'Rank', 'SalesManager', 'StoreName',
        'New', 'Used', 'Total', 'Pace', 'UnitsDayAvg',
        'BackGross', 'BackGrossPace', 'BackGrossPvr', 'GrossDayAvg'
      ];

      worksheet.mergeCells(`A${startRow}:C${startRow}`);
      worksheet.mergeCells(`D${startRow}:H${startRow}`);
      worksheet.mergeCells(`I${startRow}:L${startRow}`);
    }

    const headerRow1 = worksheet.addRow(firstHeader);
    const headerRow2 = worksheet.addRow(secondHeader);

    /* ================= HEADER STYLE ================= */

    // 🔵 GROUP HEADER (ROW 1 → #0554EF)
    headerRow1.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0554EF' } // exact color
      };
      cell.font = {
        bold: true,
        color: { argb: 'FFFFFFFF' }, // white text

      };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };

      cell.border = {
        top: { style: 'thin', color: { argb: 'FF0554EF' } },
        left: { style: 'thin', color: { argb: 'FF0554EF' } },
        bottom: { style: 'thin', color: { argb: 'FF0554EF' } },
        right: { style: 'thin', color: { argb: 'FF0554EF' } }
      };
    });

    // 🔷 COLUMN HEADER (ROW 2 → #D9E7FF)
    headerRow2.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9E7FF' } // exact color
      };
      cell.font = {
        bold: true,
        color: { argb: 'FF000000' }, // black text

      };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };

      cell.border = {
        top: { style: 'thin', color: { argb: 'FFC9D6F0' } },
        left: { style: 'thin', color: { argb: 'FFC9D6F0' } },
        bottom: { style: 'thin', color: { argb: 'FFC9D6F0' } },
        right: { style: 'thin', color: { argb: 'FFC9D6F0' } }
      };
    });
    /* ================= DATA ================= */

    const currencyFields = [
      'FrontGross', 'FrontGrossPace', 'FrontGrossPvr',
      'BackGross', 'BackGrossPace', 'BackGrossPvr',
      'GrossDayAvg',
      'TotalGross', 'TotalGrossPace', 'TotalGrossPvr'
    ];

    this.FIManagerData.forEach((item: any, i: number) => {

      const isTotalRow = item.data1 === 'Reports Total';

      const row = worksheet.addRow(
        bindingHeaders.map(k => {
          let val = item[k];
          if (val === null || val === 0) return '-';
          return val;
        })
      );

      row.eachCell((cell, col) => {

        const field = bindingHeaders[col - 1];

        /* ===== ALIGNMENT ===== */
        if (col === 1) {
          // Rank column
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        } else if (col <= 3) {
          // Sales Manager + Store
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        } else {
          // Numbers
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        }

        // CURRENCY
        if (currencyFields.includes(field) && typeof item[field] === 'number') {
          const val = Math.round(item[field]);
          cell.value = `$ ${val.toLocaleString()}`;

          if (val < 0) {
            cell.font = { color: { argb: 'FFFF0000' } };
          }
        }

        // GRID (LIGHT LIKE PDF)
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE3E7ED' } },
          left: { style: 'thin', color: { argb: 'FFE3E7ED' } },
          bottom: { style: 'thin', color: { argb: 'FFE3E7ED' } },
          right: { style: 'thin', color: { argb: 'FFE3E7ED' } }
        };

        // ALTERNATE ROW
        if (i % 2 === 0 && !isTotalRow) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF5F7FA' }
          };
        }

        // TOTAL ROW
        if (isTotalRow) {
          cell.font = { bold: true };
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFDCE6FF' }
          };
        }

      });

    });

    /* ================= COLUMN WIDTH ================= */
    worksheet.columns.forEach((col, i) => {
      if (i === 1 || i === 2) col.width = 28;
      else col.width = 16;
    });

    /* ================= EXPORT ================= */
    workbook.xlsx.writeBuffer().then(() => {
      this.shared.exportToExcel(workbook, 'Sales Manager Rankings');
    });

  }

  scrollPosition = 0;

  getScrollPosition(event: any): void {
    this.scrollPosition = event.target.scrollLeft;
    console.log(this.scrollPosition, event.target.scrollTop);
  }


}
