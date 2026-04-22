import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
// import { Stores } from '../../../../CommonFilters/stores/stores';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Subscription } from 'rxjs';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { FilterPipe } from '../../../../Core/Providers/filterpipe/filter.pipe';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import * as FileSaver from 'file-saver';
import { Workbook } from 'exceljs';
import { DatePipe } from '@angular/common';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  Appointmentdata: any = [];
  NoData: any = false;
  appointment: any = [];
  selectedStore = '0';
  selectedstrid = 0;
  // stores: any = [];
  PageCount = 1;
  LastCount!: boolean;
  Pagination: boolean = false;
  storeIds: any = '0';

  // groups: any = 1
  filteredCustomers: any;
  Appointmentsearch: any;
  searcdata: any;
  FromDate: any = '';
  ToDate: any = '';

  minDate!: Date;
  maxDate!: Date;
  todaytitle: any;
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;

  // FromDate: any;
  // ToDate: any;
  DateType: string = 'C';
  displaytime: any = new Date();
  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    Types: [
      // { 'code': '', 'name': '' },
      // { 'code': 'QTD', 'name': 'QTD' },
      // { 'code': 'YTD', 'name': 'YTD' },
      // { 'code': 'PYTD', 'name': 'PYTD' },
      // { 'code': 'LY', 'name': 'Last Year' },
      // { 'code': 'LM', 'name': 'Last Month' },
      // { 'code': 'PM', 'name': 'Same Month PY' },
    ]
  }
  StoreValues: any = 0;
  popup: any = [{ type: 'Popup' }];
  groups: any = 1;
  gridvisibility: any;
  bsRangeValue!: Date[];
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  index = '';
  groupId: any = 0;
  stores: any = []
  // activePopover: number = -1;
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  // @HostListener('document:click', ['$event'])
  // onDocumentClick(event: MouseEvent) {
  //   const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .dropdown-menu , .timeframe, .reportstores-card');
  //   if (!clickedInside) {
  //     this.activePopover = -1;
  //   }
  // }
  constructor(public shared: Sharedservice, public setdates: Setdates, private toast: ToastService,

  ) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.StoreValues = '1,2';
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.StoreValues = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
      if (this.StoreValues.toString().indexOf(',') > 0) {
        this.gridvisibility = 'DL';
      } else {
        this.gridvisibility = 'SL';
      }
    }

    localStorage.setItem('time', 'C');

    this.titleSetting()
    this.shared.setTitle(this.shared.common.titleName + '- Lead Source Report');

    this.todaytitle = this.shared.datePipe.transform(new Date(), 'MM.dd.yyyy');
    let today = new Date();
    let enddate = new Date(today.setDate(today.getDate() - 1));
    this.FromDate =
      ('0' + (enddate.getMonth() + 1)).slice(-2) +
      '/01' +
      '/' +
      enddate.getFullYear();
    this.ToDate =
      ('0' + (enddate.getMonth() + 1)).slice(-2) +
      '/' +
      ('0' + enddate.getDate()).slice(-2) +
      '/' +
      enddate.getFullYear();

    this.FromDate = this.FromDate.replace(/-/g, '/');
    this.ToDate = this.ToDate.replace(/-/g, '/');
    this.setDates(this.DateType)
    // if (localStorage.getItem('Fav') != 'Y') {
    const data = {
      title: this.dynamicTitle,
      stores: this.StoreValues,
      groups: this.groups,
      count: 0,
      fromdate: this.FromDate,
      todate: this.ToDate,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });

    this.GetLeadSourceReport();
    // }
  }
  dynamicTitle = 'Lead Source Report'
  titleSetting() {
    // if (localStorage.getItem('time') == 'C') {
    console.log(this.todaytitle, this.shared.datePipe.transform(this.FromDate, 'MM-dd-yyyy'), this.shared.datePipe.transform(this.ToDate, 'MM-dd-yyyy'));

    // if ((this.todaytitle == this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy')) && (this.todaytitle == this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy'))) {
    //   this.dynamicTitle = 'Today Lead Source Report'
    // }
    // else {
    this.dynamicTitle = 'Lead Source Report'
    // }
    const data = {
      title: this.dynamicTitle,
      stores: this.StoreValues,
      groups: this.groups,
      count: 0,
      fromdate: this.FromDate,
      todate: this.ToDate,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });

    // }
  }

  setHeaderData() {
    const data = {
      title: this.dynamicTitle,
      stores: this.storeIds,
      groups: 1,
      count: 0,
      fromdate: this.FromDate,
      todate: this.ToDate,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }
  ngOnInit(): void {
    this.getStores();
  }


  searchQuery: string = '';
  isDesc: boolean = false;
  column: string = '';

  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // // console.log(property)
    this.Appointmentdata.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
  getStoresandGroupsValues() {
    this.storesFilterData.groupsArray = this.groupsArray;
    this.storesFilterData.groupId = this.groupId;
    this.storesFilterData.storesArray = this.stores;
    this.storesFilterData.storeids = this.StoreValues;
    this.storesFilterData.groupName = this.groupName;
    this.storesFilterData.storename = this.storename;
    this.storesFilterData.storecount = this.storecount;
    this.storesFilterData.storedisplayname = this.storedisplayname;

    this.storesFilterData = {
      groupsArray: this.groupsArray,
      groupId: this.groupId,
      storesArray: this.stores,
      storeids: this.StoreValues,
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
    this.StoreValues = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
  }
  getStores() {
    this.selectedstorevalues = [];
  }

  LeadSourceData: any = [];

  GetLeadSourceReport() {
    this.LeadSourceData = [];
    this.NoData = false;
    this.shared.spinner.show();
    let obj = {};
    obj = {
      DEALERID: this.StoreValues.toString(),
      StartDate: this.shared.datePipe.transform(this.FromDate, 'dd-MMM-yyyy'),
      EndDate: this.shared.datePipe.transform(this.ToDate, 'dd-MMM-yyyy'),
      UserID: JSON.parse(localStorage.getItem('userInfo')!).user_Info.userid || 0,
    };
    console.log(obj);
    // const curl = this.shared.getEnviUrl() +this.shared.common.routeEndpoint+'GetDrivecentricAppointments';
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetLeadSourceStatusTypes', obj).subscribe(
      (res) => {
        const currentTitle = document.title;
        if (res.status == 200) {
          this.LeadSourceData = res.response;
          this.LeadSourceData.some(function (x: any) {
            if (x.data1 != undefined) {
              x.data1 = JSON.parse(x.data1);
            }
          });

          if (this.LeadSourceData.length > 1) {
            this.LeadSourceData.sort((a: any, b: any) => {
              if (a.leadSources === 'Total' && b.leadSources !== 'Total') {
                return -1;
              } else if (
                a.leadSources !== 'Total' &&
                b.leadSources === 'Total'
              ) {
                return 1;
              } else {
                return 0;
              }
            });
            let position = this.scrollCurrentposition + 5;
            setTimeout(() => {
              this.scrollcent.nativeElement.scrollTop = position;
              // console.log(position);
            }, 500);
          }
          console.log('LeadSourceData', this.LeadSourceData);
          this.shared.spinner.hide();
          if (this.LeadSourceData.length > 0) {
            this.NoData = false;
          } else {
            this.NoData = true;
          }
        } else {
          alert('Invalid Details');
          this.shared.spinner.hide();
        }
      },
      (error) => {
        console.log(error);
      }
    );
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Lead Source Report') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.StoreValues.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.StoreValues.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.StoreValues)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: { obj: { title: string; stateEmailPdf: boolean; }; }) => {
      if (this.email != undefined) {
        if (res.obj.title == this.dynamicTitle) {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: { obj: { title: string; statePrint: boolean; }; }) => {
      if (this.print != undefined) {
        if (res.obj.title == this.dynamicTitle) {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: { obj: { title: string; statePDF: boolean; }; }) => {
      if (this.Pdf != undefined) {

        if (res.obj.title == this.dynamicTitle) {
          if (res.obj.statePDF == true) {
            this.downloadPDF()
          }
        }
      }
    });

    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.excel != undefined) {
        if (res.obj.title == this.dynamicTitle) {
          if (res.obj.state == true) {
            this.exportLeadSourceExcel();
          }
        }
      }
    });
  }
  ngOnDestroy() {
    if (this.excel != undefined) {
      this.excel.unsubscribe()
    }
    if (this.Pdf != undefined) {
      this.Pdf.unsubscribe()
    }
    if (this.print != undefined) {
      this.print.unsubscribe()
    }
    if (this.email != undefined) {
      this.email.unsubscribe()
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

  //------Reports-----------//
  activePopover: number = -1;
  Teams = 'S';
  selectedstorevalues: any = [];
  labortypes: any = [];
  selectedlabortypevalues: any = [];
  Performance: any = 'Load';
  changesvalues: any = [];
  AllLabortypes: boolean = true;
  toporbottom: any = ['T'];

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.setDates(type)
  }



  updatedDates(data: any) {
    console.log(data);
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
    this.bsRangeValue = [new Date(this.FromDate), new Date(this.ToDate)];

  }
  open() {
    this.Dates = { ...this.Dates }
    this.setDates(this.DateType)
  }
  setDates(type: any) {
    this.DateType == 'C' ? this.displaytime = '( ' + this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy') + ' - ' + this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy') + ' )' :
      this.displaytime = 'Time Frame (' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ')';
    // this.maxDate = new Date();
    // this.minDate = new Date();
    // this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    // this.maxDate.setDate(this.maxDate.getDate());
    this.Dates.FromDate = this.FromDate;
    this.Dates.ToDate = this.ToDate;
    this.Dates.MinDate = this.minDate;
    this.Dates.MaxDate = this.maxDate;
    this.Dates.DateType = this.DateType;
    this.Dates.DisplayTime = this.displaytime;
    console.log(this.FromDate, this.ToDate, this.DateType, this.displaytime, '..............');
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

  viewreport() {
    if (!this.StoreValues || this.StoreValues.length === 0) {

      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      return;
    }

    this.activePopover = -1
    this.titleSetting()
    this.setHeaderData()
    this.GetLeadSourceReport()

  }
  getExcelStoreLabel(): string {
    // 0 selected
    if (!this.StoreValues || this.StoreValues.length === 0) {
      return 'Selected (0)';
    }

    // All selected → show all store names
    if (this.StoreValues.length === this.stores.length) {
      return this.stores
        .map((s: any) => s.storename)
        .join(', ');
    }

    // Single selected → store name
    if (this.StoreValues.length === 1) {
      const store = this.stores.find(
        (s: any) => s.ID === this.StoreValues[0]
      );
      return store ? store.storename : '-';
    }

    // Multiple (not all)
    return `Selected (${this.StoreValues.length})`;
  }


  formatDollar(value: any): string {
    if (value === 0 || value === null || value === undefined) {
      return '-';
    }
    return '$' + Number(value).toLocaleString('en-US', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  }
  formatDateMMDDYYYY(value: any): string {
    if (!value) return '-';
    const date = new Date(value);
    if (isNaN(date.getTime())) return '-';
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const year = date.getFullYear();
    return `${month}.${day}.${year}`;
  }
  private createLeadSourcePDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });

    /* ================= TITLE ================= */
    doc.setFontSize(14);
    doc.text('Lead Source Report', 14, 12);

    /* ================= DATE RANGE ================= */
    const from = this.formatDateMMDDYYYY(this.FromDate);
    const to = this.formatDateMMDDYYYY(this.ToDate);

    doc.setFontSize(11);
    doc.text(`${from} - ${to}`, 14, 18);

    /* ================= HEADERS ================= */

    const head = [[
      'Store',
      'Active',
      'Sold',
      'Lost',
      'Bad',
      'Complete',
      'Total Count'
    ]];

    /* ================= BODY ================= */

    const rows: any[] = [];

    (this.LeadSourceData || []).forEach((LSST: any) => {

      const isTotal = LSST.leadSources === 'Total';

      rows.push([
        {
          content: isTotal ? 'Report Totals' : (LSST.leadSources || '-'),
          styles: { fillColor: isTotal ? [200, 220, 255] : [217, 231, 255] }
        },
        { content: LSST.Active || '-', styles: { fillColor: [217, 231, 255] } },
        { content: LSST.Sold || '-', styles: { fillColor: [217, 231, 255] } },
        { content: LSST.Lost || '-', styles: { fillColor: [217, 231, 255] } },
        { content: LSST.Bad || '-', styles: { fillColor: [217, 231, 255] } },
        { content: LSST.Complete || '-', styles: { fillColor: [217, 231, 255] } },
        { content: LSST.TotalCount || '-', styles: { fillColor: [217, 231, 255] } }
      ]);

      (LSST.data1 || []).forEach((sub: any) => {
        rows.push([
          { content: '   ' + (sub.StoreName || '-'), styles: { fillColor: [255, 255, 255] } },
          sub.Active || '-',
          sub.Sold || '-',
          sub.Lost || '-',
          sub.Bad || '-',
          sub.Complete || '-',
          sub.TotalCount || '-'
        ]);
      });

    });

    const pageW = doc.internal.pageSize.getWidth();
    const margin = 6;
    const colWidth = (pageW - margin * 2) / head[0].length;

    const columnStyles: any = {};
    head[0].forEach((_: any, i: number) => {
      columnStyles[i] = { cellWidth: colWidth };
    });

    autoTable(doc, {
      startY: 22,
      head,
      body: rows,
      theme: 'grid',
      margin: { left: 6, right: 6 },

      styles: {
        fontSize: 9,
        cellPadding: 2,
        lineWidth: 0.2,
        lineColor: [180, 180, 180]
      },

      headStyles: {
        fillColor: [5, 84, 239],
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        halign: 'center'
      },

      columnStyles,

      didParseCell: (data: any) => {

        if (data.section === 'body') {

          let raw = data.cell.raw;

          if (typeof raw === 'object' && raw?.content !== undefined) {
            raw = raw.content;
          }
          const firstCol = data.row.raw?.[0];
          let firstVal = firstCol;

          if (typeof firstCol === 'object' && firstCol?.content) {
            firstVal = firstCol.content;
          }

          if (firstVal === 'Report Totals') {
            data.cell.styles.fontStyle = 'bold';
          }
          if (data.column.index === 0) {
            data.cell.styles.halign = 'left';
          } else {
            data.cell.styles.halign = 'center';
          }
          if (data.column.index > 0) {

            if (raw === '-' || raw === null || raw === undefined || raw === '') {
              data.cell.text = ['-'];
              data.cell.styles.textColor = [0, 0, 0];
              return;
            }

            const val = Number(raw);

            if (!isNaN(val)) {
              data.cell.text = [val.toLocaleString('en-US')];
              data.cell.styles.textColor = val < 0
                ? [220, 53, 69]
                : [0, 0, 0];
            }
          }
        }
      }
    });

    return doc;
  }
  downloadPDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createLeadSourcePDF();
      doc.save(`Lead Source Report.pdf`);
    } catch (e) {
      console.error(e);
      this.toast.show('Error generating PDF', 'danger', 'Error');
    } finally {
      this.shared.spinner.hide();
    }
  }

  printPDF() {
    try {
      const doc = this.createLeadSourcePDF();
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
  private exportLeadSourceExcel(): void {

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Lead Source Report');


    worksheet.addRow(['Lead Source Report']).font = { bold: true, size: 14 };
    worksheet.addRow([]);
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 5,
        topLeftCell: 'A6',
        showGridLines: false,
      },
    ];
    const from = this.formatDateMMDDYYYY(this.FromDate);
    const to = this.formatDateMMDDYYYY(this.ToDate);

    worksheet.addRow([`${from} - ${to}`]);
    worksheet.addRow([]);


    const headers = [
      'Store', 'Active', 'Sold', 'Lost', 'Bad', 'Complete', 'Total Count'
    ];

    const headerRow = worksheet.addRow(headers);

    headerRow.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0554EF' } };
      cell.font = { bold: true, color: { argb: 'FFFFFF' } };
      cell.alignment = { horizontal: 'center' };
      cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
    });


    (this.LeadSourceData || []).forEach((LSST: any) => {

      const isTotal = LSST.leadSources === 'Total';

      const row = worksheet.addRow([
        isTotal ? 'Report Totals' : (LSST.leadSources || '-'),
        LSST.Active || '-',
        LSST.Sold || '-',
        LSST.Lost || '-',
        LSST.Bad || '-',
        LSST.Complete || '-',
        LSST.TotalCount || '-'
      ]);

      row.eachCell((cell, i) => {

        cell.alignment = {
          horizontal: i === 1 ? 'left' : 'center'
        };

        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: isTotal ? 'C8DCFF' : 'D9E7FF' }
        };

        // ✅ MAKE REPORT TOTALS BOLD
        if (isTotal) {
          cell.font = {
            bold: true,
            color: { argb: '000000' }
          };
        }
        if (i > 1) {
          if (cell.value === null) {
            cell.value = '-';
          } else {
            cell.numFmt = '#,##0;[Red]-#,##0';
          }
        }
        // NEGATIVE VALUES → RED
        if (!isTotal && i > 1 && typeof cell.value === 'number' && cell.value < 0) {
          cell.font = {
            color: { argb: 'DC3545' }
          };
        }

        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };
      });

      // CHILD ROWS
      (LSST.data1 || []).forEach((sub: any) => {

        const subRow = worksheet.addRow([
          '   ' + (sub.StoreName || '-'),
          sub.Active || '-',
          sub.Sold || '-',
          sub.Lost || '-',
          sub.Bad || '-',
          sub.Complete || '-',
          sub.TotalCount || '-'
        ]);

        subRow.eachCell((cell, i) => {

          cell.alignment = {
            horizontal: i === 1 ? 'left' : 'center'
          };

          // NEGATIVE
          if (i > 1 && typeof cell.value === 'number' && cell.value < 0) {
            cell.font = { color: { argb: 'DC3545' } };
          }
          if (i > 1) {
            if (cell.value === null) {
              cell.value = '-';
            } else {
              cell.numFmt = '#,##0;[Red]-#,##0';
            }
          }
          cell.border = {
            left: { style: 'thin' },
            right: { style: 'thin' },
            bottom: { style: 'thin' }
          };
        });

      });

    });

    /* ================= WIDTH ================= */

    const widths = [30, 15, 15, 15, 15, 18, 18];
    widths.forEach((w, i) => worksheet.getColumn(i + 1).width = w);

    workbook.xlsx.writeBuffer().then(buffer => {
      const blob = new Blob([buffer]);
      FileSaver.saveAs(blob, 'Lead_Source_Report.xlsx');
    });
  }


}




