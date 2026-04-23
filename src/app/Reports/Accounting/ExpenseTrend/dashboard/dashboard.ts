import { Component, ElementRef, HostListener, Injector, ViewChild } from '@angular/core';
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
import { BsDatepickerConfig, BsDatepickerDirective, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Router } from '@angular/router';
import { ExpensetrendGraph } from '../expensetrend-graph/expensetrend-graph';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { Stores } from '../../../../CommonFilters/stores/stores';


@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  Current_Date: any;

  // gridshow:boolea=false
  // monthgridshow:boolean=false
  FSData: any = [];
  date: any;
  ExpenseTrendByStoreKeys: any = [];
  AllDatakeys: any = [];
  ExpenseTrendByStore_Excel: any;
  XpenseTrendByStoreKeys: any = [];
  ExpenseTrendByStore: any;
  ExpenseTrendByStore_ExcelMonth: any;
  XpenseTrendByStoreKeysMonth: any = [];
  ExpenseTrendByStoreKeysMonth: any = [];
  AllDatakeysMonth: any = [];
  ExpenseTrendByStoreMonth: any;
  SubFilter: any;
  SelectedTab: any = [];
  Month!: any;
  // stores: any;
  selectedstorevalues: any = [];
  selectedstorename: any;
  selectedFilters: string[] = [];
  selectedLabel: string = '( All )';
  StoreName: any = 'All Stores';
  Filter: any = ['New', 'Used', 'Service', 'Parts', 'Detail'];
  StoreValues: any = 2;
  groups: any = 1;
  PresentDayDate: string;

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }


  stores: any = [];
  selectDate: Date = new Date();
  currentMonth: any = '';
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = '0';
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
    private comm: common,
    private router: Router,
    private toast: ToastService,
    private injector: Injector,
    public shared: Sharedservice,
  ) {
    const lastMonth = new Date();
    let today = new Date();
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
    if (today.getDate() < 5) {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth() - 1));
    } else {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth()));
    }
    this.title.setTitle(this.comm.titleName + '-Expense Trend');
    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Expense Trend',
        path1: '',
        path2: '',
        path3: '',
        Month: this.date,
        stores: this.storeIds.toString(),
        store: this.storeIds,
        filter: this.Filter,
        groups: 1,
        count: 0,
      };
      this.apiSrvc.SetHeaderData({
        obj: data,
      });
    }
    const format = 'ddMMyyyy';
    const locale = 'en-US';
    const myDate = new Date();
    const formattedDate = formatDate(myDate, format, locale);
    this.PresentDayDate = formattedDate;
    this.selectDate = this.date
    this.currentMonth = this.selectDate;
    this.selectedFilters = [...this.filters];
    this.GetDataByMonths(this.currentMonth, this.selectedFilters);
  }
  roleId: any;
  ngOnInit(): void {
    this.roleId = localStorage.getItem('roleId');
    console.log('role Id', this.roleId);
    this.getStores();

  }

  StoreNamesHeadings: any = [];
  MonthsHeadings: any = [];

  Scrollpercent: any = 0;
  scrollPositionStoring: number = 0;
  scrollCurrentPosition: number = 0;

  @ViewChild('scrollcent') scrollcent!: ElementRef;

  updateVerticalScroll(event: any): void {
    this.scrollCurrentPosition = event.target.scrollTop;
  }

  // Filters list
  filters: string[] = ['New', 'Used', 'Service', 'Parts', 'Detail'];
  activePopover: number | null = null;
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month',
    maxDate: new Date()
  };
  monthPicker!: BsDatepickerDirective;
  openMonthPicker() {
    if (this.monthPicker) {
      this.monthPicker.show();
    }
  }

  togglePopover(index: number) {
    this.activePopover = this.activePopover === index ? null : index;
  }


  isSelected(filter: string): boolean {
    return this.selectedFilters.includes(filter);
  }

  toggleFilter(filter: string) {
    const index = this.selectedFilters.indexOf(filter);
    if (index >= 0) {
      this.selectedFilters.splice(index, 1);
    } else {
      this.selectedFilters.push(filter);
    }
    this.updateLabel();
  }

  updateLabel() {
    if (!this.selectedFilters || this.selectedFilters.length === 0) {
      this.selectedLabel = '( Select )';
    } else if (this.selectedFilters.length === this.filters.length) {
      this.selectedLabel = '( All )';
    } else if (this.selectedFilters.length === 1) {
      this.selectedLabel = `( ${this.selectedFilters[0]} )`;
    } else {
      this.selectedLabel = `( Selected ${this.selectedFilters.length} )`;
    }
  }



  applyDateChange() {
    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      return;
    }
    if (!this.selectedFilters || this.selectedFilters.length === 0) {
      this.toast.show('Please Select Atleast One Department', 'warning', 'Warning');
      return;
    }
    else {
      this.currentMonth = this.formatMonth(this.selectDate);
      this.GetDataByMonths(this.currentMonth, this.selectedFilters);
      this.activePopover = null;
      this.isLoading = true;
    }
  }
  formatMonth(date: Date): string {
    const year = date.getFullYear();
    const month = ('0' + (date.getMonth() + 1)).slice(-2);
    return `${year}-${month}`;
  }
  isLoading = true;
  NoData = false;
  EndDate: any;
  dates: any;
  storeName: any;
  ExpenseTrendKeys: any = [];
  GetDataByMonths(date: any, filters: any) {
    this.spinner.show();
    this.dates = [];
    this.ExpenseTrendByStoreMonth = [];
    const currentDate = new Date(date);
    const currentMonth = currentDate.getMonth();
    const currentYear = currentDate.getFullYear();
    for (let i = 0; i < 12; i++) {
      const lastMonthDate = new Date(currentYear, currentMonth - i, 1);
      this.dates.push(this.formatDate(lastMonthDate));
    }
    const DateToday = this.datepipe.transform(
      new Date(date),
      'yyyy-MM-dd'
    );
    const obj = {
      AS_IDS: this.storeIds.toString(),
      DATE: this.shared.datePipe.transform(date, 'yyyy-MM') + '-10',
      // DEPARTMENT: filters.toString()
    };
    console.log(obj);
    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetExpenseTrendReport', obj)
      .subscribe(
        (x: any) => {
          this.isLoading = false;
          if (x.status == 200) {
            this.ExpenseTrendByStoreMonth = x.response;
            const serviceKeys = Object.keys(x.response[0]).slice(10);
            this.ExpenseTrendKeys = serviceKeys;
            console.log(
              'ExpenseTrendByStoreMonth',
              this.ExpenseTrendByStoreMonth
            );
            console.log('ExpenseTrendByMonthKeys', this.ExpenseTrendKeys);
            this.spinner.hide();
            this.NoData = this.ExpenseTrendByStoreMonth.length > 6;
          } else {
            this.ExpenseTrendByStoreMonth = [];
            this.ExpenseTrendKeys = [];
            this.NoData = false;
          }
          this.spinner.hide();
        },
        () => {
          this.isLoading = false;
          this.NoData = false;
          this.ExpenseTrendByStoreMonth = [];
          this.ExpenseTrendKeys = [];
          this.spinner.hide();
        }
      );
  }
  formatDate(date: Date): string {
    const options: Intl.DateTimeFormatOptions = {
      year: 'numeric',
      month: 'short',
    };
    return date.toLocaleDateString(undefined, options);
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  SFstate: any;
  StoreCodes: any;
  block: any = '';
  Report: any = '';

  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Expense Trend') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })

    this.apiSrvc.GetReports().subscribe((data) => {
      if (data.obj.Reference == 'Expense Trend') {
        console.log(data);
        if (data.obj.header == undefined) {
          this.date = data.obj.month;
          this.Month = data.obj.month;
          this.StoreValues = data.obj.storeValues;
          this.StoreName = data.obj.Sname;
          this.Filter = data.obj.filter;
          this.SubFilter = data.obj.subfilters;
          this.StoreCodes = data.obj.storecode;
          this.groups = data.obj.groups;
          this.index = '';
          this.Scrollpercent = 0;
        } else {
          if (data.obj.header == 'Yes') {
            this.StoreValues = data.obj.storeValues;
          }
        }
        this.GetDataByMonths(this.currentMonth, this.selectedFilters);
        if (this.StoreValues != '') {
          this.goToFirstPage();
        } else {
          this.NoData = true;
          this.filteredETdetailsData = [];
        }
        const headerdata = {
          title: 'Expense Trend',
          path1: '',
          path2: '',
          path3: '',
          Month: new Date(this.Month),
          filter: this.Filter,
          stores: this.StoreValues,
          storecode: this.StoreCodes,
          Sname: this.StoreName,
          groups: this.groups,
        };
        this.apiSrvc.SetHeaderData({
          obj: headerdata,
        });
      }
    });
    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Expense Trend') return;

      this.SFstate = obj.state;

      if (!obj.state) return;

      if (this.ExpenseTrend) {
        this.exportAsXLSX();
      } else if (this.ExpenseTrendDetails) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Expense Trend') return;

      if (obj.statePrint) {
        this.printPDF();
      }
    });

    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Expense Trend') return;

      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.apiSrvc.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Expense Trend') return;

      if (!obj.stateEmailPdf) return;

      if (this.ExpenseTrend) {
        this.sendEmailData(obj.Email, obj.notes, obj.from);
      } else if (this.ExpenseTrendDetails) {
        this.sendEmailDataDetails(obj.Email, obj.notes, obj.from);
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
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });

    /* ================= TITLE ================= */
    doc.setFontSize(14);
    doc.text('Expense Trend Report', 14, 12);

    /* ================= FILTER INFO ================= */
    doc.setFontSize(10);

    const startDate = this.datepipe.transform(this.dates[0], 'MMM yyyy');
    const endDate = this.datepipe.transform(this.dates[11], 'MMM yyyy');

    // doc.text(`Date Range: ${startDate} - ${endDate}`, 14, 18);
    // doc.text(`Stores: ${this.storedisplayname || 'All Stores'}`, 14, 23);

    /* ================= HEADER ================= */
    const header = [
      `${startDate} - ${endDate}`,
      'YTD',
      'PYTD',
      'AVG',
      ...this.ExpenseTrendKeys
    ];

    /* ================= DATA ================= */
    const rows: any[] = [];

    (this.ExpenseTrendByStoreMonth || []).forEach((d: any) => {

      const row = [
        d.DISPLAY_LABLE,
        d.YTD ?? '-',
        d.PYTD ?? '-',
        d.AVG ?? '-'
      ];

      this.ExpenseTrendKeys.forEach((key: any) => {
        row.push(
          d.DISPLAYHEAD_FLAG === 1
            ? (d[key] ?? '')
            : (d[key] === '' || d[key] == null ? '-' : Number(d[key]))
        );
      });

      rows.push(row);
    });


    autoTable(doc, {
      startY: 18,
      head: [header],
      body: rows,
      theme: 'grid',

      styles: {
        fontSize: 8,
        cellPadding: 2,
        halign: 'right',
        textColor: [20, 20, 20],
        lineWidth: 0.3,
        lineColor: [200, 200, 200]
      },

      headStyles: {
        fillColor: [5, 84, 239],
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        halign: 'center',
        lineWidth: 0.3,
        lineColor: [200, 200, 200]
      },

      didParseCell: (data: any) => {

        if (data.section === 'head') {
          data.cell.styles.fillColor = [5, 84, 239];
          data.cell.styles.textColor = [255, 255, 255];
          data.cell.styles.fontStyle = 'bold';

          data.cell.styles.halign = data.column.index === 0 ? 'left' : 'center';
          return;
        }

        const rowData = this.ExpenseTrendByStoreMonth[data.row.index];
        if (!rowData) return;

        if (data.column.index === 0) {
          data.cell.styles.halign = 'left';
        } else {
          data.cell.styles.halign = 'right';
        }

        if (rowData.DISPLAYHEAD_FLAG === 1) {
          if (!data.cell.raw || data.cell.raw === '-') {
            data.cell.text = [''];
            return;
          }
        } else {
          if (!data.cell.raw || data.cell.raw === '-') {
            data.cell.text = ['-'];
            data.cell.styles.halign = 'center';
            return;
          }
        }
        const isAdvertisingRow = rowData.DISPLAY_LABLE === 'Advertising Rebates';
        const isPercentRow = rowData.DISPLAY_LABLE?.includes('%');

        if (data.column.index > 0) {

          let val = parseFloat(data.cell.raw);

          if (!isNaN(val)) {

            /* 🔹 PERCENT */
            if (isPercentRow) {

              if (isAdvertisingRow) {
                data.cell.text = val < 0
                  ? [`(${Math.abs(val)}%)`]
                  : [`${val}%`];
              } else {
                data.cell.text = [`${val}%`];
              }

            } else {

              const rounded = Math.round(val);
              const formatted = Math.abs(rounded).toLocaleString();

              /* 🔥 ACCOUNTING FORMAT */
              if (isAdvertisingRow) {
                data.cell.text = rounded < 0
                  ? [`($${formatted})`]   // ✅ HERE
                  : [`$${formatted}`];
              } else {
                data.cell.text = [`$${formatted}`];
              }

            }
          }
        }
        if (rowData.DISPLAYHEAD_FLAG === 1) {
          Object.values(data.row.cells).forEach((cell: any) => {
            cell.styles.fillColor = [217, 231, 255];
            cell.styles.fontStyle = 'bold';
            cell.styles.textColor = [0, 0, 0];
          });
        }

        if (rowData.ISHEAD_TOTAL === 'Y') {
          data.cell.styles.fontStyle = 'bold';
        }

        if (
          data.row.index % 2 === 0 &&
          rowData.DISPLAYHEAD_FLAG !== 1 &&
          rowData.ISHEAD_TOTAL !== 'Y'
        ) {
          data.cell.styles.fillColor = [245, 247, 250];
        }
      }
    });

    return doc;
  }
  generatePDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save(`Expense Trend.pdf`);
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
  closeReport() {
    this.Report = '';
  }
  getStores() {
    this.selectedstorevalues = [];
    this.stores = JSON.parse(localStorage.getItem('Stores')!);
  }

  exportAsXLSX(): void {

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Expense Trend');

    /* ================= FILTER SECTION ================= */
    const filterRowCount = this.addExcelFiltersSection(worksheet);

    /* ================= FREEZE ================= */
    worksheet.views = [{
      state: 'frozen',
      ySplit: filterRowCount + 1
    }];

    const ExpenseTrendByStoreMonth = this.ExpenseTrendByStoreMonth.map((e: any) => ({ ...e }));

    const StartDate = this.datepipe.transform(this.dates[0], 'MMM yyyy');
    const EndDate = this.datepipe.transform(this.dates[11], 'MMM yyyy');

    const Header = [StartDate + ' - ' + EndDate, 'YTD', 'PYTD', 'AVG'];

    for (let i = 0; i < this.ExpenseTrendKeys.length; i++) {
      Header.push(this.ExpenseTrendKeys[i]);
    }

    /* ================= HEADER ================= */
    const headerRow = worksheet.addRow(Header);

    headerRow.eachCell((cell: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0554EF' }
      };

      cell.font = {
        name: 'Calibri',
        size: 11,
        bold: true,
        color: { argb: 'FFFFFFFF' }
      };

      cell.alignment = { horizontal: 'center', vertical: 'middle' };

      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    /* ================= DATA ================= */
    ExpenseTrendByStoreMonth.forEach((d: any) => {

      let Obj = [
        d.DISPLAY_LABLE,
        d.YTD ?? '-',
        d.PYTD ?? '-',
        d.AVG ?? '-'
      ];

      this.ExpenseTrendKeys.forEach((e: any) => {
        Obj.push(
          d.DISPLAYHEAD_FLAG === 1
            ? (d[e] ?? '')
            : (d[e] === '' || d[e] === null ? '-' : Number(d[e]))
        );
      });

      const row = worksheet.addRow(Obj);

      row.eachCell((cell: any, colNumber: number) => {

        cell.font = { name: 'Calibri', size: 11 };

        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };

        // Alignment
        if (colNumber === 1) {
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        } else {
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        }

        // Number format + negative red
        if (colNumber >= 2 && typeof cell.value === 'number') {
          cell.numFmt = '_("$"* #,##0_);[Red]_("$"* -#,##0_);_("$"* "-"_);_(@_)';
        }

        // Remove '-' for header rows
        if (d.DISPLAYHEAD_FLAG === 1) {
          if (!cell.value || cell.value === '-') {
            cell.value = '';
          }
        } else {
          if (!cell.value || cell.value === '-') {
            cell.value = '-';
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
          }
        }
      });

      /* ===== ALTERNATE ROW ===== */
      if (row.number % 2 === 0 && d.DISPLAYHEAD_FLAG !== 1) {
        row.eachCell((cell: any) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF5F7FA' }
          };
        });
      }

      /* ===== HEAD ROW ===== */
      if (d.DISPLAYHEAD_FLAG === 1) {
        row.eachCell((cell: any) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'D9E7FF' }
          };

          cell.font = {
            name: 'Calibri',
            size: 11,
            bold: true,
            color: { argb: 'FF000000' }
          };
        });
      }

      /* ===== TOTAL ROW ===== */
      if (d.ISHEAD_TOTAL === 'Y') {
        row.eachCell((cell: any) => {
          cell.font = {
            name: 'Calibri',
            size: 11,
            bold: true
          };

          cell.fill = undefined;

          cell.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
          };
        });
      }

    });

    /* ================= COLUMN WIDTH ================= */
    worksheet.getColumn(1).width = 30;

    for (let i = 2; i <= Header.length; i++) {
      worksheet.getColumn(i).width = 18;
    }

    /* ================= EXPORT ================= */
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, 'Expense Trend' + EXCEL_EXTENSION);
    });

  }
  ETdetailsData: any = [];
  currentPage: number = 1;
  itemsPerPage: number = 100;
  maxPageButtonsToShow: number = 3;
  clickedPage: number | null = null;
  filteredETdetailsData: any[] = [];
  selectedDate: any;
  DateType: any;
  DetailsSearchName: any;
  Lable: any;
  searchText: string = '';
  ExpenseTrend: boolean = true;
  ExpenseTrendDetails: boolean = false;
  SubtypeDetailLable: any;
  FinSummaryLable: any;
  MonthDate: any;
  openMonthDetails(
    Object: any,
    date: any,
    ref: any,
    item: any,
    DateMethod: any
  ) {
    this.spinner.show();
    this.scrollPositionStoring = this.scrollCurrentPosition;
    this.ExpenseTrend = false;
    this.ExpenseTrendDetails = true;
    // this.ETdetailsData = [];
    this.DateType = DateMethod;
    this.Lable = ref;
    console.log(date);
    this.MonthDate = date;
    if (DateMethod === 'YTD') {
      this.selectedDate = new Date(date);

    } else if (DateMethod === 'PYTD') {
      const currentDate = new Date(date);
      currentDate.setFullYear(currentDate.getFullYear() - 1);
      this.selectedDate = currentDate;

    } else {
      this.selectedDate = new Date(date);
    }
    this.SubtypeDetailLable = Object.DISPLAY_LABLE;
    this.FinSummaryLable = Object.DISPLAY_PARENTLABLE;
    const Obj = {
      as_ids: this.storeIds.toString(),
      // dept: this.selectedFilters.toString(),
      dept: '',
      subtype: '',
      subtypedetail: Object.DISPLAY_LABLE,
      FinSummary: Object.DISPLAY_PARENTLABLE,
      date: this.datepipe.transform(new Date(this.selectedDate), 'MM-yyyy'),
      accountnumber: ''
    };
    console.log(Obj);
    this.NoData = false;
    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetExpenseTrendDetailsV1', Obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.ETdetailsData = res.response;
          this.filterData();
          console.log('ET Details', this.ETdetailsData);
          this.spinner.hide();
          this.NoData = true;
        }
      });
  }
  expandedIndex: number | null = null;
  FSSubDetailsMap: { [index: number]: any[] } = {};

  GetSubDetails(AcctNo: any, StoreName: any, index: number) {
    if (this.expandedIndex === index) {
      this.expandedIndex = null;
      return;
    }
    this.expandedIndex = index;

    this.spinner.show();
    const Obj = {
      as_ids: StoreName,
      // dept: this.selectedFilters.toString(),
      dept: '',
      subtype: '',
      subtypedetail: this.SubtypeDetailLable,
      FinSummary: this.FinSummaryLable,
      date: this.datepipe.transform(new Date(this.selectedDate), 'MM-yyyy'),
      accountnumber: AcctNo
    };

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetExpenseTrendDetailsV1', Obj)
      .subscribe((res) => {
        this.spinner.hide();
        if (res.status === 200) {
          this.FSSubDetailsMap[index] = res.response;
        }
      });
  }

  backtoWR() {
    this.ExpenseTrend = true;
    this.ExpenseTrendDetails = false;
    this.filteredETdetailsData = [];
    this.ETdetailsData = [];
    this.expandedIndex = null;
    if (this.StoreValues != '') {
      this.goToFirstPage();
    }
    setTimeout(() => {
      if (this.scrollcent && this.scrollcent.nativeElement) {
        this.scrollcent.nativeElement.scrollTop = this.scrollPositionStoring;
      }
    });
  }

  get postingAmountTotal(): number {
    return this.filteredETdetailsData.reduce((total, item) => {
      return total + (item.PostingAmount || 0);
    }, 0);
  }
  getPostingSubAmountTotal(index: number): number {
    const subRows = this.FSSubDetailsMap[index];
    if (!subRows || !Array.isArray(subRows)) {
      return 0;
    }

    return subRows.reduce((total, item) => {
      return total + (item.PostingAmount || 0);
    }, 0);
  }
  filterData() {
    const text = this.searchText.trim().toLowerCase();

    if (!text) {
      this.filteredETdetailsData = [...this.ETdetailsData];
    } else {
      this.filteredETdetailsData = this.ETdetailsData.filter((item: any) =>
        [
          item.StoreName,
          item.AccountNumber,
          item.AccountDescription,
          item.PostingAmount
        ].some(val =>
          val?.toString().toLowerCase().includes(text)
        )
      );
    }

    this.currentPage = 1;
  }
  get paginatedItems() {
    const start = (this.currentPage - 1) * this.itemsPerPage;
    return this.filteredETdetailsData.slice(start, start + this.itemsPerPage);
  }

  getMaxPageNumber(): number {
    return Math.max(1,
      Math.ceil(this.filteredETdetailsData.length / this.itemsPerPage)
    );
  }

  nextPage() {
    if (this.currentPage < this.getMaxPageNumber()) this.currentPage++;
  }

  prevPage() {
    if (this.currentPage > 1) this.currentPage--;
  }

  goToFirstPage() {
    this.currentPage = 1;
  }

  goToLastPage() {
    this.currentPage = this.getMaxPageNumber();
  }

  getStartRecordIndex(): number {
    return (this.currentPage - 1) * this.itemsPerPage + 1;
  }

  getEndRecordIndex(): number {
    return Math.min(
      this.getStartRecordIndex() + this.itemsPerPage - 1,
      this.filteredETdetailsData.length
    );
  }
  onPageSizeChange() {
    this.currentPage = 1;
  }

  openStoresDetails(Object: any, StoreName: any, date: any, item: any) {
    this.index = '';
    console.log(Object);
    console.log(item);
    const myDate = this.datepipe.transform(new Date(date), 'MMM-yyyy');
    let index = this.stores.filter(
      (store: any) => store.DEALER_NAME == StoreName
    );
    this.selectedstorevalues = index[0].AS_ID;
    this.selectedstorename = index[0].DEALER_NAME;
    const DetailsSF = this.ngbmodal.open({
      size: 'xl',
      backdrop: 'static',
    });
    DetailsSF.componentInstance.ETdetails = {
      TYPE: item.LABLEVAL,
      NAME: item.LABLE,
      STORES: this.selectedstorevalues,
      LatestDate: myDate,
      STORENAME: this.selectedstorename,
    };
  }
  ValueFormat: any;
  openGraph(monthname: any, dates: any, Obj: any, SummaryType: any) {
    const modalRef = this.ngbmodal.open(ExpensetrendGraph, {
      size: 'xl',
      backdrop: 'static',
      injector: Injector.create({
        providers: [
          { provide: CurrencyPipe, useClass: CurrencyPipe }
        ],
        parent: this.injector
      })
    });
    const ValueFormat =
      Obj.DISPLAY_LABLE.includes('%') ? 'Percentage' : 'Currency';

    modalRef.componentInstance.ETgraphdetails = {
      ITEM: Obj,
      DATES: dates,
      NAME: Obj.DISPLAY_LABLE,
      ValueFormat,
      STORES: this.StoreValues,
      SUMMARYTYPE: SummaryType,
    };
  }

  openStoreGraph(
    Object: any,
    StoreName: any,
    date: any,
    item: any,
    SummaryType: any
  ) {
    console.log(Object, StoreName, date, item, SummaryType);
    const DetailsSF = this.ngbmodal.open({
      size: 'xl',
      backdrop: 'static',
    });
    var DisplayName = item.DISPLAY_LABLE;
    DetailsSF.componentInstance.ETgraphdetails = {
      ITEM: item,
      TYPE: item.DISPLAY_LABLE,
      NAME: SummaryType,
      SUMMARYTYPE: SummaryType,
    };
  }

  index = '';
  commentobj: any = {};
  close(data: any) {
    console.log(data);
    this.index = '';
  }
  addcmt(data: any) {
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
            console.log(data);
          },
          (reason) => {
            console.log(reason);

            if (reason == 'O') {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
            } else {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
              if (this.Filter == 'VariableTrendsvsBudget') {
                this.GetDataByMonths(this.currentMonth, this.selectedFilters);
              }
            }
            // // on dismiss

            // const Data = {
            //   state: true,
            // };
            // this.apiSrvc.setBackgroundstate({ obj: Data });
            // this.GetData();
          }
        );
    }
    if (data == 'AD') {
      if (this.Filter == 'VariableTrendsvsBudget') {
        this.GetDataByMonths(this.currentMonth, this.selectedFilters);
      }
    }
  }

  ExcelStoreNames: any = [];


  Favreports: any = [];

  isDesc: boolean = false;
  column: string = 'CategoryName';

  sort(property: string, data: any[], state?: any) {
    if (state === undefined) {
      this.isDesc = this.column === property ? !this.isDesc : false;
    }
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    data.sort((a, b) => {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
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
      const pdfBlob = this.generatePDFBlob();

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
      const pdfFile = this.blobToFile(pdfBlob, 'Expense Trend.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Expense Trend');
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
      title: 'Expense Trend',
      filters: [
        {
          label: 'Store',
          value: this.getSelectedStoreNames() || 'All Stores'
        },
        {
          label: 'Group',
          value: this.groupName || ''
        },
        // {
        //   label: 'Department',
        //   value: this.Filter && this.Filter.length
        //     ? this.Filter.join(', ')
        //     : 'All'
        // },
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
  exportToExcel() {

    const FSDetailsData = [...this.filteredETdetailsData];
    const FSSubDetailsMap = this.FSSubDetailsMap;

    let storeValue = '';

    const selectedStoreIds: string[] =
      this.storeIds && this.storeIds.length
        ? this.storeIds.map((id: any) => id.toString())
        : [];

    const allStores: any[] = Array.isArray(this.stores) ? this.stores : [];

    storeValue = allStores
      .filter((s: any) => selectedStoreIds.includes(s.ID.toString()))
      .map((s: any) => s.storename.trim())
      .filter(Boolean)
      .join(', ');

    if (!storeValue && selectedStoreIds.length) {
      storeValue = selectedStoreIds.join(', ');
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Expense Trend Details');

    const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');
    const DateToday = this.datepipe.transform(new Date(), 'MM.dd.yyyy h:mm:ss a');

    /* ================= FREEZE ================= */
    worksheet.views = [{
      state: 'frozen',
      ySplit: 1
    }];
    const totalColumns = 4;

    /* ================= BORDER FUNCTION ================= */
    const applyBorder = (cell: any) => {
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    };

    /* ================= FILTER STYLE ================= */
    const TOTAL_COLUMNS = 4;

    const applyFullRowBorder = (row: any) => {
      for (let i = 1; i <= TOTAL_COLUMNS; i++) {
        const cell = row.getCell(i);

        // ✅ Force cell creation
        if (cell.value === undefined || cell.value === null) {
          cell.value = '';
        }

        // ✅ Apply border
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };

        // Optional alignment
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      }
    };

    /* ================= HEADER SECTION (FIXED LIKE OLD CODE) ================= */

    let rowCount = 0;

    /* ===== TITLE ===== */
    const titleRow = worksheet.addRow(['Expense Trend Details']);
    titleRow.font = { bold: true, size: 14, name: 'Calibri' };
    titleRow.alignment = { horizontal: 'left', vertical: 'middle' };

    // ✅ Merge like your old code (IMPORTANT)
    worksheet.mergeCells(`A${rowCount + 1}:D${rowCount + 1}`);

    // ✅ Apply border to all merged cells
    for (let i = 1; i <= 4; i++) {
      const cell = titleRow.getCell(i);
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    }

    rowCount++;

    /* ===== DATE ===== */
    const dateRow = worksheet.addRow([DateToday]);
    dateRow.font = { name: 'Calibri', size: 11 };

    for (let i = 1; i <= 4; i++) {
      const cell = dateRow.getCell(i);
      if (!cell.value) cell.value = '';
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
    rowCount++;

    /* ===== SELECTED DETAILS ===== */
    const selectedRow = worksheet.addRow(['Selected Details:']);
    selectedRow.font = { bold: true, size: 11, name: 'Calibri' };

    for (let i = 1; i <= 4; i++) {
      const cell = selectedRow.getCell(i);
      if (!cell.value) cell.value = '';
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
    rowCount++;

    /* ===== TYPE ===== */
    const typeRow = worksheet.addRow(['Type:', this.Lable]);
    typeRow.getCell(1).font = { bold: true, name: 'Calibri' };

    for (let i = 1; i <= 4; i++) {
      const cell = typeRow.getCell(i);
      if (!cell.value) cell.value = '';
      cell.font = { name: 'Calibri', size: 11 };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
    rowCount++;

    /* ===== DATE FILTER ===== */
    const dateFilterRow = worksheet.addRow(['Date:', this.selectedDate]);
    dateFilterRow.getCell(1).font = { bold: true, name: 'Calibri' };

    for (let i = 1; i <= 4; i++) {
      const cell = dateFilterRow.getCell(i);
      if (!cell.value) cell.value = '';
      cell.font = { name: 'Calibri', size: 11 };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
    rowCount++;

    /* ===== STORE ===== */
    const storeRow = worksheet.addRow(['Store:', storeValue]);
    storeRow.getCell(1).font = { bold: true, name: 'Calibri' };

    for (let i = 1; i <= 4; i++) {
      const cell = storeRow.getCell(i);
      if (!cell.value) cell.value = '';
      cell.font = { name: 'Calibri', size: 11 };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
    rowCount++;

    /* ===== EMPTY ROW ===== */
    const emptyRow = worksheet.addRow([]);
    for (let i = 1; i <= 4; i++) {
      const cell = emptyRow.getCell(i);
      cell.value = '';
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
    rowCount++;
    /* ================= GRID HEADER ================= */
    const headers = [
      'STORE NAME',
      'ACCOUNT NUMBER',
      'DESCRIPTION',
      'BALANCE',
    ];

    const headerRow = worksheet.addRow(headers);

    headerRow.eachCell((cell: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9E7FF' }
      };

      cell.font = {
        name: 'Calibri',
        size: 11,
        bold: true,
        color: { argb: 'FF000000' }
      };

      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      applyBorder(cell);
    });

    let totalBalance = 0;
    const columnWidths = new Array(headers.length).fill(10);

    /* ================= MAIN DATA ================= */
    FSDetailsData.forEach((item: any, index: number) => {

      const rowData = [
        item.StoreName || '-',
        item.AccountNumber || '-',
        item.AccountDescription || '-',
        item.PostingAmount ?? '-',
      ];

      const mainRow = worksheet.addRow(rowData);

      mainRow.eachCell((cell: any, colNumber: number) => {

        cell.font = { name: 'Calibri', size: 11 };

        if (colNumber === 4 && typeof cell.value === 'number') {
          cell.numFmt = '_("$"* #,##0_);[Red]_("$"* -#,##0_);_("$"* "-"_);_(@_)';
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        } else {
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        }

        if (!cell.value || cell.value === '') {
          cell.value = '-';
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }

        applyBorder(cell);
      });

      if (typeof item.PostingAmount === 'number') {
        totalBalance += item.PostingAmount;
      }

      if (index % 2 !== 0) {
        mainRow.eachCell((cell: any) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF5F7FA' }
          };
        });
      }

      rowData.forEach((val, i) => {
        const length = val?.toString().length || 0;
        columnWidths[i] = Math.max(columnWidths[i], length);
      });

      /* ================= SUB DETAILS ================= */
      const subDetails = FSSubDetailsMap[index];

      if (subDetails?.length && this.expandedIndex != null) {

        const subheaderRow = worksheet.addRow([
          'Control', 'Date', 'Detail Description', 'Balance'
        ]);

        subheaderRow.eachCell((cell: any) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '0554EF' }
          };
          cell.font = {
            name: 'Calibri',
            size: 11,
            bold: true,
            color: {
              argb: 'FFFFFFFF'
            }
          };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          applyBorder(cell);
        });

        let subTotal = 0;

        subDetails.forEach((sub: any) => {

          const subRow = worksheet.addRow([
            sub.Control || '-',
            sub.AccountingDate || '-',
            sub.DetailDescription || '-',
            sub.PostingAmount ?? '-',
          ]);

          subRow.eachCell((cell: any, colNumber: number) => {

            cell.font = { name: 'Calibri', size: 11 };

            if (colNumber === 4 && typeof cell.value === 'number') {
              cell.numFmt = '_("$"* #,##0_);[Red]_("$"* -#,##0_);_("$"* "-"_);_(@_)';
              cell.alignment = { horizontal: 'right', vertical: 'middle' };
            } else {
              cell.alignment = { horizontal: 'left', vertical: 'middle' };
            }

            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'D9E7FF' }
            };

            applyBorder(cell);
          });

          if (typeof sub.PostingAmount === 'number') {
            subTotal += sub.PostingAmount;
            totalBalance += sub.PostingAmount;
          }
        });

        const subTotalRow = worksheet.addRow(['', '', 'Sub Total:', subTotal]);

        subTotalRow.eachCell((cell: any, colNumber: number) => {
          cell.font = { name: 'Calibri', size: 11, bold: true };

          if (colNumber === 4) {
            cell.numFmt = '_("$"* #,##0_);[Red]_("$"* -#,##0_);_("$"* "-"_);_(@_)';
            cell.alignment = { horizontal: 'right', vertical: 'middle' };
          }

          applyBorder(cell);
        });
      }
    });

    /* ================= TOTAL ================= */
    const totalRow = worksheet.addRow(['', '', 'Total:', this.postingAmountTotal]);

    totalRow.eachCell((cell: any, colNumber: number) => {
      cell.font = { name: 'Calibri', size: 11, bold: true };

      if (colNumber === 4) {
        cell.numFmt = '_("$"* #,##0_);[Red]_("$"* -#,##0_);_("$"* "-"_);_(@_)';
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
      }

      applyBorder(cell);
    });

    /* ================= FORCE FULL SHEET BORDERS ================= */
    worksheet.eachRow((row: any) => {
      for (let i = 1; i <= totalColumns; i++) {
        const cell = row.getCell(i);

        if (!cell.value) cell.value = '';

        applyBorder(cell);
      }
    });

    /* ================= COLUMN WIDTH ================= */
    columnWidths.forEach((width, i) => {
      worksheet.getColumn(i + 1).width = Math.max(20, width + 2);
    });

    /* ================= EXPORT ================= */
    workbook.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, `Expense_Trend_Details`);
    });
  }

  sendEmailDataDetails(Email: any, notes: any, from: any) {
    this.spinner.show();
    const printContents = document.getElementById(
      'ExpenseTrendDetailsDownload'
    )!.innerHTML;
    const iframe = document.createElement('iframe');

    // Make the iframe invisible
    iframe.style.position = 'absolute';
    iframe.style.width = '0px';
    iframe.style.height = '0px';
    iframe.style.border = 'none';

    document.body.appendChild(iframe);

    const doc = iframe.contentWindow?.document;
    if (!doc) {
      console.error('Failed to create iframe document');
      return;
    }
    doc.open();
    doc.write(`
        <html>
            <head>
            <title>Expense Trend</title>
                 <style>
                 @font-face {
                  font-family: 'GothamBookRegular';
                  src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                @font-face {
                  font-family: 'Roboto';
                  src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                @font-face {
                  font-family: 'RobotoBold';
                  src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                       url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
                }
                                            .comment {
                        width: fit-content;
                        background-color: #d7d7d7;
                        border-radius: 0 15px 0 15px;
                        padding: .2rem .5rem;
                        margin-left: 2rem;
                      }
                      .bdr{
                        border:1px solid white !important
                      }
                                            .notes { 
                        width: fit-content;  
                        padding: 0.2rem 0.5rem;
                        margin-right: 3rem;
                        float: inline-end;
                      }
                              .performance-scorecard-details {
                        display: flex;
                        flex-direction: column;
                        height: auto;
                        /* Adjust based on your needs */
                        width: 100%;
                      }
                      .performance-scorecard-details .table > :not(:first-child) {
                        border-top: 0px solid #ffa51a;
                      }
                      .performance-scorecard-details .table {
                        text-align: center;
                        text-transform: capitalize;
                        border: transparent;
                        width: 100%;
                      }
                      .performance-scorecard-details .table th, .performance-scorecard-details .table td {
                        white-space: nowrap;
                        vertical-align: top;
                      }
                      .performance-scorecard-details .table th:first-child, .performance-scorecard-details .table td:first-child {
                        left: 0;
                        z-index: 1;
                      }
                      .performance-scorecard-details .table tr:nth-child(odd) td:first-child, .performance-scorecard-details .table tr:nth-child(odd) td:nth-child(2) {
                        background-color: #e9ecef;
                      }
                      .performance-scorecard-details .table tr:nth-child(even) td:first-child, .performance-scorecard-details .table tr:nth-child(even) td:nth-child(2) {
                        background-color: #fff;
                      }
                      .performance-scorecard-details .table tr:nth-child(odd) {
                        background-color: #e9ecef;
                      }
                      .performance-scorecard-details .table tr:nth-child(even) {
                        background-color: #fff;
                      }
                      .performance-scorecard-details .table .spacer {
                        background-color: #cfd6de !important;
                        border-left: 1px solid #cfd6de !important;
                        border-bottom: 1px solid #cfd6de !important;
                        border-top: 1px solid #cfd6de !important;
                      }
                      .performance-scorecard-details .table .hidden {
                        display: none !important;
                      }
                      .performance-scorecard-details .table .bdr-rt {
                        border-right: 1px solid #abd0ec;
                      }
                      .performance-scorecard-details .table thead {
                        position: sticky;
                        top: 0;
                        z-index: 99;
                        font-family: 'FaktPro-Bold';
                        font-size: 0.8rem;
                      }
                      .performance-scorecard-details .table thead th {
                        padding: 5px 10px;
                        margin: 0px;
                        border-right: 1px solid #abd0ec;
                      }
                      .performance-scorecard-details .table thead .bdr-btm {
                        border-bottom: #005fa3;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(1) {
                        background-color: #fff !important;
                        color: #000;
                        text-transform: uppercase;
                        border-bottom: #cfd6de;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(1) {
                        background-color: #fff !important;
                        color: #000;
                        text-transform: uppercase;
                        border-bottom: #cfd6de;
                        box-shadow: inset 0 1px 0 0 #cfd6de;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(1) th {
                        box-shadow: inset 0 -2px 0 #337ab7;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(3) {
                        background-color: #fff !important;
                        color: #000;
                        text-transform: uppercase;
                        border-bottom: #cfd6de;
                        box-shadow: inset 0 1px 0 0 #cfd6de;
                      }
                      .performance-scorecard-details .table thead tr:nth-child(3) th :nth-child(1) {
                        background-color: #337ab7 !important;
                        color: #fff;
                      }
                      .performance-scorecard-details .table tbody {
                        font-family: 'FaktPro-Normal';
                        font-size: 0.9rem;
                      }
                      .performance-scorecard-details .table tbody td {
                        padding: 2px 10px;
                        margin: 0px;
                        border: 1px solid #cfd6de;
                      }
                      .performance-scorecard-details .table tbody tr {
                        border-bottom: 1px solid #37a6f8;
                        border-left: 1px solid #37a6f8;
                      }
                      .performance-scorecard-details .table tbody td:first-child {
                        text-align: start;
                        box-shadow: inset -1px 0 0 0 #cfd6de;
                      }
                      .performance-scorecard-details .table tbody .sub-title {
                        font-size: 0.8rem !important;
                      }
                      .performance-scorecard-details .table tbody .sub-subtitle {
                        font-size: 0.7rem !important;
                      }
                      .performance-scorecard-details .table tbody .alignright {
                        text-align: right;
                        padding-right: 1rem;
                      }
                      .performance-scorecard-details .table tbody .alignleft {
                        text-align: left;
                        padding-left: 1rem;
                      }
                      .performance-scorecard-details .table tbody .text-bold {
                        font-family: 'FaktPro-Bold';
                      }
                      .performance-scorecard-details .table tbody .darkred-bg {
                        background-color: #282828 !important;
                        color: #fff;
                      }
                      .performance-scorecard-details .table tbody .lightblue-bg {
                        background-color: #646e7a !important;
                        color: #fff;
                      }
                      .performance-scorecard-details .table tbody .gold-bg {
                        background-color: #ffa51a;
                        color: #fff;
                      }
                 </style>
            </head>
            <body id='content'>
                ${printContents}
            </body>
        </html>
    `);

    doc.close();

    const div = doc.getElementById('content');
    if (!div) {
      console.error('Element not found');
      return;
    }

    const options = {
      logging: true,
      allowTaint: false,
      useCORS: true,
      scale: 1, // Adjust scale to fit the page better
    };

    html2canvas(div, options)
      .then((canvas) => {
        let imgWidth = 285;
        let pageHeight = 204;
        let imgHeight = (canvas.height * imgWidth) / canvas.width;
        let heightLeft = imgHeight;
        const contentDataURL = canvas.toDataURL('image/png');
        let pdfData = new jsPDF('l', 'mm', 'a4', true);
        let position = 5;
        pdfData.addImage(
          contentDataURL,
          'PNG',
          5,
          position,
          imgWidth,
          imgHeight,
          undefined,
          'FAST'
        );
        heightLeft -= pageHeight;
        while (heightLeft >= 0) {
          position = heightLeft - imgHeight;
          pdfData.addPage();
          pdfData.addImage(
            contentDataURL,
            'PNG',
            5,
            position,
            imgWidth,
            imgHeight,
            undefined,
            'FAST'
          );
          heightLeft -= pageHeight;
        }

        const pdfBlob = pdfData.output('blob');
        const pdfFile = this.blobToFile(pdfBlob, 'Expense Trend Details.pdf');
        const formData = new FormData();
        formData.append('to_email', Email);
        formData.append('subject', 'Expense Trend Details');
        formData.append('file', pdfFile);
        formData.append('notes', notes);
        formData.append('from', from);
        this.apiSrvc
          .postmethod(this.comm.routeEndpoint + 'mail', formData)
          .subscribe(
            (res: any) => {
              console.log('Response:', res);
              if (res.status === 200) {

                this.toast.show(res.response, 'success', 'Success');
              } else {

                this.toast.show('Invalid Details.', 'danger', 'Error');
              }
            },
            (error) => {
              console.error('Error:', error);
            }
          );
      })
      .catch((error) => {
        console.error('html2canvas error:', error);
      })
      .finally(() => {
        this.spinner.hide();
        // popupWin.close();
      });
  }
}

