import { Component, Injector, HostListener } from '@angular/core';
import { Api } from '../../../../Core/Providers/Api/api';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { Subscription } from 'rxjs';
import { Workbook } from 'exceljs';
import FileSaver from 'file-saver';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
import numeral from 'numeral';
import pdfMake from 'pdfmake/build/pdfmake';
import pdfFonts from 'pdfmake/build/vfs_fonts';
(pdfMake as any)['vfs'] = (pdfFonts as any)['vfs'];
import autoTable from 'jspdf-autotable';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { common } from '../../../../common';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {

  GLData: any = [];
  grandTotal: number = 0;
  accountNumber: string = '';
  NoData = '';
  actionType: any = 'N'
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
  activePopover: number = -1;

  month!: Date;
  DuplicatDate!: Date;
  minDate!: Date;
  maxDate!: Date;
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month'
  };
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  constructor(public apiSrvc: Api, private ngbmodalActive: NgbActiveModal,
    private toast: ToastService, private injector: Injector, public shared: Sharedservice,
    private comm: common) {

    let today = new Date();
    let enddate = new Date(today.setDate(today.getDate() - 1));
    this.month = new Date(enddate.setMonth(enddate.getMonth() - 1))
    this.maxDate = new Date();
    this.minDate = new Date();
    this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    this.maxDate.setMonth(this.maxDate.getMonth() - 1);
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
    this.shared.setTitle(this.shared.common.titleName + '-GL Lookup')

    this.setHeaderData()
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

  onOpenCalendar(container: any) {
    container.setViewMode('month');
    container.monthSelectHandler = (event: any): void => {
      container.value = event.date;
      this.month = event.date;
      return;
    };
  }
  changeDate(e: any) {
    console.log(e);
    this.month = e;
  }
  setHeaderData() {
    const data = {
      title: 'GL Lookup',
      Month: this.month,
      store: this.storeIds,
      groups: this.groupId,
    };
    this.apiSrvc.SetHeaderData({
      obj: data,
    });
  }
  searchGLData() {
    this.GLData = [];
    const acc = (this.accountNumber || '').trim();


    this.shared.spinner.show();
    const obj = {
      Store: this.storeIds.toString(),
      Account_Number: acc,
      Date: this.shared.datePipe.transform(this.month, 'yyyy-MM-dd'),

    };

    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetGLLookup', obj).subscribe(
      (res: any) => {
        this.shared.spinner.hide();
        if (res.status === 200 && res.response.length > 0) {
          this.GLData = res.response;
          this.NoData = '';
          this.grandTotal = this.GLData.reduce((sum: any, r: any) => {
            return sum + (r.postingamount ? Number(r.postingamount) : 0);
          }, 0);

        } else {
          this.GLData = [];
          this.NoData = 'No Data Found!!';
        }
      },
      () => {
        this.shared.spinner.hide();
        this.toast.show('Error fetching GL data', 'danger', 'Error');
        this.NoData = 'No Data Found!!';

      }
    );
  }
  getSubTotal(rows: any[]): number {
    if (!rows) return 0;

    return rows.reduce((sum, r) => {
      return sum + (r.postingamount ? Number(r.postingamount) : 0);
    }, 0);
  }
  expandedGLRows: any[] = [];

  expandGL(i: number, row: any) {
    if (this.expandedGLRows.includes(i)) {
      this.expandedGLRows = this.expandedGLRows.filter(x => x !== i);
      return;
    }
    this.expandedGLRows.push(i);
    if (!row.childRows) {
      try {
        row.childRows = JSON.parse(row.AccountDetails);
      } catch {
        row.childRows = [];
      }
    }
  }


  isDesc: boolean = false;
  column: string = '';

  sort(property: any, data: any, block: any) {

    this.isDesc = !this.isDesc;  // toggle ASC/DESC
    this.column = property;
    let direction = this.isDesc ? 1 : -1;

    // ===================== GL MAIN BLOCK ("M") ===================== //
    if (block === 'M') {
      let fakedata = [...data];
      fakedata.sort((a: any, b: any) => {
        if (a[property] < b[property]) return -1 * direction;
        if (a[property] > b[property]) return 1 * direction;
        return 0;
      });
      this.GLData = [...fakedata];
      this.expandedGLRows = []
      return;
    }

    // ===================== GL CHILD TABLE BLOCK ("GL") ===================== //
    if (block === 'GL') {
      if (!data || data.length === 0) return;
      data.sort((a: any, b: any) => {
        let valA = a[property];
        let valB = b[property];
        if (property === 'accountingdate') {
          valA = new Date(valA);
          valB = new Date(valB);
        }

        valA = valA ?? '';
        valB = valB ?? '';

        if (valA < valB) return -1 * direction;
        if (valA > valB) return 1 * direction;
        return 0;
      });

      return;
    }

    // ===================== DEFAULT SIMPLE SORT ===================== //
    data.sort((a: any, b: any) => {
      if (a[property] < b[property]) return -1 * direction;
      if (a[property] > b[property]) return 1 * direction;
      return 0;
    });
  }
  viewreport() {
    console.log(this.storeIds);
    this.activePopover = -1
    if (this.storeIds.length == 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    }
    if (!this.accountNumber) {
      this.toast.show('Please enter an account number', 'warning', 'Warning');
      return;
    }

    else {
      if (this.storeIds != '') {
        this.expandedGLRows = [];
        this.setHeaderData()
        this.searchGLData();
        this.actionType = 'Y'
      } else {
        // this.NoData = '';
      }
    }
  }
  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'GL Lookup') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })

    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'GL Lookup') return;
      if (obj.state) {
        this.exportAsXLSX();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'GL Lookup') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'GL Lookup') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'GL Lookup') return;
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

  private createPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    doc.setFontSize(14);
    doc.setFont('helvetica', 'bold');
    doc.text('GL Lookup', 14, 10);

    const pageW = doc.internal.pageSize.getWidth();
    const margin = 6;
    const colCount = 4;
    const colWidth = (pageW - margin * 2) / colCount;

    const columnStyles: any = {};
    for (let i = 0; i < colCount; i++) {
      columnStyles[i] = { cellWidth: colWidth };
    }
    const head = [[
      'ACCOUNT NUMBER',
      'ACCOUNT NAME',
      'STORE',
      'POSTING AMOUNT'
    ]];

    const rows: any[] = [];

    this.GLData.forEach((item: any, index: number) => {

      rows.push([
        item.Accountnumber || '-',
        item.AccountName || '-',
        item.Store || '-',
        item.postingamount ?? 0
      ]);

    });
    rows.push([
      '',
      '',
      'TOTAL',
      this.grandTotal
    ]);
    autoTable(doc, {
      startY: 16,
      head,
      body: rows,
      theme: 'grid',
      margin: { left: margin, right: margin },
      columnStyles,

      styles: {
        fontSize: 9,
        cellPadding: 3,
        textColor: [0, 0, 0],
        lineWidth: 0.2,
        lineColor: [180, 180, 180]
      },

      headStyles: {
        fillColor: [217, 231, 255],
        textColor: [0, 0, 0],
        fontStyle: 'bold',
        halign: 'center'
      },

      didParseCell: (data: any) => {
        if (data.section === 'body') {
          const col = data.column.index;
          if (col === 3) {
            data.cell.styles.halign = 'right';
          } else {
            data.cell.styles.halign = 'left';
          }

          let val = Number(data.cell.raw);
          if (col === 3 && !isNaN(val)) {

            if (val === 0) {
              data.cell.text = ['-'];
              data.cell.styles.halign = 'center';
              return;
            }

            const formatted = Math.round(val).toLocaleString('en-US');

            data.cell.text = [
              val < 0 ? `-$${Math.abs(Math.round(val)).toLocaleString('en-US')}` : `$${formatted}`
            ];
            if (val < 0) {
              data.cell.styles.textColor = [220, 53, 69];
            }
          }
          const rowText = data.row.raw?.[2];

          if (rowText && rowText.toString().toLowerCase() === 'total') {
            data.cell.styles.fillColor = [217, 231, 255];
            data.cell.styles.fontStyle = 'bold';
          }
        }
        if (data.section === 'body' && data.row.index % 2 === 0) {
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
      doc.save('GL Lookup.pdf');
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
      const pdfFile = this.blobToFile(pdfBlob, 'GL Lookup.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'GL Lookup');
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
  getSelectedStoreNames(): string {
    if (!this.storeIds || this.storeIds.length === 0) return '';

    const ids = this.storeIds.toString().split(',');

    const selectedStores = this.stores.filter((s: any) =>
      ids.includes(s.ID.toString())
    );

    return selectedStores.map((s: any) => s.storename).join(', ');
  }
  getReportFilters(): { title: string; filters: any[] } {

    const formattedMonth = this.month
      ? this.shared.datePipe.transform(this.month, 'MMMM yyyy')
      : '';

    return {
      title: 'GL Lookup',
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
          label: 'Account Number',
          value: this.accountNumber || 'All Accounts'
        },
        {
          label: 'Month',
          value: formattedMonth || ''
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
  exportAsXLSX() {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet("GL Lookup");
    const DATE_EXTENSION = this.shared.datePipe.transform(new Date(), 'MMddyyyy');
    const TOTAL_COLUMNS = 4;
    const filterRowCount = this.addExcelFiltersSection(worksheet);
    const headers = [
      "ACCOUNT NUMBER",
      "ACCOUNT NAME",
      "STORE",
      "POSTING AMOUNT"
    ];

    const headerRow = worksheet.addRow(headers);

    headerRow.eachCell((cell: any) => {

      if (typeof cell.value === 'string') {
        cell.value = cell.value.toUpperCase();
      }

      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFD9E7FF" }
      };

      cell.font = {
        name: "Calibri",
        size: 11,
        bold: true,
        color: { argb: "FF000000" }
      };

      cell.alignment = { horizontal: "center", vertical: "middle" };

      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" }
      };
    });

    worksheet.views = [{ state: "frozen", ySplit: 7 }];


    worksheet.columns = [
      { key: "Accountnumber", width: 25 },
      { key: "AccountName", width: 30 },
      { key: "Store", width: 30 },
      { key: "postingamount", width: 20 }
    ];

    let dataRowIndex = filterRowCount + 2;

    this.GLData.forEach((item: any) => {

      const row = worksheet.addRow({
        Accountnumber: item.Accountnumber || "-",
        AccountName: item.AccountName || "-",
        Store: item.Store || "-",
        postingamount: item.postingamount ?? 0
      });

      row.eachCell((cell: any, colNumber: number) => {

        cell.font = { name: "Calibri", size: 11 };

        if (colNumber === 4 && typeof cell.value === "number") {
          cell.numFmt = '_("$"* #,##0.00_);[Red]_("$"* -#,##0.00_);_("$"* "-"_);_(@_)';
          cell.alignment = { horizontal: "right", vertical: "middle" };
        } else {
          cell.alignment = { horizontal: "left", vertical: "middle" };
        }

        if (!cell.value) cell.value = "-";

        cell.border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" }
        };
      });

      // ✅ Zebra rows
      if (dataRowIndex % 2 === 0) {
        row.eachCell((cell: any) => {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFF5F7FA" }
          };
        });
      }

      dataRowIndex++;
    });

    /* ================= TOTAL ================= */

    const totalRow = worksheet.addRow([
      "",
      "",
      "TOTAL",
      this.grandTotal
    ]);

    totalRow.eachCell((cell: any, colNumber: number) => {

      cell.font = { name: "Calibri", size: 11, bold: true };

      if (colNumber === 4) {
        cell.numFmt = '_("$"* #,##0.00_);[Red]_("$"* -#,##0.00_);_("$"* "-"_);_(@_)';
        cell.alignment = { horizontal: "right", vertical: "middle" };
      }

      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFD9E7FF" }
      };

      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" }
      };
    });

    /* ================= EXPORT ================= */

    workbook.xlsx.writeBuffer().then((buffer) => {
      FileSaver.saveAs(
        new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }),
        `GLLookup_Report_${DATE_EXTENSION}.xlsx`
      );
    });
  }



  // ---------------------------  STORE NAME HELPER ---------------------------
  getStoreNames(): string {
    const allStores = this.shared.common.groupsandstores.flatMap((g: any) => g.Stores);

    // REMOVE DUPLICATES
    const uniqueIds = [...new Set(this.storeIds)];

    // Check if all stores selected
    if (uniqueIds.length === allStores.length) {
      return "All Stores";
    }


    const selected = allStores
      .filter((s: any) => uniqueIds.includes(s.ID))
      .map((s: any) => s.storename);

    return selected.join(", ");
  }
}
