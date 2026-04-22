import {
  Component,
  ElementRef,
  Injector,
  Input,
  OnInit,
  SimpleChanges,
  ViewChild,
  HostListener
} from '@angular/core';
import { ActivatedRoute, Router } from '@angular/router';
import { NgxSpinnerService } from 'ngx-spinner';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { Title } from '@angular/platform-browser';
import { CommonModule, CurrencyPipe, DatePipe, NgStyle } from '@angular/common';
import * as FileSaver from 'file-saver';
import { Workbook } from 'exceljs';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import html2canvas from 'html2canvas';
import { Subscription } from 'rxjs';
import { Api } from '../../../../Core/Providers/Api/api';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { environment } from '../../../../../environments/environment';
import { common } from '../../../../common';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { FormsModule } from '@angular/forms';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  imports: [CommonModule, SharedModule, BsDatepickerModule, FormsModule],
  standalone: true,
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {

  LendersData: any;
  filteredData: any[] = [];
  selectedstrid = 0;
  selectedDate: Date = new Date();
  currentMonth!: Date;
  Month: any = '';
  groups: any = 1;
  StoreName: any;
  selectedstorevalues: any = [];
  gridvisibility: any;
  bsRangeValue!: Date[];
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = '0';
  stores: any = []
  activePopover: number = -1;

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  constructor(
    public apiSrvc: Api,
    private spinner: NgxSpinnerService,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private toast: ToastService,
    private title: Title,
    private datepipe: DatePipe,
    private comm: common,
    public shared: Sharedservice,
  ) {
    const data = {
      title: 'Lender Mapping',
      path1: '',
      path2: '',
      path3: '',
      stores: this.storeIds,
      groups: this.groups,
      count: 0
    };
    this.apiSrvc.SetHeaderData({
      obj: data,
    });
    this.title.setTitle(this.comm.titleName + '-Lender Mapping');
    this.GetLenders(this.selectedstrid);
  }
  NoData: any = false;
  searchText: any = '';
  GetLenders(strid: any) {
    this.LendersData = [];
    this.NoData = false;
    this.spinner.show();
    const obj = {
      store_id: strid,
      accountname: this.searchText,
      Lendertype: '0',
      LenderCategory: '',
      STATUS: 'A',
      UserID: 0,
    };

    this.apiSrvc.postmethod(this.comm.routeEndpoint + 'GetLenderTab', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          this.spinner.hide();
          this.LendersData = res.response;
          this.filteredData = this.LendersData;
        } else {
          this.toast.show(res.status, 'danger', 'Error');
          this.spinner.hide();
          this.NoData = true;
        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        this.spinner.hide();
        this.NoData = true;
      }
    );
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

  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit(): void {

    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Lender Mapping') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Lender Mapping') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Lender Mapping') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Lender Mapping') return;
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
  onSearch() {
    const value = this.searchText.toLowerCase();

    this.filteredData = this.LendersData.filter((item: any) =>
      (item.DMS_Lender_Text || '').toLowerCase().includes(value) ||
      (item.LenderName || '').toLowerCase().includes(value) ||
      (item.LenderType || '').toLowerCase().includes(value) ||
      (item.Category || '').toLowerCase().includes(value)
    );
  }
  ExcelStoreNames: any = [];
  generatePDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save('Lender Mapping.pdf'); // ✅ only here
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
      const pdfFile = this.blobToFile(pdfBlob, 'Lender Mapping.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Lender Mapping');
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
  private createPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

    /* ================= TITLE ================= */
    doc.setFontSize(14);
    doc.text('LENDER MAPPING', 14, 12);

    let startY = 20;

    /* ================= HEADERS (same as Excel) ================= */
    const head = [[
      'LENDER NAME IN DMS',
      'LENDER NAME',
      'LENDER TYPE',
      'LENDER CATEGORY'
    ]];

    /* ================= BODY (same data as Excel) ================= */
    const body: any[] = [];

    this.LendersData.forEach((d: any) => {
      body.push([
        d.DMS_Lender_Text || '-',
        d.LenderName || '-',
        d.LenderType || '-',
        d.Category || '-'
      ]);
    });

    /* ================= TABLE ================= */
    autoTable(doc, {
      startY: startY,
      head: head,
      body: body,
      theme: 'grid',

      styles: {
        fontSize: 10,
        cellPadding: 3,
        valign: 'middle',
        textColor: [20, 20, 20], // 👈 slightly richer black
        lineWidth: 0.1
      },

      /* ✅ COLUMN WIDTH (same as Excel) */
      columnStyles: {
        0: { cellWidth: 60 }, // DMS
        1: { cellWidth: 100 }, // Name
        2: { cellWidth: 50 }, // Type
        3: { cellWidth: 50 }  // Category
      },

      didParseCell: (data: any) => {

        /* ================= HEADER STYLE ================= */
        if (data.section === 'head') {
          data.cell.styles.fillColor = [42, 145, 240]; // Excel blue
          data.cell.styles.textColor = [255, 255, 255];
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.halign = 'center';
          return;
        }

        const rowIndex = data.row.index;

        /* ================= ALIGN ================= */
        data.cell.styles.halign = 'left';

        /* ================= ZEBRA ROW ================= */
        if (rowIndex % 2 === 1) {
          data.cell.styles.fillColor = [245, 247, 250]; // same as Excel
        }
      }
    });

    return doc;
  }
  exportToExcel() {

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet("Lender Mapping");

    const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');

    let rowIndex = 1;

    /* ================= TITLE ================= */

    const titleRow = worksheet.addRow(['LENDER MAPPING']);

    titleRow.font = {
      name: 'Calibri',
      size: 14,
      bold: true
    };

    titleRow.alignment = {
      horizontal: 'center',
      vertical: 'middle'
    };

    worksheet.mergeCells(`A${rowIndex}:D${rowIndex}`);

    // Borders for title
    for (let c = 1; c <= 4; c++) {
      const cell = titleRow.getCell(c);

      if (!cell.value) cell.value = '';

      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" }
      };
    }

    rowIndex++;

    /* ================= GAP ROW ================= */

    worksheet.addRow([]);   // ✅ empty row
    rowIndex++;

    /* ================= HEADER ================= */

    const headers = [
      'LENDER NAME IN DMS',
      'LENDER NAME',
      'LENDER TYPE',
      'LENDER CATEGORY'
    ];

    const headerRow = worksheet.addRow(headers);

    headerRow.eachCell((cell: any) => {

      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "2A91F0" }
      };

      cell.font = {
        name: "Calibri",
        size: 11,
        bold: true,
        color: { argb: "FFFFFF" }
      };

      cell.alignment = { horizontal: "center", vertical: "middle" };

      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" }
      };
    });

    /* ================= FREEZE HEADER ================= */

    worksheet.views = [{
      state: 'frozen',
      ySplit: 1
    }];

    /* ================= COLUMNS ================= */

    worksheet.columns = [
      { key: "dms", width: 30 },
      { key: "name", width: 50 },
      { key: "type", width: 25 },
      { key: "category", width: 25 }
    ];

    /* ================= DATA ================= */

    let dataRowIndex = rowIndex + 1;

    this.LendersData.forEach((d: any) => {

      const row = worksheet.addRow({
        dms: d.DMS_Lender_Text || '-',
        name: d.LenderName || '-',
        type: d.LenderType || '-',
        category: d.Category || '-'
      });

      row.eachCell((cell: any) => {

        cell.font = { name: "Calibri", size: 11 };

        cell.alignment = {
          horizontal: "left",
          vertical: "middle",
          wrapText: true
        };

        cell.border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" }
        };
      });

      // Zebra rows
      if (dataRowIndex % 2 === 0) {
        row.eachCell((cell: any) => {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "F5F7FA" }
          };
        });
      }

      dataRowIndex++;
    });

    /* ================= EXPORT ================= */

    workbook.xlsx.writeBuffer().then((buffer) => {
      FileSaver.saveAs(
        new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }),
        `Lender_Mapping_${DATE_EXTENSION}.xlsx`
      );
    });
  }
}


