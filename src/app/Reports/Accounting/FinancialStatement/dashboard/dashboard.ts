import { Component, ElementRef, HostListener, Renderer2 } from '@angular/core';
import {
  BsDatepickerConfig,
  BsDatepickerModule,
} from 'ngx-bootstrap/datepicker';
import { NgxSpinnerService } from 'ngx-spinner';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { CurrencyPipe, DatePipe, DecimalPipe } from '@angular/common';
import { common } from '../../../../common';
import { Title } from '@angular/platform-browser';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { Api } from '../../../../Core/Providers/Api/api';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { environment } from '../../../../../environments/environment';
import { Subscription } from 'rxjs';
import { Stores } from '../../../../CommonFilters/stores/stores';
import numeral from 'numeral';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import * as ExcelJS from 'exceljs';
import * as FileSaver from 'file-saver';

import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  providers: [CurrencyPipe, DatePipe, DecimalPipe],
  templateUrl: './dashboard.html',
  styleUrls: ['./dashboard.scss']
})
export class Dashboard {
  dynamicTitle = 'Financial Statement';
  FSData: any[] = [];
  Month: string | null = null;
  // storeIds: any = '';
  storeName: any;
  presentDate: Date | null = null;
  lastFourMonthsArray: string[] = [];
  FsNegativeData: any[] = [];
  FinancialStatementExcel: any[] = [];
  NodataFound = false;
  NoData = false;
  selectedDate: Date = new Date();
  currentMonth!: Date;
  // groupId: any = [1];
  storeIds: any = [4]
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  groups: any = 1;

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': this.storeIds, 'type': 'S', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  activePopover: number = -1;
  ngChanges: any;

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  constructor(
    public apiSrvc: Api,
    public shared: Sharedservice,
    private spinner: NgxSpinnerService,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private renderer: Renderer2,
    private el: ElementRef,
    private title: Title,
    private datepipe: DatePipe,
    private currencyPipe: CurrencyPipe,
    private decimalPipe: DecimalPipe,
    private toast: ToastService,
    private comm: common
  ) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')[0]
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')[0]
    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }

    this.shared.setTitle(this.shared.common.titleName + '-Financial Statement');
    // if (localStorage.getItem('Fav') !== 'Y') {
    const data = { title: this.dynamicTitle, stores: '2' };
    this.shared.api.SetHeaderData({ obj: data });
    const lastMonth = new Date();
    let today = new Date();
    if (today.getDate() < 7) {
      this.selectedDate = new Date(lastMonth.setMonth(lastMonth.getMonth() - 1));
    } else {
      this.selectedDate = new Date(lastMonth.setMonth(lastMonth.getMonth()));
    }
    this.currentMonth = this.selectedDate;
    this.onMonthChange(this.currentMonth);
    // }
  }



  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month',
    maxDate: new Date()
  };

  applyDateChange() {
    if (!this.storeIds || this.storeIds.length === 0) {

      this.toast.show(
        'Please Select Atleast One Store',
        'warning',
        'Warning'
      );
      return;
    }
    else {
      this.currentMonth = this.selectedDate;
      this.onMonthChange(this.currentMonth);
    }
  }

  onMonthChange(monthDate: Date) {
    this.spinner.show();
    this.lastFourMonthsArray = [];
    if (!monthDate) return;
    this.selectedDate = monthDate;
    this.lastFourMonthsArray = [];
    for (let i = 1; i <= 3; i++) {
      const lastMonthDate = new Date(monthDate.getFullYear(), monthDate.getMonth() - i, 1);
      const formattedDate = lastMonthDate.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long'
      });
      this.lastFourMonthsArray.push(formattedDate);
    }
    const obj = {
      stores: this.storeIds.toString(),
      date: '10-'+this.shared.datePipe.transform(monthDate,'MMM-yyyy'),
    };
    this.apiSrvc.postmethod(this.comm.routeEndpoint + 'GetFbFinancialStatementReport', obj).subscribe({
      next: (res) => {
        this.spinner.hide();

        if (res.status !== 200) {

          this.toast.show('Invalid Details', 'danger', 'Error');
          return;
        }

        this.FSData = res.response || [];
        this.FinancialStatementExcel = this.FSData;
        this.FsNegativeData = this.FSData;
        this.FSData = this.FSData.map((v: any) => ({
          ...v,
          Fs_Titles: 'SubTitle',
          ValueType: '$',
        }));
        this.FSData.forEach((val: any) => {
          const sentence = val.DisplayName || '';
          if (sentence.match(/Units?|Retail/i)) val.ValueType = '#';
          else if (sentence.match(/Percent|Absorption/i)) val.ValueType = '%';
          else val.ValueType = '$';
        });
        this.FSData.forEach((val: any) => {
          const label = val.LabelName;
          if (
            [
              'TotalSellingGross-New', 'SellingGross%VariableGross-New', 'SellingGrossPUS-New',
              'TotalSellingGross-Used', 'SellingGross%VariableGross-Used', 'SellingGrossPUS-Used',
              'Sales-Fixed', 'Gross-Fixed', 'Gross%Sales-Fixed', 'Gross%TotalDeptExpIncludingFixedExp-Fixed',
              'SellingGross-AllDepts', 'SellingGross%Gross-AllDepts', 'TotalFixedExpenses',
              'FixedExpenses%Gross-AllDepts', 'OpertingProfit-AllDepts', 'OpertingProfit%Gross-AllDepts',
              'DealerSalary', 'DealerSalary%Gross-AllDepts', 'OtherAdditions&Reductions',
              'OtherAddDeducts%Gross-AllDepts', 'NetIncome-BeforeTaxes',
              'NetIncomeBeforeTaxes%Gross-AllDepts', 'Payroll-AllDepts', 'TotalDealershipSales',
              'Payroll%Gross-AllDepts', 'SellingGrossPercent', 'SellingGross-Service',
              'SellingGross%DeptGross-Service', 'SellingGross-Parts', 'SellingGross%DeptGross-Parts',
            ].includes(label)
          ) {
            val.Fs_Titles = 'Title';
          } else if (
            ['TotalGrossProfit', 'SellingGross', 'Payroll'].includes(label)
          ) {
            val.Fs_Titles = 'TitleCenter';
          }
        });
        if (this.FSData.length > 0) {
          this.FSData.forEach((x: any) => {
            (x.SubData = []), (x.data2sign = '+');
          });
          this.NoData = false;
        } else {
          this.NoData = true;
        }
        console.log('FS Data', this.FSData);
        this.NodataFound = true;
        this.NoData = this.FSData.length === 0;
      },
      error: (err) => {
        console.error(err);
        this.spinner.hide();
      },
    });
  }

  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit() {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Financial Statement') {
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
      if (!obj || obj.title !== 'Financial Statement') return;
      if (obj.state) {
        this.exportToExcelFinancialStatement();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Financial Statement') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Financial Statement') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Financial Statement') return;
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
      'type': 'S', 'others': 'N'
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
  expandorcollapse(ind: any, e: any, ref: any, Item: any, Department: any) {
    let id = (e.target as Element).id;
    console.log(ref, Item);
    if (id == 'D_' + ind) {
      console.log(id);
      if (ref == '-') {
        this.FSData[ind].data2sign = '+';
        this.FSData[ind].SubData = [];
      }
      if (ref == '+') {
        this.FSData[ind].data2sign = '-';
        this.GetFinancialStatementLayerOne(Item, ind, Department);
      }
    }
    if (id == 'DN_' + ind) {
      if (ref == '-') {
        this.FinancialStatementLayerOne[ind].dataLayer2sign = '+';
        this.FinancialStatementLayerOne[ind].SubLayerData = [];
      }
      if (ref == '+') {
        this.FinancialStatementLayerOne[ind].dataLayer2sign = '-';
        // this.GetFixedExpenseLayerTwo(parentData, Item, ind);
        this.FSData[ind].data2sign = '-';
      }
    }
  }
  FinancialStatementLayerOne: any;
  FinancialStatementLayerOneKeys: any;
  GetFinancialStatementLayerOne(Object: any, index: number, Department: any): void {
    this.spinner.show();
    this.FinancialStatementLayerOne = [];
    this.FinancialStatementLayerOneKeys = [];
    const DateToday = this.datepipe.transform(
      new Date(this.selectedDate),
      'dd-MMM-yyyy'
    );
    const obj = {
      "Store": this.storeIds,
      "Department": Object.Department,
      "DATE": '10-'+this.shared.datePipe.transform(new Date(this.selectedDate),'MMM-yyyy'),
      "Headname": Object.DisplayName
    };

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetFinStatementExpenseReportDetail', obj)
      .subscribe(
        (x: any) => {
          if (x.status === 200) {
            this.FinancialStatementLayerOne = x.response;
            this.FinancialStatementLayerOne = this.FinancialStatementLayerOne.map((v: any) => ({
              ...v,
              Fs_Titles: 'SubTitle',
              ValueType: '$'
            }));
            this.FinancialStatementLayerOne.forEach((val: any) => {
              const sentence: string = val['SUBTYPE DETAIL'];
              const Units: string = 'Units';
              const Percentage: string = 'Percent';
              const index: number = sentence.indexOf(Units);
              const index2: number = sentence.indexOf(Percentage);
              if (index !== -1) {
                val.ValueType = '#'
              } else if (index2 !== -1) {
                val.ValueType = '%'
              } else {
                val.ValueType = '$'
              }
            })
            console.log('GetFinStatementExpenseDetail', this.FinancialStatementLayerOne);

            if (
              this.FinancialStatementLayerOne &&
              this.FinancialStatementLayerOne.length > 0
            ) {
              // this.FinancialStatementLayerOne.forEach((item: any) => {
              //   item.SubLayerData = [];
              //   item.dataLayer2sign = '+';
              // });
              this.FSData.forEach((val: any) => {
                val.SubData = [];
                val.data2sign = '+';
                this.FinancialStatementLayerOne.forEach((ele: any) => {
                  console.log(ele.Headname);
                  if (val.DisplayName.replace(/\s+/g, '') === ele.Headname && val.Department === Department) {
                    console.log(ele.Headname);
                    val.SubData.push(ele);
                    val.data2sign = '-';
                  }
                });
              });
              this.NoData = false;
            } else {
              this.NoData = false;
            }
            this.spinner.hide();
          }
        },
        (error: any) => {
          console.error(error);
          this.spinner.hide();
        }
      );
  }

  isSpecialRow(name: string): boolean {
    return [
      'Total Selling Gross (New)',
      'Total Selling Gross (Used)',
      'Net Income Before Taxes',
      'Total Operating Department Profit'
    ].includes(name);
  }

  getStyle(value: any, valueType: string) {
    if (value === null || value === '' || value === 0 || value === '0') {
      return { 'justify-content': 'center' };
    }
    switch (valueType) {
      case '$':
        return { 'justify-content': 'space-between' };
      case '#':
        return { 'justify-content': 'flex-end' };
      case '%':
        return { 'justify-content': 'center' };
      default:
        return { 'justify-content': 'initial' };
    }
  }
  formatValue(value: any, valueType: string): string {
    if (value === null || value === undefined || value === 0 || value === '0') {
      return '-';
    } else if (valueType === '#') {
      return typeof value === 'number' ? numeral(value).format('0,0') : value;
    } else if (valueType === '$') {
      return typeof value === 'number' ? numeral(value).format('0,0') : value;
    } else if (valueType === '%') {
      return typeof value === 'number' ? value.toLocaleString() + '%' : value;
    } else {
      return value;
    }
  }
  index: string = '';
  selectRow(i: number) {
    this.index = i.toString();
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  formatExcelValue(value: any, valueType: string) {
    if (value === null || value === '' || value === undefined || value === 0 || value === '0') {
      return '-';
    }
    switch (valueType) {
      case '#':
        return value;
      case '$':
        return value;
      case '%':
        return value / 100;
      default:
        return value;
    }
  }
  generatePDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save('Financial Statement.pdf'); // ✅ only here
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
      const pdfFile = this.blobToFile(pdfBlob, 'Financial Statement.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Financial Statement');
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
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });

    /* ================= TITLE ================= */
    doc.setFontSize(14);
    doc.text('Financial Statement', 14, 12);

    /* ================= FILTER INFO ================= */
    doc.setFontSize(10);

    // const filters = this.getReportFilters().filters;
    // let startY = 18;

    // filters.forEach((f: any) => {
    //   doc.text(`${f.label}: ${f.value}`, 14, startY);
    //   startY += 5;
    // });

    // startY += 5;

    /* ================= HEADER ================= */
    const filtersObj = this.getReportFilters();

    const monthValue =
      filtersObj.filters.find((f: any) => f.label === 'Month')?.value
      || 'Selected Date';

    const storeValue =
      filtersObj.filters.find((f: any) => f.label === 'Store')?.value
      || 'All Stores';

    const headers = [
      [
        monthValue,   // 👈 instead of Selected Date
        storeValue,   // 👈 dynamic store name
        '',
        '',
        '',
        '',
        ''
      ],
      [
        'Department',
        'Actual',
        'Budget',
        'Variance',
        this.lastFourMonthsArray[0],
        this.lastFourMonthsArray[1],
        this.lastFourMonthsArray[2]
      ]
    ];
    /* ================= BODY ================= */
    const body: any[] = [];

    this.FSData.forEach((data: any) => {
      const isHeader = data.IsHeader === 'Y';

      body.push({
        rowData: data,
        values: [
          data.DisplayName,
          isHeader ? '' : data.Actual,
          isHeader ? '' : data.Budget,
          isHeader ? '' : data.Variance,
          isHeader ? '' : data.LM1,
          isHeader ? '' : data.LM2,
          isHeader ? '' : data.LM3
        ]
      });

      /* ===== SUB ROWS ===== */
      if (Array.isArray(data.SubData)) {
        data.SubData.forEach((sub: any) => {
          body.push({
            rowData: sub,
            isSub: true,
            values: [
              '   ' + (sub['SUBTYPE DETAIL'] || ''),
              sub.Actual,
              sub.Budget,
              sub.Variance,
              sub.LM1,
              sub.LM2,
              sub.LM3
            ]
          });
        });
      }
    });

    /* ================= TABLE ================= */
    autoTable(doc, {
      startY: 18,
      head: headers,
      body: body.map(r => r.values),
      theme: 'grid',

      styles: {
        fontSize: 8,
        cellPadding: 2,
        halign: 'right',
        lineWidth: 0.1,
        lineColor: [200, 200, 200],
        textColor: [20, 20, 20], // 👈 slightly richer black
        valign: 'middle'
      },

      headStyles: {
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        lineWidth: 0.1,
        lineColor: [200, 200, 200]
      },

      didParseCell: (data: any) => {

        const rowObj = body[data.row.index];
        const rowData = rowObj?.rowData;

        /* ================= HEADER ================= */
        if (data.section === 'head') {

          if (data.row.index === 0) {
            data.cell.styles.fillColor = [5, 84, 239];
            data.cell.styles.textColor = [255, 255, 255];
            data.cell.styles.halign = 'center';
            data.cell.styles.fontStyle = 'bold';

            if (data.column.index === 0) data.cell.colSpan = 1;
            if (data.column.index === 1) data.cell.colSpan = 6;
            if (data.column.index > 1) data.cell.text = '';
          }

          if (data.row.index === 1) {
            data.cell.styles.fillColor = [69, 132, 255];
            data.cell.styles.textColor = [255, 255, 255];
            data.cell.styles.fontStyle = 'bold';
            data.cell.styles.halign = data.column.index === 0 ? 'left' : 'center';
          }

          return;
        }

        if (!rowData) return;

        const isHeader = rowData.IsHeader === 'Y';

        if (data.column.index === 0) {
          data.cell.styles.halign = 'left';

          if (isHeader) {
            data.cell.styles.fontStyle = 'bold';
          }

          return;
        }

        if (!data.cell.raw || data.cell.raw === '-') {
          data.cell.text = isHeader ? [''] : ['-'];
          data.cell.styles.halign = 'center';
          return;
        }

        if (data.column.index === 0) return;

        const valueType = rowData.ValueType || '$';

        let rawValue = data.cell.raw?.toString().replace(/[$,% ,]/g, '');
        let val = parseFloat(rawValue);

        if (!isNaN(val)) {

          let formattedText = '';

          if (valueType === '$') {
            const rounded = Math.round(val);
            formattedText = `$ ${rounded.toLocaleString()}`;

            const name = rowData.DisplayName?.toLowerCase() || '';

            const isTargetRow =
              name.includes('total selling gross (new)') ||
              name.includes('total selling gross (used)') ||
              name.includes('total operating department profit');

            if (rounded < 0 && isTargetRow) {
              data.cell.styles.textColor = [220, 53, 69];
            }
          }

          else if (valueType === '#') {
            formattedText = Math.round(val).toLocaleString();
          }

          else if (valueType === '%') {
            formattedText = `${val.toFixed(1)}%`;
            data.cell.styles.halign = 'center';
          }

          data.cell.text = [formattedText];
        }

        data.cell.styles.halign =
          valueType === '%'
            ? 'center'
            : 'right';

        if (isHeader) {
          Object.values(data.row.cells).forEach((cell: any) => {
            cell.styles.fillColor = [217, 231, 255];
            cell.styles.fontStyle = 'bold';
          });
        }

        if (rowData.DisplayName?.toLowerCase().includes('total')) {
          Object.values(data.row.cells).forEach((cell: any) => {
            cell.styles.fontStyle = 'bold';
          });
        }

        if (
          data.row.index % 2 === 0 &&
          !isHeader &&
          !rowData.DisplayName?.toLowerCase().includes('total')
        ) {
          Object.values(data.row.cells).forEach((cell: any) => {
            cell.styles.fillColor = [245, 247, 250];
          });
        }

      }
    });

    return doc;
  }
  getReportFilters(): { title: string; filters: any[] } {
    return {
      title: 'Financial Statement',
      filters: [
        {
          label: 'Store',
          value: this.storename || 'All Stores'
        },
        {
          label: 'Group',
          value: this.groupName || 'All Groups'
        },
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
  exportToExcelFinancialStatement() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Financial Statement');

    /* ================= 1. FILTERS AT TOP ================= */
    const filterRowCount = this.addExcelFiltersSection(worksheet);

    /* ================= 2. HEADER ================= */
    const headerStartRow = filterRowCount + 1;

    // Header Row 1
    worksheet.addRow([
      this.datepipe.transform(this.currentMonth, 'MMMM yyyy'),
      this.storename || 'All Stores'
    ]);
    worksheet.mergeCells(`B${headerStartRow}:G${headerStartRow}`);

    // Header Row 2
    worksheet.addRow([
      'Department',
      'Actual',
      'Budget',
      'Variance',
      this.lastFourMonthsArray[0],
      this.lastFourMonthsArray[1],
      this.lastFourMonthsArray[2]
    ]);

    /* ================= HEADER STYLE ================= */
    [headerStartRow, headerStartRow + 1].forEach(r => {
      worksheet.getRow(r).eachCell(cell => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: r === headerStartRow ? '0554EF' : '4584FF' }
        };
        cell.font = {
          bold: true,
          color: { argb: 'FFFFFF' },
          name: 'Calibri'
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    /* ================= FORMAT FUNCTION ================= */
    const formatRow = (
      row: any,
      valueTypes: any[] = [],
      isSpecial = false,
      isTitleRow = false,
      isHeaderRow = false
    ) => {
      row.eachCell((cell: any, colNumber: number) => {

        if (colNumber === 1) {
          cell.alignment = { horizontal: 'left', vertical: 'middle', indent: isTitleRow ? 2 : 3 };
          cell.font = { name: 'Calibri', bold: isTitleRow || isSpecial };
          return;
        }

        const valueType = valueTypes[colNumber - 2];

        if (!cell.value || cell.value === '-') {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          return;
        }

        let num = Number(cell.value);

        if (!isNaN(num)) {
          if (valueType === '$') cell.numFmt = '"$" * #,##0; "$" * -#,##0';
          if (valueType === '#') cell.numFmt = '#,##0';
          if (valueType === '%') {
            cell.numFmt = '0.0%';

          }

          if (isSpecial && colNumber === 4 && num < 0) {
            cell.font = { color: { argb: 'FF0000' }, bold: true };
          }
        }

        cell.alignment = valueType === '%'
          ? { horizontal: 'center', vertical: 'middle' }
          : { horizontal: 'right', vertical: 'middle' };
      });
    };

    /* ================= DATA ================= */
    this.FSData.forEach((data: any, i: number) => {

      const isHeaderRow = data.IsHeader === 'Y';
      const isTitleRow = data.IsHeader === 'T';
      const isSpecial = this.isSpecialRow(data.DisplayName);

      const row = worksheet.addRow([
        data.DisplayName,
        isHeaderRow ? '' : this.formatExcelValue(data.Actual, data.ValueType),
        isHeaderRow ? '' : this.formatExcelValue(data.Budget, data.ValueType),
        isHeaderRow ? '' : this.formatExcelValue(data.Variance, data.ValueType),
        isHeaderRow ? '' : this.formatExcelValue(data.LM1, data.ValueType),
        isHeaderRow ? '' : this.formatExcelValue(data.LM2, data.ValueType),
        isHeaderRow ? '' : this.formatExcelValue(data.LM3, data.ValueType)
      ]);

      if (isHeaderRow) {
        row.eachCell(cell => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9E7FF' } };
          cell.font = { bold: true };
        });
      }

      if (isTitleRow) {
        row.eachCell(cell => cell.font = { bold: true });
      }

      if (!isHeaderRow) {
        formatRow(row, Array(6).fill(data.ValueType), isSpecial, isTitleRow, false);
      }

      /* ===== SUB ROWS ===== */
      if (Array.isArray(data.SubData)) {
        data.SubData.forEach((sub: any) => {
          const row2 = worksheet.addRow([
            sub['SUBTYPE DETAIL'] || '',
            this.formatExcelValue(sub.Actual, sub.ValueType),
            this.formatExcelValue(sub.Budget, sub.ValueType),
            this.formatExcelValue(sub.Variance, sub.ValueType),
            this.formatExcelValue(sub.LM1, sub.ValueType),
            this.formatExcelValue(sub.LM2, sub.ValueType),
            this.formatExcelValue(sub.LM3, sub.ValueType)
          ]);
          row2.outlineLevel = 1;
          formatRow(row2, Array(6).fill(sub.ValueType));
        });
      }
    });

    /* ================= FINAL SETTINGS ================= */
    worksheet.views = [{ state: 'frozen', xSplit: 1, ySplit: headerStartRow + 1 }];

    worksheet.columns = [
      { width: 65 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 18 }
    ];

    workbook.xlsx.writeBuffer().then(data => {
      FileSaver.saveAs(new Blob([data]), 'FinancialStatement.xlsx');
    });
  }

}
