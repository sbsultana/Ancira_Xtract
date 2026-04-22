import { ChangeDetectorRef, Component, ElementRef, HostListener, Inject, PLATFORM_ID, ViewChild } from '@angular/core';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { ToastContainer } from '../../../../Layout/toast-container/toast-container';
import { NgxSpinnerModule, NgxSpinnerService } from 'ngx-spinner';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { Api } from '../../../../Core/Providers/Api/api';
import { HttpClient } from '@angular/common/http';
import { NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { common } from '../../../../common';
import { FormBuilder } from '@angular/forms';
import { Router } from '@angular/router';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Workbook } from 'exceljs';
import * as FileSaver from 'file-saver';
import { Subscription } from 'rxjs';
import * as ExcelJS from 'exceljs';
import autoTable from 'jspdf-autotable';
import jsPDF from 'jspdf';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, ToastContainer, NgxSpinnerModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  @ViewChild('chartContainer3') chartContainer3!: ElementRef<HTMLDivElement>;
  @ViewChild('chartContainer2') chartContainer2!: ElementRef<HTMLDivElement>;
  storeIds: any = [71];
  stores: any = [];
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
  ngChanges: any;


  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  tableData: { variable: string; months: number[] }[] = [];
  // backupData: { variable: string; months: number[] }[] = [];
  constructor(
    public shared: Sharedservice,
    @Inject(PLATFORM_ID) private platformId: Object,
    private toast: ToastService,
    private spinner: NgxSpinnerService,
    public apiSrvc: Api,
    private http: HttpClient,
    public ngbModal: NgbModal,
    private comm: common,
    private fb: FormBuilder,
    private router: Router,
    private cd: ChangeDetectorRef,
  ) {
    this.shared.setTitle(this.shared.common.titleName + '- Budget Forecast');
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
    let obj = {
      title: 'Budget Forecast',
      stores: '71'
    }
    this.apiSrvc.SetHeaderData({ obj });
    this.bsConfig = {
      dateInputFormat: 'YYYY',
      minMode: 'year',
      maxDate: this.maxDate
    };
    this.getBudgetForecastData(this.selectedFilter, this.selectedYear);
    this.setAvgMonths();
  }

  //////////////////////////////////////////Store Filter Code Start//////////////////////////////////////////////////
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit() {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Budget Forecast') {
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
      if (!obj || obj.title !== 'Budget Forecast') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Budget Forecast') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Budget Forecast') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Budget Forecast') return;
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
  }
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
  }

  //////////////////////////////////////////Store Filter End////////////////////////////////////////////////////

  PDCForecastData: any[] = [];
  NoPDCForecastdata = false;
  selectedstorevalues: number = 60;
  variableType: string = '';
  isEditMode = false;
  backupData: any;

  filters: string[] = ['sales', 'expenses'];
  selectedFilter: string = 'sales';

  activePopover: number | null = null;
  selectedYear: any = new Date().getFullYear();

  bsConfig!: Partial<BsDatepickerConfig>;
  maxDate: Date = new Date();
  bsValue: Date = new Date();


  selectFilter(f: string) {
    this.selectedFilter = f;
    this.activePopover = null;
  }
  onYearSelected(event: Date) {
    if (event) {
      this.selectedYear = event.getFullYear();
    }
  }
  applyFilterAndYear() {
    if (!this.storeIds?.length) {
      this.toast.show(
        'Please Select Atleast One Store',
        'warning',
        'Warning'
      );
      return;
    }
    console.log('Filters', this.storeIds.toString(), this.selectedFilter, this.selectedYear);
    this.PDCForecastData = [];
    this.getBudgetForecastData(this.selectedFilter, this.selectedYear);
  }
  cancelEdit() {
    this.isEditMode = false;

    // Restore from backup if available
    if (this.backupData && this.backupData.length) {
      this.PDCForecastData = JSON.parse(JSON.stringify(this.backupData));
    } else {
      // If no backup, just clear monthly fields
      const monthKeys = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      this.PDCForecastData.forEach(row => {
        monthKeys.forEach(key => {
          row[key] = '';
        });
      });
    }
  }



  // getRowTotal(row: { variable: string; months: number[] }): number {
  //   return row.months.reduce((sum, val) => sum + val, 0);
  // }

  // getGrandTotal(): number {
  //   return this.tableData.reduce((acc, row) => acc + this.getRowTotal(row), 0);
  // }
  // getColumnTotal(index: number): number {
  //   return this.tableData.reduce((sum, row) => sum + (row.months[index] || 0), 0);
  // }

  // Add these helper functions in your component (below getBudgetForecastData)

  getRowTotal(row: any): number {
    const months = [
      row.Jan, row.Feb, row.Mar, row.Apr, row.May, row.Jun,
      row.Jul, row.Aug, row.Sep, row.Oct, row.Nov, row.Dec
    ].map((m: any) => Number(m) || 0);
    return months.reduce((a: number, b: number) => a + b, 0);
  }

  getColumnTotal(index: number): number {
    const monthKeys = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const key = monthKeys[index];
    return this.PDCForecastData.reduce((sum, row) => sum + (Number(row[key]) || 0), 0);
  }

  getGrandTotal(): number {
    return this.PDCForecastData.reduce((sum, row) => sum + (row.sum || 0), 0);
  }

  selectedDt: number = 2025;

  getBudgetForecastData(type: string, year: number): void {
    this.variableType = type;
    const obj = {
      IV_StoreID: this.storeIds.toString(),
      IV_V_Type: type === 'sales' ? 'store' : type,
      IV_Year: year?.toString() || new Date().getFullYear().toString()
    };

    this.spinner.show();
    this.apiSrvc.postmethod('IncomeTargetValues/GetIncomeTargetValuesExpenseV2', obj)
      .subscribe({
        next: (res: any) => {
          this.spinner.hide();
          if (res.status === 200 && Array.isArray(res.response)) {
            this.PDCForecastData = [...res.response].sort((a: any, b: any) => {
              const catA = a.CategorySequence ?? 0;
              const catB = b.CategorySequence ?? 0;
              const seqA = a.SequenceNo ?? 0;
              const seqB = b.SequenceNo ?? 0;
              return catA === catB ? seqA - seqB : catA - catB;
            });
            this.cd.detectChanges();
            const userId = JSON.parse(localStorage.getItem('userInfo')!)?.user_Info?.userid;

            this.PDCForecastData.forEach(ele => {
              if (!ele.IV_Variable_ID || ele.IV_Variable_ID === 'null') {
                ele.IV_Variable_ID = ele.ITV_Id;
                ele.IV_StoreID = this.selectedstorevalues;
                ele.IV_Seq = ele.SequenceNo ?? 0;
                ele.IV_Formula_Status = ele.ITV_Formula_Status;
                ele.IV_User_ID = userId;
                ele.IV_Year = year.toString();
              }
              const monthKeys = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
              monthKeys.forEach(key => {
                ele[key] = ele[key] != null && ele[key] !== '' ? Number(ele[key]) : null;
              });
              const monthValues = monthKeys.map(key => Number(ele[key]) || null);
              ele.sum = monthValues.reduce((a: any, b: any) => a + b, null);
            });
            this.setAvgMonths();
            console.log('Forecast Data', this.PDCForecastData);

          } else {
            this.PDCForecastData = [];
          }
          this.NoPDCForecastdata = this.PDCForecastData.length === 0;
        },
        error: err => {
          console.error('API Error:', err);
          this.spinner.hide();
          this.NoPDCForecastdata = true;
        }
      });
  }

  EditSave() {
    if (!this.isEditMode) {
      this.backupData = JSON.parse(JSON.stringify(this.PDCForecastData));
      this.isEditMode = true;
    } else {
      const storeId = this.storeIds.toString();
      const userId = JSON.parse(localStorage.getItem('userInfo')!)?.user_Info?.userid;

      const payload = this.PDCForecastData.map(ele => {
        const newRow = { ...ele };
        if (!newRow.IV_Variable_ID || newRow.IV_Variable_ID === 'null') {
          newRow.IV_Variable_ID = newRow.ITV_Id;
          newRow.IV_StoreID = storeId;
          newRow.IV_Seq = newRow.SequenceNo ?? 0;
          newRow.IV_Formula_Status = newRow.ITV_Formula_Status;
          newRow.IV_User_ID = userId;
          newRow.IV_Year = this.selectedYear.toString();
        } else {
          newRow.IV_StoreID = storeId;
        }
        const monthKeys = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        monthKeys.forEach(key => {
          newRow[key] = newRow[key] != null && newRow[key] !== '' ? newRow[key] : null;
        });
        const monthValues = monthKeys.map(key => Number(newRow[key]) || null);
        newRow.sum = monthValues.reduce((a: any, b: any) => a + b, null);

        return newRow;
      });
      console.log('Pay Load Obj', payload);
      this.spinner.show();
      this.apiSrvc.postmethod('IncomeTargetValues/AddIncomeTargetValuesV2', payload)
        .subscribe({
          next: res => {
            this.spinner.hide();
            if (res.status === 200) {
              this.toast.show('Data saved successfully!', 'success', 'Success');
              this.onYearSelected(new Date(this.selectedYear, 0, 1));
              this.isEditMode = false;
            } else {
              this.toast.show('Invalid inputs!', 'danger', 'Error');
            }
          },
          error: err => {
            console.error('API error: ', err);
            this.toast.show('Something went wrong!', 'danger', 'Error');
            this.spinner.hide();
          }
        });
    }
  }
  monthsShort = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  onPaste(event: ClipboardEvent, rowIndex: number, startColIndex: number): void {
    event.preventDefault();

    const clipboardData = event.clipboardData || (window as any).clipboardData;
    const pastedText = clipboardData.getData('text');
    const rows = pastedText.split('\n').filter((r: any) => r.trim() !== '');
    const values = rows[0].split('\t');
    if (!this.PDCForecastData[rowIndex]) return;
    for (let i = 0; i < values.length; i++) {
      const colIndex = startColIndex + i;
      if (colIndex >= this.monthsShort.length) break;
      const monthKey = this.monthsShort[colIndex];
      this.PDCForecastData[rowIndex][monthKey] = values[i].trim();

      this.onKeydownEvent(
        new InputEvent('input'),
        this.PDCForecastData[rowIndex],
        rowIndex,
        monthKey,
        colIndex
      );
    }
  }


  avgMonths: string[] = [];
  setAvgMonths() {

    if (!this.PDCForecastData?.length) {
      this.avgMonths = [];
      return;
    }

    const today = new Date();
    const currentYear = today.getFullYear();
    const selectedYear = Number(this.PDCForecastData[0]?.IV_Year);

    if (selectedYear < currentYear) {
      this.avgMonths = [];
      return;
    }

    this.avgMonths = this.monthsShort.filter(m =>
      this.PDCForecastData.some((x: any) =>
        x['Avg_' + m] !== null &&
        x['Avg_' + m] !== undefined
      )
    );

    console.log("Avg Months:", this.avgMonths);
  }

  private calcTimer: any;

  onKeydownEvent(evnt: any, ele: any, ind1: any, mnth: any, ind2: any) {
    if (evnt.type === 'input') {
      if (
        ele.ITV_Variable_Name?.toLowerCase().includes('advertising rebates') ||
        ele.ITV_Variable_Name?.toLowerCase().includes('floor plan assistance')
      ) {
        let value = Number(ele[mnth]);
        if (!isNaN(value) && value > 0) {
          ele[mnth] = -value;
        }
      }
      this.PDCForecastData.filter((vr: any) => {
        if (vr.ITV_Formula_Status === 'Y' && vr.ITV_Formula) {
          const result = vr.ITV_Formula.split(/(?<=\S) (\-|\+|\*) (?=\S)/).map((item: any) => item.trim());

          if (this.PDCForecastData.length > 0 && result[0] && result[0] !== undefined) {
            const variables: any[] = [];
            const operators: string[] = [];

            for (let i = 0; i < result.length; i++) {
              if (i % 2 === 0) {
                const idx = this.PDCForecastData.findIndex((x: any) => x.ITV_Variable_Name === result[i]);
                if (idx > -1) {
                  variables.push(this.PDCForecastData[idx][mnth]);
                } else {
                  return;
                }
              } else {
                operators.push(result[i]);
              }
            }

            let expression = '';
            for (let j = 0; j < variables.length; j++) {
              expression += variables[j];
              if (j < operators.length) {
                expression += ' ' + operators[j] + ' ';
              }
            }
            try {
              const value = new Function(`return (${expression});`)();
              vr[mnth] = this.formatToTwoDecimal(value);
            } catch (e) {
              console.error('Invalid formula evaluation:', expression, e);
            }
          }
        }
      });
    }
  }

  formatToTwoDecimal(value: any): number {
    const num = Number(value);
    return isNaN(num) ? 0 : Number(num.toFixed(2));
  }


  formatNumber(value: any): string {
    const num = Number(value);
    if (isNaN(num)) return '0.00';

    return new Intl.NumberFormat('en-US', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(num);
  }

  BacktoPDCDashboard() {
    this.router.navigate(['']);
  }
  isEmptyValue(val: any): boolean {
    return val === null || val === undefined || val === '' || val === 0 || val === '0';
  }
  exportToExcel(): void {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Budget Forecast - Input');

    const isEditMode = this.isEditMode;
    const isStore = this.variableType === 'store';
    const data = isEditMode ? this.PDCForecastData : this.PDCForecastData;

    const storeName = this.storename ?? 'Store Name';
    const variableTypeLabel = isStore ? 'Store Variables' : 'Expense Variables';
    const year = this.selectedYear ?? 'Year';
    // ================= TITLE =================
    const titleRow = worksheet.addRow(['Budget Forecast']);
    titleRow.font = { size: 14, bold: true };
    worksheet.mergeCells('A1:R1');
    worksheet.addRow([]);


    // 1. Add Filter Info Rows (top 3 rows)
    worksheet.addRow(['Store Name:', storeName]);
    worksheet.addRow(['Variable Type:', variableTypeLabel]);
    worksheet.addRow(['Year:', year]);

    // Style Filter Info Rows
    for (let i = 3; i <= 6; i++) {
      const row = worksheet.getRow(i);
      row.getCell(1).font = { bold: true };
      row.getCell(1).alignment = { indent: 1 };
      row.getCell(2).alignment = { indent: 1 };
    }

    // 2. Header Row (row 4)
    const headerRow: any[] = ['Department', 'Description'];

    this.monthsShort.forEach((mn: any) => {
      if (this.avgMonths.includes(mn) && isEditMode) {
        headerRow.push('3 Months Avg');
      }
      headerRow.push(mn);
    });

    if (!isEditMode) {
      headerRow.push('Total');
    }

    const header = worksheet.addRow(headerRow);

    // Styling
    header.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF337ab7' },
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    // 3. Data Rows
    data.forEach((row: any, index: number) => {

      const showDept =
        index === 0 ||
        (this.variableType === 'sales'
          ? row.Category !== data[index - 1]?.Category
          : row.ITV_CategoryName !== data[index - 1]?.ITV_CategoryName);

      const department =
        this.variableType === 'sales'
          ? (showDept ? row.Category : '')
          : (showDept ? row.ITV_CategoryName : '');

      const rowData: any[] = [
        department,
        row.ITV_Variable_Name || ''
      ];

      // Months
      this.monthsShort.forEach((mn: any) => {

        if (this.avgMonths.includes(mn) && isEditMode) {
          rowData.push(row['Avg_' + mn] ?? '');
        }

        rowData.push(row[mn] ?? '');
      });

      // Total
      if (!isEditMode) {
        rowData.push(row.sum ?? '');
      }

      const addedRow = worksheet.addRow(rowData);

      // 🔹 Formatting
      const isGross = row.ValueType === 'G';
      const numFmt = isGross ? '"$"#,##0' : '#,##0';

      rowData.forEach((val, colIndex) => {
        const cell = addedRow.getCell(colIndex + 1);

        if (typeof val === 'number') {
          cell.numFmt = numFmt;
          cell.alignment = { horizontal: 'right' };
        } else {
          cell.alignment = { horizontal: colIndex < 2 ? 'left' : 'right' };
        }
      });
    });

    // 4. Auto width with minimum size
    worksheet.columns.forEach((column: any) => {
      let maxLength = 10;
      column.eachCell({ includeEmpty: true }, (cell: any) => {
        const val = cell.value ? cell.value.toString() : '';
        maxLength = Math.max(maxLength, val.length);
      });
      column.width = Math.max(15, maxLength + 2); // set min width of 15 for better visibility
    });

    // 5. Freeze rows after filter + header (top 4 rows)
    worksheet.views = [{ state: 'frozen', ySplit: 7 }];

    // 6. Export as .xlsx
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, `Budget_Forecast_${new Date().getTime()}.xlsx`);
    });
  }

  private createPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });

    doc.setFontSize(14);
    doc.text('Budget Forecast', 14, 12);

    // =========================
    // 🔹 HEADERS
    // =========================
    const headerRow: any[] = ['Department', 'Description'];

    this.monthsShort.forEach((mn: any) => {
      if (this.avgMonths.includes(mn) && this.isEditMode) {
        headerRow.push('3 Months Avg');
      }
      headerRow.push(mn);
    });

    headerRow.push('Total');

    // =========================
    // 🔹 BODY
    // =========================
    const rows: any[] = [];

    this.PDCForecastData.forEach((row: any, i: number) => {

      const r: any[] = [];

      // Department
      if (this.variableType === 'sales') {
        r.push(row.Category || '');
      } else {
        r.push(row.ITV_CategoryName || '');
      }

      // Description
      r.push(row.ITV_Variable_Name || '');

      // Months
      this.monthsShort.forEach((mn: any) => {

        if (this.avgMonths.includes(mn) && this.isEditMode) {
          r.push(row['Avg_' + mn] ?? '');
        }

        r.push(row[mn] ?? '');
      });

      // Total
      r.push(this.getRowTotal(row) ?? '');

      rows.push(r);
    });

    let prevDept = '';
    autoTable(doc, {
      head: [headerRow],
      body: rows,
      startY: 18,

      styles: {
        fontSize: 8,
        cellPadding: 2,
        valign: 'middle',
        lineWidth: 0.2,
        lineColor: [200, 200, 200]
      },

      columnStyles: {
        0: { halign: 'left' },
        1: { halign: 'left' }
      },
      didParseCell: (data: any) => {
        if (data.section === 'head') {
          data.cell.styles.fillColor = [217, 231, 255];
          data.cell.styles.textColor = 0;
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.halign = 'center';
        }
        if (data.section === 'body') {

          const rowIndex = data.row.index;
          const val = data.cell.raw;
          const col = data.column.index;
          const rowData = this.PDCForecastData[rowIndex];
          if (col === 0) {

            if (val === prevDept) {
              data.cell.text = [''];
            } else {
              prevDept = val;
              data.cell.text = [val];
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.textColor = [5, 84, 239];
            }

            data.cell.styles.halign = 'left';
          }
          if (rowIndex % 2 === 0) {
            data.cell.styles.fillColor = [249, 251, 255];
          } else {
            data.cell.styles.fillColor = [255, 255, 255];
          }
          if (col === 0 || col === 1) {
            data.cell.styles.halign = 'left';
          } else {
            data.cell.styles.halign = 'right';
          }
          if (val === '' || val === null || val === undefined) {
            data.cell.text = ['-'];
            return;
          }
          if (!isNaN(val)) {

            if (rowData?.ValueType === 'G') {
              data.cell.text = [`$ ${Number(val).toLocaleString()}`];
            } else {
              data.cell.text = [Number(val).toLocaleString()];
            }
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
      doc.save('Budget Forecast.pdf'); // ✅ only here
    } catch (e) {
      console.error(e);
      this.toast.show('Error generating PDF', 'danger', 'Error');
    } finally {
      this.shared.spinner.hide();
    }
  }
  printPDF() {
    if (this.storeIds.length >= 6) {
      this.toast.show('Please select only up to 5 stores', 'warning', 'Warning');
      return;
    }
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
        this.spinner.hide();
        return;
      }
      const pdfFile = this.blobToFile(pdfBlob, 'Budget Forecast.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Budget Forecast');
      formData.append('file', pdfFile);
      formData.append('notes', notes);
      formData.append('from', from);

      this.shared.api.postmethod(this.comm.routeEndpoint + 'mail', formData)
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
}
