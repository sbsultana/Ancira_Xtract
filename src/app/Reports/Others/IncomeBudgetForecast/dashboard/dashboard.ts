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
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '71', 'type': 'M', 'others': 'N',
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
    if (typeof window !== 'undefined') {
      const userInfoStr = localStorage.getItem('userInfo');
      const flag = localStorage.getItem('flag');
      if (userInfoStr) {
        const userInfo = JSON.parse(userInfoStr);
        if (flag === 'V') {
          this.storeIds = [];
          console.log(userInfo, userInfo.user_Info, 'Widget Stores............');
          this.groupId = userInfo?.WidgetData?.groupid;
          const budgetStoreId = userInfo?.WidgetData?.BudgetStoreId;
          if (budgetStoreId && budgetStoreId > 0) {
            this.storeIds = [budgetStoreId];   // always array
          }
          this.selectedYear = userInfo?.WidgetData?.BudgetYear;
        } else {
          const stores = userInfo?.user_Info?.ustores;
          if (stores) {
            const lastStore = stores.split(',').pop();
            this.storeIds = lastStore ? [Number(lastStore)] : [];
          }
        }
        this.getBudgetForecastData(this.selectedFilter, this.selectedYear);
      }
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
  }

  //////////////////////////////////////////Store Filter Code Start//////////////////////////////////////////////////
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
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: { obj: { title: string; state: boolean; }; }) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Budget Forecast') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
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

  filters: string[] = ['store', 'expenses'];
  selectedFilter: string = 'store';

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
    if (!this.selectedYear) {
      this.toast.show(
        'Please Select Atleast One Store',
        'warning',
        'Warning'
      );
      return;
    }
    console.log('Filters', this.storeIds.toString(), this.selectedFilter, this.selectedYear);

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
      IV_V_Type: type,
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

    // 1. Add Filter Info Rows (top 3 rows)
    worksheet.addRow(['Store Name:', storeName]);
    worksheet.addRow(['Variable Type:', variableTypeLabel]);
    const row = worksheet.addRow(['Year:', year]);
    row.eachCell(cell => cell.alignment = { horizontal: 'left' });



    // Style Filter Info Rows
    for (let i = 1; i <= 3; i++) {
      const row = worksheet.getRow(i);
      row.getCell(1).font = { bold: true };
      row.getCell(1).alignment = { indent: 1 };
      row.getCell(2).alignment = { indent: 1 };
    }

    // 2. Header Row (row 4)
    const headerRow = [
      isStore ? 'Category' : 'FinSummary',
      'Store Variables',
      '', // Spacer
      ...this.monthsShort
    ];
    if (!isEditMode) {
      headerRow.push('Total');
    }

    const header = worksheet.addRow(headerRow);
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
      const showCategory = index === 0 || row.Category !== data[index - 1]?.Category;
      const showSummary = index === 0 || row.ITV_FinSummary !== data[index - 1]?.ITV_FinSummary;
      const category = isStore ? (showCategory ? row.Category : '') : (showSummary ? row.ITV_FinSummary : '');

      const rowData: (string | number | null)[] = [
        category,
        row.ITV_Variable_Name,
        ''
      ];

      this.monthsShort.forEach((mn) => {
        const val = row[mn];
        rowData.push(val !== '' && val != null ? +val : '');
      });

      if (!isEditMode) {
        const totalVal = row.sum;
        rowData.push(totalVal !== '' && totalVal != null ? +totalVal : '');
      }

      const addedRow = worksheet.addRow(rowData);

      const isGross = row.ITV_value_type === 'G' || row.ValueType === 'G';
      const numFmt = isGross ? '"$"#,##0' : '#,##0';

      this.monthsShort.forEach((_, i) => {
        const cell = addedRow.getCell(4 + i);
        if (typeof rowData[3 + i] === 'number') {
          cell.numFmt = numFmt;
          cell.alignment = { horizontal: 'right' };
        }
      });

      if (!isEditMode) {
        const totalCell = addedRow.getCell(4 + this.monthsShort.length);
        const val = row.sum;
        if (typeof val === 'number') {
          totalCell.numFmt = numFmt;
          totalCell.alignment = { horizontal: 'right' };
          totalCell.font = { bold: true };
        }
      }
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
    worksheet.views = [{ state: 'frozen', ySplit: 4 }];

    // 6. Export as .xlsx
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, `Budget_Forecast_${new Date().getTime()}.xlsx`);
    });
  }

}
