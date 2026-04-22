import {
  Component,
  ElementRef,
  Input,
  OnInit,
  Renderer2,
  ViewChild,
} from '@angular/core';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import * as FileSaver from 'file-saver';
import { Workbook } from 'exceljs';
import { formatDate, DatePipe, NgStyle, NgFor, NgIf } from '@angular/common';
import { environment } from '../../../../../environments/environment.prod';
import { Api } from '../../../../Core/Providers/Api/api';
import { common } from '../../../../common';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Router } from '@angular/router';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-enterpriseincomebyexpensedepartment-details',
  standalone: true,
  imports: [NgStyle, SharedModule],
  templateUrl: './enterpriseincomebyexpensedepartment-details.html',
  styleUrl: './enterpriseincomebyexpensedepartment-details.scss',
})
export class EnterpriseincomebyexpensedepartmentDetails {
  @Input() DetailsET: any;
  as_id: any = [];
  DetailsETData: any = [];
  pageNumber: any = 1;
  LatestDate: any;
  solutionurl: any = '';
  details: any = [];
  NoData: any;
  loadMore = 1;
  spinnerLoader: boolean = true;
  componentState: boolean = false;
  Opacity: any = 'N';
  Store_Name: any;
  control: any;
  DetailsSearchName: any;
  SubDetailsSearchName: any;
  CurrentRoute: any;
  SetTitle: any;
  constructor(
    private ngbmodel: NgbModal,
    private renderer: Renderer2,
    private apiSrvc: Api,
    private ngbmodalActive: NgbActiveModal,
    private datepipe: DatePipe,
    private comm: common,
    private router: Router,
    private shared: Sharedservice
  ) {
    this.renderer.listen('window', 'click', (e: Event) => {
      const TagName = e.target as HTMLButtonElement;
      console.log(TagName.className);
      if (TagName.className === 'modal fade bd-example-modal-xl') {
        this.Opacity = 'N';
      }
      if (TagName.className === 'fa-solid fa-xmark') {
        this.Opacity = 'N';
      }
      if (TagName.className === 'close-btn ms-auto me-0') {
        this.Opacity = 'N';
      }
    });
    this.CurrentRoute = this.router.url.substring(1);
    switch (this.CurrentRoute) {
      case 'FixedIncomeExpense':
        this.SetTitle = 'Fixed Income Expense';
        break;
      case 'VariableIncomeExpense':
        this.SetTitle = 'Variable Income Expense';
        break;
      case 'EnterpriseIncomeExpense':
        this.SetTitle = 'Enterprise Income Expense';
        break;
      default:
        this.SetTitle = 'Income Expense';
        break;
    }
  }

  ngOnInit(): void {
    console.log(this.DetailsET);
    if (this.DetailsET.REFERENCE === 'LayerOneDetails') {
      this.GetLayerOneDetails();
    } else if (this.DetailsET.REFERENCE === 'LayerThreeDetails') {
      this.GetLayerThreeDetails();
    }
  }

  getStoresId(Value: any) {
    console.log(Value);
  }
  SelectMonthYear: any;
  GetLayerOneDetails() {
    this.NoData = false;
    this.Opacity = 'N';
    const SelectedMY = this.datepipe.transform(
      new Date(this.DetailsET.LatestDate),
      'yyyy-MM-dd'
    );
    this.SelectMonthYear = SelectedMY;
    const Obj = {
      LEVEL: '0',
      AS_IDs: '',
      STORENAMES: this.DetailsET.STORENAME,
      DEPARTMENT: this.DetailsET.DEPARTMENT,
      DATE: this.shared.datePipe.transform(this.DetailsET.LatestDate, 'yyyy-MM') + '-10',
      SUBTYPEDETAIL: this.DetailsET.LAYERONE.DISPLAY_LABLE,
      ACCTTYPEDETAIL: '',
      LABLECODE: this.DetailsET.LAYERONE.PARENTLABLECODE,
      ACCTNUM: '',
      ACCTDESC: '',
      Control: '',
    };
    console.log(Obj);

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetEnterpriseIncomeExpenseDetailByLevel', Obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.details = res.response;
          this.details.some(function (x: any) {
            if (x['JSON_F52E2B61-18A1-11d1-B105-00805F49916B'] != undefined) {
              x.detailsdata = JSON.parse(
                x['JSON_F52E2B61-18A1-11d1-B105-00805F49916B']
              );
            }
          });
          if (res.response.length > 0) {
            this.DetailsETData = [...this.DetailsETData, ...this.details];
          }
          console.log(this.DetailsETData);
          this.spinnerLoader = false;
          this.NoData = true;
        }
      });
  }
  GetLayerThreeDetails() {
    this.NoData = false;
    this.Opacity = 'N';
    const SelectedMY = this.datepipe.transform(
      new Date(this.DetailsET.LatestDate),
      'yyyy-MM-dd'
    );
    this.SelectMonthYear = SelectedMY;
    const Obj = {
      LEVEL: '3',
      AS_IDs: '',
      STORENAMES: this.DetailsET.STORENAME,
      DEPARTMENT: this.DetailsET.DEPARTMENT,
      DATE: this.shared.datePipe.transform(this.DetailsET.LatestDate, 'yyyy-MM') + '-10',
      SUBTYPEDETAIL: this.DetailsET.LAYERONE.DISPLAY_LABLE,
      ACCTTYPEDETAIL: '',
      LABLECODE: this.DetailsET.LAYERONE.PARENTLABLECODE,
      ACCTNUM: this.DetailsET.LAYERTHREE.AcctNum,
      ACCTDESC: this.DetailsET.LAYERTWO.Desc,
      Control: this.DetailsET.LAYERTHREE.Cntrl,
    };
    console.log(Obj);

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetEnterpriseIncomeExpenseDetailByLevel', Obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.details = res.response;
          if (res.response.length > 0) {
            this.DetailsETData = [...this.DetailsETData, ...this.details];
          }
          console.log(this.DetailsETData);
          this.spinnerLoader = false;
          this.NoData = true;
        }
      });
  }

  get layerFourAmountTotal(): number {
    let total = 0;

    if (!this.DetailsETData || this.DetailsETData.length === 0) return 0;

    const data = this.DetailsETData[0]?.detailsdata;

    for (const level0 of data) {
      for (const level1 of level0.Level0 || []) {
        for (const level2 of level1.Level1 || []) {
          for (const level3 of level2.Level2 || []) {
            for (const level4 of level3.Level3 || []) {
              if (level4.Amount && typeof level4.Amount === 'number') {
                total += level4.Amount;
              }
            }
          }
        }
      }
    }
    return total;
  }
  get postingAmountTotal(): number {
    return this.DetailsETData.reduce((total: any, item: any) => {
      return total + (item.Amount || 0);
    }, 0);
  }
  updateVerticalScroll(event: any): void {
    //  console.log(event);
    if (
      event.target.scrollTop + event.target.clientHeight >=
      event.target.scrollHeight
    ) {
      if (this.pageNumber == 1 && this.details.length >= 100) {
        this.spinnerLoader = true;
        this.pageNumber++;
        this.GetLayerOneDetails();
      } else {
        if (this.details.length >= 100) {
          this.spinnerLoader = true;
          this.pageNumber++;
          this.GetLayerOneDetails();
        }
      }
    }
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  close() {
    this.ngbmodalActive.close();
    console.log(this.Opacity);
  }
  onclose() {
    this.ngbmodel.dismissAll();
    console.log(this.Opacity);
  }

  Account_Details: any = [];
  AcctDetails: any = [];
  Acct_ID: any;
  Obj: any;

  AccountSubDetails(AcctArry: any) {
    let token = localStorage.getItem('token');
    localStorage.setItem('token', token!);
    let index = window.location.href.lastIndexOf('/');
    this.solutionurl =
      window.location.href.substring(0, index) + '/fs-accountdetails';
    window.open(this.solutionurl, AcctArry);
    console.log(
      this.solutionurl,
      window.location.href.substring(index),
      window.location.href,
      index
    );
    localStorage.setItem('Id', JSON.stringify(this.Obj));
    localStorage.setItem('date', this.DetailsET.LatestDate);
  }


  exportToExcel() {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet(this.SetTitle);

    // Row-by-row heading
    const title = `Title #: ${this.DetailsET?.LAYERONE?.DISPLAY_LABLE || '-'}`;
    const store = `Store: ${this.DetailsET?.STORENAME || '-'}`;
    const latestDate = `Date: ${this.DetailsET?.LatestDate
      ? new Date(this.DetailsET.LatestDate).toLocaleDateString('en-US', {
        month: 'short',
        year: 'numeric',
      })
      : '-'
      }`;

    worksheet.getCell('A1').value = title;
    worksheet.getCell('A1').font = { bold: true, size: 12 };
    worksheet.getCell('A1').alignment = { horizontal: 'left' };

    worksheet.getCell('A2').value = store;
    worksheet.getCell('A2').font = { bold: true, size: 12 };
    worksheet.getCell('A2').alignment = { horizontal: 'left' };

    worksheet.getCell('A3').value = latestDate;
    worksheet.getCell('A3').font = { bold: true, size: 12 };
    worksheet.getCell('A3').alignment = { horizontal: 'left' };

    worksheet.addRow([]);

    const headerRow = ['Acct. Details', 'Reference', 'Entry Desc.', 'Amount', 'Acct. Date'];
    const header = worksheet.addRow(headerRow);

    header.eachCell((cell) => {
      cell.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFCCE5FF' },
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    this.DetailsETData[0]?.detailsdata.forEach((layerZero: any) => {
      layerZero.Level0.forEach((layerOne: any) => {
        layerOne.Level1.forEach((layerTwo: any) => {
          const acctDesc = layerTwo.AcctDesc || '-';
          const acctDescRow = worksheet.addRow([acctDesc, '', '', '', '']);
          acctDescRow.font = { bold: true };
          const acctNumber = layerTwo.accountnumber || '-';
          worksheet.addRow(['   ' + acctNumber, '', '', '', '']);

          layerTwo.Level2.forEach((layerThree: any) => {
            worksheet.addRow(['      ' + (layerThree.Control || '-'), '', '', '', '']);

            if (layerThree.Control2 && layerThree.Control2 !== 0) {
              worksheet.addRow(['         ' + layerThree.Control2, '', '', '', '']);
            }

            layerThree.Level3.forEach((layerFour: any) => {
              const amount = layerFour.Amount !== null && layerFour.Amount !== undefined
                ? +layerFour.Amount
                : null;

              const dateValue = layerFour.Date ? new Date(layerFour.Date) : null;

              worksheet.addRow([
                '',
                layerFour.Refer || '-',
                layerFour.EntryDesc || '-',
                amount,
                dateValue
              ]);
            });
          });
        });
      });
    });

    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
        cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      });
    });

    // Auto-width per column
    worksheet.columns.forEach((column: any) => {
      let maxLength = 10;
      column.eachCell({ includeEmpty: true }, (cell: any) => {
        const val = cell.value;
        const length = val ? val.toString().length : 0;
        if (length > maxLength) {
          maxLength = length;
        }
      });
      column.width = maxLength + 2;
    });

    worksheet.getColumn(4).numFmt = '"$"#,##0.00;[Red]-"$"#,##0.00';
    worksheet.getColumn(4).alignment = { horizontal: 'right' };

    worksheet.getColumn(5).numFmt = 'mm.dd.yyyy';
    worksheet.getColumn(5).alignment = { horizontal: 'center' };
    worksheet.getColumn(5).width = 20;

    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, `${this.SetTitle} Details${latestDate}.xlsx`);
    });
  }
  exportToSubExcel() {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet(this.SetTitle);

    const formattedDate = this.DetailsET?.LatestDate
      ? new Date(this.DetailsET.LatestDate).toLocaleDateString('en-US', { month: 'short', year: 'numeric' })
      : '-';
    worksheet.addRow([`Control #: ${this.DetailsET?.LAYERTHREE?.Desc || '-'}`]);
    worksheet.addRow([`Store: ${this.DetailsET?.STORENAME || '-'}`]);
    worksheet.addRow([`Account #: ${this.DetailsET?.LAYERTHREE?.AcctNum || '-'}`]);
    worksheet.addRow([`Date: ${formattedDate}`]);
    worksheet.addRow([]);

    for (let i = 1; i <= 4; i++) {
      worksheet.getRow(i).font = { bold: true };
    }

    const headerRow = worksheet.addRow([
      'Acct. Date', 'Control 2', 'Reference', 'Entry Desc.', 'Amount', 'Post Desc.', 'Journal'
    ]);
    headerRow.font = { bold: true };
    headerRow.eachCell((cell) => {
      cell.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFCCE5FF' },
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    this.DetailsETData.forEach((row: any) => {
      const acctDate = row['Acct. Date'] ? new Date(row['Acct. Date']) : null;

      const dataRow = worksheet.addRow([
        acctDate,
        row['control2'] || '-',
        row['Reference'] || '-',
        row['Entry Desc'] || '-',
        row['Amount'] != null ? +row['Amount'] : null,
        row['Post Desc'] || '-',
        row['Journal'] || '-'
      ]);

      dataRow.getCell(5).numFmt = '"$"#,##0.00';
      dataRow.getCell(5).alignment = { horizontal: 'right' };
      if (acctDate) {
        dataRow.getCell(1).numFmt = 'mm.dd.yyyy';
        dataRow.getCell(1).alignment = { horizontal: 'left' };
      }
    });
    worksheet.columns.forEach((column: any) => {
      let maxLength = 10;
      column.eachCell({ includeEmpty: true }, (cell: any) => {
        const columnLength = cell.value ? cell.value.toString().length : 10;
        if (columnLength > maxLength) {
          maxLength = columnLength;
        }
      });
      column.width = maxLength + 2;
    });

    workbook.xlsx.writeBuffer().then(buffer => {
      const blob = new Blob([buffer], { type: 'application/octet-stream' });
      FileSaver.saveAs(blob, `${this.SetTitle} Account Details${formattedDate}.xlsx`);
    });
  }
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
}
