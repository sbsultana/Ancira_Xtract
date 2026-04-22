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
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-incomestatementstore-details',
  standalone: true,
  imports: [NgStyle, SharedModule],
  templateUrl: './incomestatementstore-details.html',
  styleUrl: './incomestatementstore-details.scss'
})
export class IncomestatementstoreDetails {
  @Input() Fsdetails: any;
  as_id: any = [];
  FSDetailsData: any = [];
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
  payrollInfo: string = 'N';
  constructor(
    private ngbmodel: NgbModal,
    private renderer: Renderer2,
    private apiSrvc: Api,
    private ngbmodalActive: NgbActiveModal,
    private datepipe: DatePipe,
    private comm: common,
  ) {
    this.renderer.listen('window', 'click', (e: Event) => {
      const TagName = e.target as HTMLButtonElement;
      console.log(TagName.className);
      // if (TagName.className === 'container-fluid d-flex justify-content-center align-items-center') {
      //   this.onclose()
      // }
      if (TagName.className === 'modal fade bd-example-modal-xl') {
        // this.componentState = false
        this.Opacity = 'N';
      }
      if (TagName.className === 'fa-solid fa-xmark') {
        this.Opacity = 'N';
      }
      if (TagName.className === 'close-btn ms-auto me-0') {
        this.Opacity = 'N';
      }
    });
    this.apiSrvc.payroll$.subscribe(value => {
      this.payrollInfo = value;
    });
  }

  ngOnInit(): void {
    console.log(this.Fsdetails);
    this.GetDetails();
    // this.GetSubDetails();
  }

  getStoresId(Value: any) {
    console.log(Value);
  }
  SelectMonthYear: any;
  ETdetailsData: any = [];
  currentPage: number = 1;
  itemsPerPage: number = 100;
  maxPageButtonsToShow: number = 3;
  clickedPage: number | null = null;
  filteredFSdetailsData: any[] = [];
  selectedDate: any;
  DateType: any;
  // DetailsSearchName: any;
  Lable: any;
  searchText: string = '';
  GetDetails() {
    this.NoData = false;
    this.Opacity = 'N';
    const format = 'MM-yyyy';
    const locale = 'en-US';
    const myDate = this.Fsdetails.LatestDate.replace(/-/g, '/');
    const formattedDate = formatDate(myDate, format, locale);
    this.LatestDate = formattedDate;
    const SelectedMY = this.datepipe.transform(
      new Date(this.Fsdetails.LatestDate),
      'MMM-yy'
    );
    this.SelectMonthYear = SelectedMY;
    // const Obj = {
    //   Type: this.Fsdetails.TYPE,
    //   Date: this.LatestDate,
    //   Stores:
    //     this.Fsdetails.STORES == undefined ||
    //     this.Fsdetails.STORES == null ||
    //     this.Fsdetails.STORES == 0
    //       ? ''
    //       : this.Fsdetails.STORES,
    //   PageNumber: this.pageNumber,
    //   PageSize: '100',
    // };
    const Obj = {
      "as_ids": this.Fsdetails.STORES == undefined ||
        this.Fsdetails.STORES == null ||
        this.Fsdetails.STORES == 0
        ? ''
        : this.Fsdetails.STORES,
      "dept": "",
      "subtype": "",
      "subtypedetail": this.Fsdetails.NAME,
      "FinSummary": this.Fsdetails.TYPE,
      "date": this.LatestDate,
      "accountnumber": ""
    }
    console.log(Obj);

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetFinancialSummaryDetailsV3', Obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.FSDetailsData = res.response.map((item: any) => ({
            ...item,
            AccountDescription: item.AccountDescription
              ? item.AccountDescription
                .toLowerCase()
                .replace(/\b\w/g, (char: string) => char.toUpperCase())
              : item.AccountDescription
          }));

          // this.filterData();
          this.filteredFSdetailsData = this.FSDetailsData || [];
          console.log(this.FSDetailsData);
          this.spinnerLoader = false;
          // if (this.FSDetailsData.length > 0) {
          //   this.NoData = false;
          // } else {
          this.NoData = true;
          // }
        }
      });
  }
  expandedIndex: number | null = null;
  FSSubDetailsMap: { [index: number]: any[] } = {};

  GetSubDetails(AcctNo: any, StoreName: any, index: number) {
    // Toggle: If same row clicked again, collapse it
    if (this.expandedIndex === index) {
      this.expandedIndex = null;
      return;
    }

    this.expandedIndex = index;

    this.spinnerLoader = true;
    const format = 'MM-yyyy';
    const locale = 'en-US';
    const myDate = this.Fsdetails.LatestDate.replace(/-/g, '/');
    const formattedDate = formatDate(myDate, format, locale);
    this.LatestDate = formattedDate;

    const SelectedMY = this.datepipe.transform(
      new Date(this.Fsdetails.LatestDate),
      'MMM-yy'
    );
    this.SelectMonthYear = SelectedMY;

    const Obj = {
      as_ids: StoreName,
      dept: '',
      subtype: '',
      subtypedetail: this.Fsdetails.NAME,
      FinSummary: this.Fsdetails.TYPE,
      date: this.LatestDate,
      accountnumber: AcctNo,
    };

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetFinancialSummaryDetailsV3', Obj)
      .subscribe((res) => {
        this.spinnerLoader = false;
        if (res.status === 200) {
          this.FSSubDetailsMap[index] = res.response.map((sub: any) => ({
            ...sub,
            DetailDescription: sub.DetailDescription
              ? sub.DetailDescription
                .toLowerCase()
                .replace(/\b\w/g, (char: string) => char.toUpperCase())
              : sub.DetailDescription
          }));
        }
      });
  }
  get postingAmountTotal(): number {
    return this.filteredFSdetailsData.reduce((total, item) => {
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
      this.filteredFSdetailsData = [...this.FSDetailsData];
    } else {
      this.filteredFSdetailsData = this.FSDetailsData.filter((item: any) =>
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
  get hasPayrollData(): boolean {
    if (this.payrollInfo === 'Y') {
      return this.paginatedItems?.some(
        (item: any) => item.HasPayroll === ''
      ) || false;
    }
    if (this.payrollInfo === 'N') {
      return this.paginatedItems?.some(
        (item: any) => item.HasPayroll === 'Y'
      ) || false;
    }
    return false;
  }
  get paginatedItems() {
    const start = (this.currentPage - 1) * this.itemsPerPage;
    return this.filteredFSdetailsData.slice(start, start + this.itemsPerPage);
  }

  getMaxPageNumber(): number {
    return Math.max(1,
      Math.ceil(this.filteredFSdetailsData.length / this.itemsPerPage)
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
      this.filteredFSdetailsData.length
    );
  }
  onPageSizeChange() {
    this.currentPage = 1;
  }

  resetExpand() {
    this.clickedPage = null;
    this.expandedIndex = null;
  }
  updateVerticalScroll(event: any): void {
    if (
      event.target.scrollTop + event.target.clientHeight >=
      event.target.scrollHeight
    ) {
      if (this.pageNumber == 1 && this.details.length >= 100) {
        this.spinnerLoader = true;
        this.pageNumber++;
        this.GetDetails();
      } else {
        if (this.details.length >= 100) {
          this.spinnerLoader = true;
          this.pageNumber++;
          this.GetDetails();
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
    this.filteredFSdetailsData = [];
    if (this.Fsdetails.STORES != '') {
      this.goToFirstPage();
    }
  }
  onclose() {
    this.ngbmodel.dismissAll();
    console.log(this.Opacity);
  }

  Account_Details: any = [];
  AcctDetails: any = [];
  Acct_ID: any;
  Obj: any;
  AccountDetails(Object: any) {
    this.NoData = false;
    this.Opacity = 'Y';
    this.Account_Details = [];
    this.componentState = true;
    console.log(Object);
    console.log(this.LatestDate);
    this.spinnerLoader = true;
    this.Acct_ID = Object.Account_ID;
    this.Obj = {
      AccountNumber: Object.Account_ID,
      DateVal: this.LatestDate,
      Store:
        Object.STORE_ID == undefined ||
          Object.STORE_ID == null ||
          Object.STORE_ID == 0
          ? ''
          : Object.STORE_ID,
    };

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetFinacialSummaryTransactionDetails', this.Obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.AcctDetails = res.response;
          if (res.response.length > 0) {
            this.Account_Details = [
              ...this.Account_Details,
              ...this.AcctDetails,
            ];
          }
          console.log(this.Account_Details);
          this.spinnerLoader = false;
          // if (this.Account_Details.length > 0) {
          //   this.NoData = false;
          // } else {
          this.NoData = true;
          // }
        }
      });
  }
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
    localStorage.setItem('date', this.Fsdetails.LatestDate);
  }

  ExcelStoreNames: any = [];
  Details_ExportAsXLSX() {
    const FSDetailsData = [...this.filteredFSdetailsData];
    const FSSubDetailsMap = this.FSSubDetailsMap;

    // Setup Excel
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Income Statement Store Details');
    const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');
    const DateToday = this.datepipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a');

    worksheet.views = [{ state: 'frozen', ySplit: 10, topLeftCell: 'A11', showGridLines: false }];

    // Header section (above grid)
    worksheet.addRow([]);
    const titleRow = worksheet.addRow(['Income Statement Store Details']);
    titleRow.font = { bold: true, size: 12 };
    worksheet.addRow([]);
    worksheet.addRow([DateToday]).font = { size: 9 };
    worksheet.addRow(['Selected Details:']).font = { bold: true, size: 10 };
    worksheet.addRow(['Type:', this.Fsdetails.NAME]);
    worksheet.addRow(['Date:', this.LatestDate]);
    worksheet.addRow(['Store:', this.Fsdetails.STORENAMES]);
    worksheet.addRow([]);

    // Grid Header
    const headers = [
      'Store Name',
      'Account Number',
      'Description',
      `Balance (${this.SelectMonthYear})`,
    ];
    const headerRow = worksheet.addRow(headers);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0554ef' },
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle', indent: 1 };
    });

    let totalBalance = 0;
    const columnWidths = new Array(headers.length).fill(10);

    // Main Data Rows
    FSDetailsData.forEach((item: any, index: number) => {
      const rowData = [
        item.StoreName || '-',
        item.AccountNumber || '-',
        item.AccountDescription || '-',
        item.PostingAmount ?? '-',
      ];
      const mainRow = worksheet.addRow(rowData);
      mainRow.font = { size: 9 };

      if (typeof item.PostingAmount === 'number') {
        mainRow.getCell(4).numFmt = '$#,##0.00';
        totalBalance += item.PostingAmount;
      }
      mainRow.getCell(4).alignment = { horizontal: 'right', indent: 1 };

      mainRow.eachCell((cell) => {
        cell.alignment = { ...cell.alignment, indent: 1 };
      });

      if (index % 2 !== 0) {
        mainRow.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e5e5e5' },
          };
        });
      }

      rowData.forEach((val, i) => {
        const length = val?.toString().length || 0;
        columnWidths[i] = Math.max(columnWidths[i], length);
      });

      const subDetails = FSSubDetailsMap[index];

      if (subDetails?.length && this.expandedIndex != null) {
        // Add Subtable Header
        const Subheaders = [
          'Control',
          'Date',
          'Detail Description',
          `Balance`,
        ];
        const subheaderRow = worksheet.addRow(Subheaders);
        subheaderRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
        subheaderRow.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '0554ef' },
          };
          cell.alignment = { horizontal: 'center', vertical: 'middle', indent: 1 };
        });

        // Add Sub Rows
        let subTotal = 0;

        subDetails.forEach((sub) => {
          const subData = [
            sub.Control || '-',
            sub.AccountingDate || '-',
            sub.DetailDescription || '-',
            sub.PostingAmount ?? '-',
          ];
          const subRow = worksheet.addRow(subData);
          subRow.font = { size: 9 };

          if (typeof sub.PostingAmount === 'number') {
            subTotal += sub.PostingAmount;
            totalBalance += sub.PostingAmount;
            subRow.getCell(4).numFmt = '$#,##0.00';
          }
          subRow.getCell(4).alignment = { horizontal: 'right', indent: 1 };

          subRow.eachCell((cell) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'd2deed' },
            };
            cell.alignment = { ...cell.alignment, indent: 1 };
          });

          subData.forEach((val, i) => {
            const length = val?.toString().length || 0;
            columnWidths[i] = Math.max(columnWidths[i], length);
          });
        });

        // Add Subtotal Row
        const subTotalRow = worksheet.addRow(['', '', 'Sub Total:', subTotal]);
        subTotalRow.font = { bold: true };
        subTotalRow.getCell(4).numFmt = '$#,##0.00';
        subTotalRow.getCell(4).alignment = { horizontal: 'right', indent: 1 };
        subTotalRow.eachCell((cell) => {
          cell.alignment = { ...cell.alignment, indent: 1 };
        });
      }
    });

    // Final Total Row
    const totalRow = worksheet.addRow(['', '', 'Total:', this.postingAmountTotal]);
    totalRow.font = { bold: true };
    totalRow.getCell(4).numFmt = '$#,##0.00';
    totalRow.getCell(4).alignment = { horizontal: 'right', indent: 1 };
    totalRow.eachCell((cell) => {
      cell.alignment = { ...cell.alignment, indent: 1 };
    });

    // Set Column Widths
    columnWidths.forEach((width, i) => {
      worksheet.getColumn(i + 1).width = Math.max(15, width + 2);
    });

    // Export to Excel File
    workbook.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, `Financial summary details ${DATE_EXTENSION}.xlsx`);
    });
  }





  AcctDetasil_ExportAsXLSX() {
    let AcctDetails = this.AcctDetails.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Income Statement Store Details');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 9,
        topLeftCell: 'A10',
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Income Statement Store Account Details']);
    titleRow.eachCell((cell, number) => {
      cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    worksheet.addRow('');

    const DateToday = this.datepipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.datepipe.transform(
      new Date(),
      'MMddyyyy'
    );
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };

    const ReportFilter = worksheet.addRow(['Selected Details :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const Type = worksheet.getCell('A6');
    Type.value = 'Account No :';
    Type.font = { name: 'Arial', family: 4, size: 9, bold: true };
    const type = worksheet.getCell('B6');
    type.value = this.Acct_ID;
    type.font = { name: 'Arial', family: 4, size: 9 };
    type.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

    const DateMonth = worksheet.getCell('A7');
    DateMonth.value = 'Date :'
    const date = worksheet.getCell('B7');
    date.value = this.LatestDate
    date.font = { name: 'Arial', family: 4, size: 9 };
    date.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    DateMonth.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };

    worksheet.addRow('');

    let Headings = [
      'Control',
      'Detail Description',
      'Accounting Date',
      'Refer',
      'Posting Amount'
    ];
    const headerRow = worksheet.addRow(Headings);
    headerRow.height = 20;
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 8,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'center' };
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '2a91f0' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'dotted' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    for (const d of AcctDetails) {
      const AcctDate = this.datepipe.transform(
        new Date(d.accountingdate),
        'MM.dd.yyyy'
      );
      const Data1 = worksheet.addRow([
        d.control == '' ? '-' : d.control == null ? '-' : d.control,
        d.detaildescription == '' ? '-' : d.detaildescription == null ? '-' : d.detaildescription,
        AcctDate == '' ? '-' : AcctDate == null ? '-' : AcctDate,
        d.refer == '' ? '-' : d.refer == null ? '-' : d.refer,
        d.PostingAmount == '' ? '-' : d.PostingAmount == null ? '-' : d.PostingAmount,
      ]);
      Data1.font = { name: 'Arial', family: 4, size: 8 };

      // Set alignment for each cell individually
      Data1.getCell(1).alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }; // Left align
      Data1.getCell(2).alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }; // Center align
      Data1.getCell(3).alignment = { indent: 1, vertical: 'middle', horizontal: 'center' }; // Left align
      Data1.getCell(4).alignment = { indent: 1, vertical: 'middle', horizontal: 'center' }; // Right align
      Data1.getCell(5).alignment = { indent: 1, vertical: 'middle', horizontal: 'right' }; // Right align
      Data1.getCell(5).numFmt = '$#,##0';
      Data1.eachCell((cell, number) => {
        cell.border = { right: { style: 'dotted' } };
      });

      if (Data1.number % 2) {
        Data1.eachCell((cell, number) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e5e5e5' },
            bgColor: { argb: 'FF0000FF' },
          };
        });
      }
    }
    worksheet.eachRow((row, rowIndex) => {
      row.eachCell((cell, colIndex) => {
        if (rowIndex > 1 && rowIndex < 9) {
          if (colIndex === 1) {
            cell.alignment = {
              horizontal: 'left',
              vertical: 'middle',
              indent: 1,
            };
          }
        }
      });
    });
    worksheet.getColumn(1).width = 15;
    worksheet.getColumn(2).width = 40;
    worksheet.getColumn(3).width = 20;
    worksheet.getColumn(4).width = 20;
    worksheet.getColumn(5).width = 20;

    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, 'Income Statement Store Account Details_' + DATE_EXTENSION + EXCEL_EXTENSION);
    });
  }
}
