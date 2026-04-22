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
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-enterprisetracking-details',
  standalone: true,
  imports: [NgStyle, SharedModule],
  templateUrl: './enterprisetracking-details.html',
  styleUrl: './enterprisetracking-details.scss'
})
export class EnterprisetrackingDetails {
  @Input() ETdetails: any;
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
  constructor(
    private ngbmodel: NgbModal,
    private renderer: Renderer2,
    private apiSrvc: Api,
    private ngbmodalActive: NgbActiveModal,
    private datepipe: DatePipe,
    private comm: common,
    private shared: Sharedservice
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
  }

  ngOnInit(): void {
    console.log(this.ETdetails);
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
    const myDate = String(this.ETdetails.MONTH).replace(/-/g, '/');
    const formattedDate = formatDate(myDate, format, locale);
    this.LatestDate = formattedDate;
    const SelectedMY = this.datepipe.transform(
      new Date(this.ETdetails.MONTH),
      'MMM-yy'
    );
    this.SelectMonthYear = SelectedMY;
    let departmentValue = this.ETdetails.DEPARTMENT == 'TOTAL' ? null : this.ETdetails.DEPARTMENT;
    const Obj = {
      "Month_Year": this.shared.datePipe.transform(this.ETdetails.MONTH,'MM-yyyy'),
      "Companyid": this.ETdetails.COMPANYID,
      "Department": departmentValue,
      "Type": this.ETdetails.TYPE,
      "Accountnumber": ""
    }
    console.log(Obj);
    console.log(Obj);

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetEnterpriseTrackingNetprofitDetailsV1', Obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.FSDetailsData = res.response;
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
    if (this.expandedIndex === index) {
      this.expandedIndex = null;
      return;
    }
    this.expandedIndex = index;
    this.spinnerLoader = true;
    const format = 'MM-yyyy';
    const locale = 'en-US';
    const myDate = String(this.ETdetails.MONTH).replace(/-/g, '/');
    const formattedDate = formatDate(myDate, format, locale);
    this.LatestDate = formattedDate;
    const SelectedMY = this.datepipe.transform(
      new Date(this.ETdetails.MONTH),
      'MMM-yy'
    );
    this.SelectMonthYear = SelectedMY;
    let departmentValue = this.ETdetails.DEPARTMENT == 'TOTAL' ? null : this.ETdetails.DEPARTMENT;
    const Obj = {
      "Month_Year": this.shared.datePipe.transform(this.ETdetails.MONTH,'MM-yyyy'),
      "Companyid": this.ETdetails.COMPANYID,
      "Department": departmentValue,
      "Type": this.ETdetails.TYPE,
      "Accountnumber": AcctNo
    }
    console.log(Obj);

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetEnterpriseTrackingNetprofitDetailsV1', Obj)
      .subscribe((res) => {
        this.spinnerLoader = false;
        if (res.status === 200) {
          this.FSSubDetailsMap[index] = res.response;
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
    if (this.ETdetails.STORES != '') {
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
    localStorage.setItem('date', this.ETdetails.MONTH);
  }

  ExcelStoreNames: any = [];
  Details_ExportAsXLSX() {
    const FSDetailsData = [...this.filteredFSdetailsData];
    const FSSubDetailsMap = this.FSSubDetailsMap;

    // Setup Excel
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Enterprise Tracking Details');
    const DATE_EXTENSION = this.datepipe.transform(new Date(), 'MMddyyyy');
    const DateToday = this.datepipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a');

    worksheet.views = [{ state: 'frozen', ySplit: 10, topLeftCell: 'A11', showGridLines: false }];

    // Header section (above grid)
    worksheet.addRow([]);
    const titleRow = worksheet.addRow(['Enterprise Tracking Details']);
    titleRow.font = { bold: true, size: 12 };
    worksheet.addRow([]);
    worksheet.addRow([DateToday]).font = { size: 9 };
    worksheet.addRow(['Selected Details:']).font = { bold: true, size: 10 };
    worksheet.addRow(['Type:', this.ETdetails.DEPARTMENT === 'TOTAL' ? 'New, Used, Service, Parts, Collision' : this.ETdetails.DEPARTMENT]);
    worksheet.addRow(['Date:', this.LatestDate]);
    worksheet.addRow(['Store:', this.ETdetails.STORE]);
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
      FileSaver.saveAs(blob, `Enterprise Tracking Details ${DATE_EXTENSION}.xlsx`);
    });
  }

}
