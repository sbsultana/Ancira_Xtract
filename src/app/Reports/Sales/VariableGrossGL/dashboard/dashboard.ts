import { Component, HostListener, OnInit, signal } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Api } from '../../../../Core/Providers/Api/api';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Subscription } from 'rxjs';
import { NgbActiveModal } from '@ng-bootstrap/ng-bootstrap';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { DatePipe } from '@angular/common';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrls: ['./dashboard.scss']
})
export class Dashboard implements OnInit {
  SalesData: any = [];
  IndividualSalesGross: any = [];
  TotalSalesGross: any = [];
  TotalReport: any = 'B';
  TotalSortPosition: any = 'B';
  saleType: any = ['Retail'];
  Department: any = ['New', 'Used'];
  date: any = '';
  DupDate: any = ''
  CurrentDate = new Date();
  NoData: boolean = false;


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


  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  month!: Date;
  DuplicatDate!: Date;
  minDate!: Date;
  maxDate!: Date;
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/yyyy',
    minMode: 'month'
  };

  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  constructor(
    public apiSrvc: Api, private ngbmodalActive: NgbActiveModal, private toast: ToastService, public shared: Sharedservice, private datepipe: DatePipe) {
    this.shared.setTitle(this.shared.common.titleName + '-Variable Gross GL');
    let today = new Date();
    let lastMonth = new Date()
    let enddate = new Date(today.setDate(today.getDate() - 1));
    this.month = new Date(enddate.setMonth(enddate.getMonth() - 1))
    this.maxDate = new Date();
    this.minDate = new Date();
    this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    this.maxDate.setMonth(this.maxDate.getMonth());
    if (today.getDate() < 5) {
      this.date = new Date(enddate.setMonth(enddate.getMonth() - 1));
    } else {
      this.date = new Date(enddate.setMonth(enddate.getMonth()));
    }
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
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
    this.setHeaderData()
    this.getSalesData();
  }
  ngOnInit(): void {
  }
  setHeaderData() {
    const data = {
      title: 'Variable Gross GL',
      stores: this.storeIds,
      saleType: this.saleType.toString(),
      toporbottom: this.TotalReport,
      groups: this.groupId,
      Department: this.Department,
      Month: this.date,
    };
    this.apiSrvc.SetHeaderData({ obj: data });
  }
  getSalesData() {
    if (this.storeIds != '') {
      this.shared.spinner.show();
      this.GetData();
      // this.GetTotalData();
    } else {
      this.NoData = true;
    }
  }
  Datebinding: any = ''
  GetData() {
    this.IndividualSalesGross = [];
    this.DupDate = this.date
    let date = new Date()
    let enddate = new Date(date.setDate(date.getDate() - 1));
    this.shared.datePipe.transform(this.date, 'yyyy-MM') == this.shared.datePipe.transform(enddate, 'yyyy-MM') ? this.Datebinding = 'MTD' : this.Datebinding = this.shared.datePipe.transform(this.date, 'MMM yy')
    const obj = {
      DATE: this.shared.datePipe.transform(this.date, 'yyyy-MM') + '-' + ('0' + enddate.getDate()).slice(-2),
      AS_IDS: this.storeIds.toString(),
      DEALTYPE: this.Department.indexOf('New') >= 0 && this.Department.indexOf('Used') >= 0 ? '' : this.Department.toString(),
      SALETYPE: this.saleType.toString() == 'All' ? '' : this.saleType.toString()
    };
    this.apiSrvc
      .postmethod(this.shared.common.routeEndpoint + 'GetSalesGrossGLSummaryReport', obj)
      .subscribe(
        (res) => {
          const currentTitle = document.title;
          if (res.status == 200) {
            this.IndividualSalesGross = [];
            this.TotalSalesGross = [];
            if (res.response != undefined) {
              if (res.response.length > 0) {
                let array = res.response.map((v: any) => ({
                  ...v,
                  Dealer: '-',
                }));
                let data = array.reduce(
                  (r: any, { STORE, ...rest }: any) => {
                    if (!r.some((o: any) => o.STORE == STORE)) {
                      r.push({
                        STORE,
                        ...rest,
                        subdata: array.filter(
                          (v: any) => v.STORE == STORE
                        ),
                      });
                    }
                    return r;
                  },
                  []
                );
                this.TotalSalesGross = data.filter(
                  (e: any) => e.STORE == 'Report Total'
                );
                this.IndividualSalesGross = data.filter(
                  (i: any) => i.STORE != 'Report Total'
                );
                this.SalesData = this.IndividualSalesGross
                if (this.TotalReport == 'B') {
                  this.TotalSalesGross.forEach((val: any) => {
                    this.IndividualSalesGross.push(val);
                  });
                } else {
                  this.TotalSalesGross.forEach((val: any) => {
                    this.IndividualSalesGross.unshift(val);
                  });
                }
                this.shared.spinner.hide();
              } else {
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            this.toast.show(res.status, 'danger', 'Error');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }
  DateType: any = 'MTD';
  datetype() {
    if (this.DateType == 'PM') {
      return 'SP';
    }
    else if (this.DateType == 'C') {
      return 'C'
    }
    return this.DateType;
  }
  Salesdetails: any = []
  openDetails(Item: any, ParentItem: any, cat: any) {

    this.Salesdetails = [{
      StartDate: this.shared.datePipe.transform(this.date, 'yyyy-MM') + '-01',
      saletype: this.saleType,

      var1: Item.DEPT,
      var2: this.saleType,
      var3: '',

      var1Value: Item.data1,
      var2Value: '',
      var3Value: '',

      userName: Item.STORE,
      storeids: Item.STOREID == 0 ? this.storeIds : Item.STOREID,

      DepartmentN: this.Department.includes('New') ? 'N' : '',
      DepartmentU: this.Department.includes('Used') ? 'U' : '',

      parent: ParentItem,
      category: cat
    }];
    this.expandedIndex = null;
    this.FSSubDetailsMap = {};
    this.currentPage = 1;
    this.GetDetails('', 0);
  }
  details: any = [];
  spinnerLoader: boolean = true;
  DetailsSearchName: any;
  SubDetailsSearchName: any;
  acctNo: any = '';
  SalesdetailsData: any = [];
  filteredSalesdetailsData: any[] = [];
  searchText: string = '';
  ETdetailsData: any = [];
  currentPage: number = 1;
  itemsPerPage: number = 100;
  maxPageButtonsToShow: number = 3;
  clickedPage: number | null = null;
  Opacity: any = 'N';


  GetDetails(acctno: any, index: number) {

    // 🔹 MAIN DATA LOAD
    if (!acctno) {
      this.spinnerLoader = true;

      const obj = {
        AS_IDS: this.Salesdetails[0].storeids,
        DATE: this.Salesdetails[0].StartDate,
        VAR1: this.Salesdetails[0].var1,
        VAR2: this.Salesdetails[0].var2.toString(),
        Accountnumber: ''
      };

      this.apiSrvc
        .postmethod(this.shared.common.routeEndpoint + 'GetSalesGrossGLDetailsV1', obj)
        .subscribe((res: any) => {

          if (res.status === 200) {

            this.SalesdetailsData = (res.response || []).map((item: any) => ({
              Store: item.Store || item.As_dealername || '-',
              AccountNumber: item.AccountNumber || '-',
              AccountDescription: item['Account Description'] || '-',
              Balance: item.Balance || 0
            }));

            this.filterData();
            this.spinnerLoader = false;
            this.NoData = this.SalesdetailsData.length === 0;
          }
        });

      return;
    }

    // 🔹 TOGGLE CLOSE (if already open)
    if (this.expandedIndex === index) {
      this.expandedIndex = null;
      return;
    }

    // 🔹 IF DATA ALREADY LOADED → JUST OPEN (no API call again)
    if (this.FSSubDetailsMap[index]) {
      this.expandedIndex = index;
      return;
    }

    // 🔹 LOAD SUB DETAILS
    this.spinnerLoader = true;

    const obj = {
      AS_IDS: this.Salesdetails[0].storeids,
      DATE: this.Salesdetails[0].StartDate,
      VAR1: this.Salesdetails[0].var1,
      VAR2: this.Salesdetails[0].var2.toString(),
      Accountnumber: acctno
    };

    this.apiSrvc
      .postmethod(this.shared.common.routeEndpoint + 'GetSalesGrossGLDetailsV1', obj)
      .subscribe((res: any) => {

        if (res.status === 200) {

          this.FSSubDetailsMap[index] = (res.response || []).map((item: any) => ({
            Control: item.Control || '-',
            Date: item.Date,
            AccountDescription: item['Account Description'] || '-',
            Balance: item.Balance || 0
          }));

          this.expandedIndex = index;
          this.spinnerLoader = false;
        }
      });
  }
  expandedIndex: number | null = null;
  FSSubDetailsMap: { [index: number]: any[] } = {};


  get postingAmountTotal(): number {
    return this.filteredSalesdetailsData.reduce((total, item) => {
      return total + (item.Balance || 0);
    }, 0);
  }
  getPostingSubAmountTotal(index: number): number {
    const subRows = this.FSSubDetailsMap[index];
    if (!subRows || !Array.isArray(subRows)) {
      return 0;
    }

    return subRows.reduce((total, item) => {
      return total + (item.Balance || 0);
    }, 0);
  }

  filterData() {
    const text = this.searchText.trim().toLowerCase();

    if (!text) {
      this.filteredSalesdetailsData = [...this.SalesdetailsData];
    } else {
      this.filteredSalesdetailsData = this.SalesdetailsData.filter((item: any) =>
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
    return this.filteredSalesdetailsData.slice(start, start + this.itemsPerPage);
  }

  getMaxPageNumber(): number {
    return Math.max(1,
      Math.ceil(this.filteredSalesdetailsData.length / this.itemsPerPage)
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
      this.filteredSalesdetailsData.length
    );
  }
  onPageSizeChange() {
    this.currentPage = 1;
  }

  resetExpand() {
    this.clickedPage = null;
    this.expandedIndex = null;
  }


  close() {
    this.ngbmodalActive.close();
    console.log(this.Opacity);
    this.filteredSalesdetailsData = [];
    if (this.Salesdetails.STORES != '') {
      this.goToFirstPage();
    }
  }
  onclose() {
    this.shared.ngbmodal.dismissAll();
    console.log(this.Opacity);
  }

  Account_Details: any = [];
  AcctDetails: any = [];
  Acct_ID: any;
  Obj: any;



  isDesc: boolean = false;
  column: string = 'CategoryName';
  sort(property: any, data: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // //console.log(direction);
    if (direction == -1) {
      this.TotalSortPosition = 'T';
    } else {
      this.TotalSortPosition = 'B';
    }
    data.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
  subdataindex: any = 0;
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Variable Gross GL') {
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
      if (!obj || obj.title !== 'Variable Gross GL') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Variable Gross GL') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Variable Gross GL') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Variable Gross GL') return;
      if (obj.stateEmailPdf) {
        this.sendEmailData(obj.Email, obj.notes, obj.from);
      }
    });
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
      this.date = event.date;
      return;
    };
  }
  changeDate(e: any) {
    console.log(e);
    this.date = e;
  }

  ngOnDestroy(): void {
    this.excel?.unsubscribe();
    this.print?.unsubscribe();
    this.pdf?.unsubscribe();
    this.email?.unsubscribe();
  }

  multipleorsingle(block: any, e: any) {

    if (block == 'RL') {
      this.saleType = []
      this.saleType.push(e);
    }

    if (block == 'TB') {
      this.TotalReport = e;

    }

    if (block == 'Dept') {
      const index = this.Department.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Department.splice(index, 1);
        if (this.Department.length == 0) {
          this.toast.show('Please Select Atleast One Department', 'warning', 'Warning');
        }
      } else {
        this.Department.push(e);
      }
    }
  }
  activePopover: number = -1;
  togglePopover(popoverIndex: number) {

    if (this.activePopover === popoverIndex) {
      // If the same popover is clicked, close it
      this.activePopover = -1;
    } else {
      // Open the selected popover and close others
      this.activePopover = popoverIndex;
    }
  }
  viewreport() {

    this.activePopover = -1
    if (this.storeIds.length == 0 || this.Department.length == 0) {
      if (this.Department.length == 0) {
        this.toast.show('Please Select Atleast One Department', 'warning', 'Warning');
      }
      if (this.storeIds.length == 0) {
        this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      }
    } else {
      this.setHeaderData()
      this.getSalesData()
    }
    // }
    // }
  }
  Scrollpercent: any = 0;
  updateVerticalScroll(event: any): void {
    const scrollDemo = document.querySelector('#scrollcent');
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo!.clientHeight)) *
      100
    );
  }
  index = '';
  ExcelStoreNames: any = [];
  generatePDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save('Variable Gross GL.pdf'); // ✅ only here
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
  generatePDFBlob(): Blob | null {
    try {
      const doc = this.createPDF()
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
      const pdfFile = this.blobToFile(pdfBlob, 'Variable Gross GL.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Variable Gross GL');
      formData.append('file', pdfFile);
      formData.append('notes', notes);
      formData.append('from', from);

      this.shared.api.postmethod(this.shared.common.routeEndpoint + 'mail', formData)
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
  private formatNumber(val: any): string {
    if (val == null || val === '' || val === 0) return '-';
    return Math.round(Number(val)).toLocaleString('en-US');
  }

  private formatCurrency(val: any): string {
    if (val == null || val === '' || val === 0) return '-';
    return '$ ' + Math.round(Number(val)).toLocaleString('en-US');
  }

  private formatPercent(val: any): string {
    if (val == null || val === '') return '-';

    const num = Number(val);

    // If whole number → no decimals
    if (Number.isInteger(num)) {
      return num + '%';
    }

    // If decimal → show max 2 decimals (no trailing zeros)
    return parseFloat(num.toFixed(2)) + '%';
  }

  private createPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });

    doc.setFontSize(14);
    doc.text('Variable Gross GL', 14, 12);

    const startY = 18;

    /* ================= DATE ================= */
    const d = new Date(this.DupDate);
    const monthYear =
      `${d.toLocaleString('default', { month: 'long' })} - ${d.getFullYear()}`;

    /* ================= HEADER ================= */
    const head = [
      [
        { content: '', rowSpan: 1 },
        { content: 'Units', colSpan: 8, styles: { halign: 'center' } },
        { content: 'Gross', colSpan: 11, styles: { halign: 'center' } }
      ],
      [
        monthYear,
        this.Datebinding || 'MTD',
        'Pace', 'Target', 'Diff', 'LY', 'YOY%', 'LM', 'MOM%',
        'Front Gross', 'Back Gross', 'Total Gross',
        'PVR', 'Pace', 'Target', 'Diff',
        'LY', 'YOY%', 'LM', 'MOM%'
      ]
    ];

    const body: any[] = [];

    /* ================= BODY ================= */
    this.SalesData.forEach((d: any) => {

      if (!d.STORE || d.STORE.trim() === '-' || d.STORE.trim() === '') return;

      const isTotalRow = d.STORE.trim().toLowerCase() === 'report total';

      const mainRow = [
        d.STORE,

        // UNITS
        this.formatNumber(d.UNITS_MTD),
        this.formatNumber(d.UNITS_PACE),
        this.formatNumber(d.UNITS_TARGET),
        this.formatNumber(d.UNITS_DIFF),
        this.formatNumber(d.UNITS_LY),
        this.formatPercent(d.UNITS_YOY),
        this.formatNumber(d.UNITS_LM),
        this.formatPercent(d.UNITS_MOM),

        // GROSS ($)
        this.formatCurrency(d.GROSS_FG),
        this.formatCurrency(d.GROSS_BG),
        this.formatCurrency(d.GROSS_MTD),
        this.formatCurrency(d.GROSS_PVR),
        this.formatCurrency(d.GROSS_PACE),
        this.formatCurrency(d.GROSS_TARGET),
        this.formatCurrency(d.GROSS_DIFF),
        this.formatCurrency(d.GROSS_LY),
        this.formatPercent(d.GROSS_YOY),
        this.formatCurrency(d.GROSS_LM),
        this.formatPercent(d.GROSS_MOM)
      ];

      body.push({
        row: mainRow,
        indent: 0,
        isTotalRow,
        isStoreRow: true
      });

      /* ===== SUB ROWS ===== */
      if (d.subdata && d.subdata.length > 0) {
        d.subdata.forEach((sd: any) => {

          if (!sd.DEPT || sd.DEPT.trim() === '-') return;

          const subRow = [
            sd.DEPT,

            // UNITS
            this.formatNumber(sd.UNITS_MTD),
            this.formatNumber(sd.UNITS_PACE),
            this.formatNumber(sd.UNITS_TARGET),
            this.formatNumber(sd.UNITS_DIFF),
            this.formatNumber(sd.UNITS_LY),
            this.formatPercent(sd.UNITS_YOY),
            this.formatNumber(sd.UNITS_LM),
            this.formatPercent(sd.UNITS_MOM),

            // GROSS
            this.formatCurrency(sd.GROSS_FG),
            this.formatCurrency(sd.GROSS_BG),
            this.formatCurrency(sd.GROSS_MTD),
            this.formatCurrency(sd.GROSS_PVR),
            this.formatCurrency(sd.GROSS_PACE),
            this.formatCurrency(sd.GROSS_TARGET),
            this.formatCurrency(sd.GROSS_DIFF),
            this.formatCurrency(sd.GROSS_LY),
            this.formatPercent(sd.GROSS_YOY),
            this.formatCurrency(sd.GROSS_LM),
            this.formatPercent(sd.GROSS_MOM)
          ];

          body.push({ row: subRow, indent: 1 });
        });
      }

    });

    /* ================= TABLE ================= */
    autoTable(doc, {
      startY,
      head: head as any,
      body: body.map(b => b.row),
      theme: 'grid',

      styles: {
        fontSize: 8,
        cellPadding: 2,
        halign: 'right',
        lineWidth: 0.3,
        lineColor: [200, 200, 200],
        textColor: [20, 20, 20],
        valign: 'middle'
      },

      headStyles: {
        fillColor: [5, 84, 239],
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        lineWidth: 0.3,
        lineColor: [200, 200, 200]
      },

      didParseCell: (data: any) => {

        const b = body[data.row.index] || {};

        /* ===== 2ND HEADER ===== */
        if (data.section === 'head' && data.row.index === 1) {
          data.cell.styles.fillColor = [69, 132, 255]; // #4584FF
          data.cell.styles.textColor = [255, 255, 255];
          data.cell.styles.halign = 'center';
        }

        if (data.section === 'body') {

          // indent sub rows
          if (b?.indent && data.column.index === 0) {
            data.cell.styles.paddingLeft = 6;
          }

          // alignment
          data.cell.styles.halign = data.column.index === 0 ? 'left' : 'right';

          // REPORT TOTAL
          if (b?.isTotalRow) {
            data.cell.styles.fillColor = [141, 180, 255];
            data.cell.styles.fontStyle = 'bold';
          }

          // STORE ROWS
          else if (b?.isStoreRow) {
            data.cell.styles.fillColor = [217, 231, 255];
          }

          // NEGATIVE VALUES
          const raw = String(data.cell.raw).replace(/[^0-9.-]/g, '');
          const val = parseFloat(raw);

          if (!isNaN(val) && val < 0) {
            data.cell.styles.textColor = [255, 0, 0];
          }
        }
      }
    });

    return doc;
  }
  exportToExcel() {

    /* ================= STORE NAMES ================= */
    let storeNames: any = [];
    let store = this.storeIds;

    storeNames = this.shared.common.groupsandstores
      .filter((v: any) => v.sg_id == this.groupId)[0]
      .Stores.filter((item: any) =>
        store.some((cat: any) => cat === item.ID.toString())
      );

    this.ExcelStoreNames =
      store.length ==
        this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.length
        ? 'All Stores'
        : storeNames.map((a: any) => a.storename);

    const SalesData = this.SalesData.map((e: any) => ({ ...e }));

    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Variable Gross GL');

    /* ================= HEADER STYLE ================= */
    const applyHeaderStyle = (row: any, isTop: boolean) => {
      row.eachCell((cell: any) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: isTop ? '0554EF' : '4584FF' }
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
    };

    /* ================= DATA FORMAT ================= */
    const formatRow = (row: any, level: number = 0) => {
      row.eachCell((cell: any, colNumber: number) => {

        if (colNumber === 1) {
          cell.alignment = {
            horizontal: 'left',
            vertical: 'middle',
            indent: level * 2
          };
          return;
        }

        /* ZERO → '-' */
        if (cell.value === 0) {
          cell.value = '-';
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          return;
        }

        if (typeof cell.value === 'number') {

          /* % columns */
          if ([7, 9, 18, 20].includes(colNumber)) {
            cell.numFmt = '0.##"%"';
          }

          /* Currency */
          else if (colNumber >= 10) {
            cell.numFmt = '"$" * #,##0;[Red]"$" * -#,##0';
          }

          /* Normal */
          else {
            cell.numFmt = '#,##0';
          }

          /* Negative red */
          if (cell.value < 0) {
            cell.font = { color: { argb: 'FF0000' } };
          }
        }

        cell.alignment = { horizontal: 'right', vertical: 'middle' };
      });
    };

    /* ================= TITLE ================= */
    worksheet.addRow([]);
    const titleRow = worksheet.addRow(['Variable Gross GL']);
    titleRow.font = { size: 14, bold: true };
    worksheet.mergeCells('A2:T2');

    /* ================= FILTER INFO ================= */
    worksheet.addRow([]);

    const PresentMonth = this.shared.datePipe.transform(this.date, 'MMMM yyyy');
    const DateToday = this.shared.datePipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a');

    worksheet.addRow([DateToday]);
    worksheet.addRow(['Report Controls :']).font = { bold: true };
    worksheet.addRow(['Timeframe :', PresentMonth]);
    worksheet.addRow(['Stores :', this.ExcelStoreNames.toString()]);
    worksheet.addRow(['Department :', this.Department.toString()]);
    worksheet.addRow([]);

    /* ================= HEADER (MATCH PDF) ================= */

    const headerTop = worksheet.addRow([
      '',
      'Units', '', '', '', '', '', '', '',
      'Gross', '', '', '', '', '', '', '', '', '', ''
    ]);

    // ✅ EXACT MATCH
    worksheet.mergeCells('B10:I10'); // Units (8 cols)
    worksheet.mergeCells('J10:T10'); // Gross (11 cols)

    const headerRow = worksheet.addRow([
      PresentMonth,
      this.Datebinding || 'MTD',
      'Pace', 'Target', 'Diff', 'LY', 'YOY%', 'LM', 'MOM%',
      'Front Gross', 'Back Gross', 'Total Gross',
      'PVR', 'Pace', 'Target', 'Diff',
      'LY', 'YOY%', 'LM', 'MOM%'
    ]);

    applyHeaderStyle(headerTop, true);
    applyHeaderStyle(headerRow, false);

    /* ================= DATA ================= */
    SalesData.forEach((d: any) => {

      const Data1 = worksheet.addRow([
        d.STORE,
        d.UNITS_MTD, d.UNITS_PACE, d.UNITS_TARGET, d.UNITS_DIFF,
        d.UNITS_LY, d.UNITS_YOY, d.UNITS_LM, d.UNITS_MOM,
        d.GROSS_FG, d.GROSS_BG, d.GROSS_MTD, d.GROSS_PVR,
        d.GROSS_PACE, d.GROSS_TARGET, d.GROSS_DIFF,
        d.GROSS_LY, d.GROSS_YOY, d.GROSS_LM, d.GROSS_MOM
      ]);

      /* Store BG */
      Data1.eachCell((cell: any) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'D9E7FF' }
        };
      });

      /* Report Total */
      if (d.data1 === 'Report Total') {
        Data1.eachCell((cell: any) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '8DB4FF' }
          };
          cell.font = { bold: true };
        });
      }

      formatRow(Data1, 0);

      /* SUB DATA */
      if (d.subdata && d.subdata.length > 0) {

        d.subdata.forEach((d1: any) => {
      
          // 🚫 REMOVE ROW IF DEPT IS NULL / EMPTY
          if (!d1 || !d1.DEPT || d1.DEPT.toString().trim() === '' || d1.DEPT.toString().trim() === '-') {
            return;
          }
      
          const Data2 = worksheet.addRow([
            d1.DEPT,
            d1.UNITS_MTD, d1.UNITS_PACE, d1.UNITS_TARGET, d1.UNITS_DIFF,
            d1.UNITS_LY, d1.UNITS_YOY, d1.UNITS_LM, d1.UNITS_MOM,
            d1.GROSS_FG, d1.GROSS_BG, d1.GROSS_MTD, d1.GROSS_PVR,
            d1.GROSS_PACE, d1.GROSS_TARGET, d1.GROSS_DIFF,
            d1.GROSS_LY, d1.GROSS_YOY, d1.GROSS_LM, d1.GROSS_MOM
          ]);
      
          Data2.outlineLevel = 1;
      
          formatRow(Data2, 1);
        });
      }
    });

    /* ================= BORDERS ================= */
    worksheet.eachRow((row: any) => {
      row.eachCell((cell: any) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    /* ================= FREEZE ================= */
    worksheet.views = [{ state: 'frozen', xSplit: 1, ySplit: 11 }];

    /* ================= COLUMN WIDTH ================= */
    worksheet.columns.forEach((col: any, index: number) => {
      col.width = index === 0 ? 35 : 15;
    });

    worksheet.properties.outlineLevelRow = 1;

    /* ================= DOWNLOAD ================= */
    workbook.xlsx.writeBuffer().then(() => {
      this.shared.exportToExcel(workbook, 'Variable Gross GL');
    });
  }
  AcctDetasil_ExportAsXLSX() {
    let localarray = this.details.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Variable Gross GL');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 5, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A13', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('')

    const DateToday = this.shared.datePipe.transform(new Date(), 'MM/dd/yyyy h:mm:ss a');

    const titleRow = worksheet.getCell("A2"); titleRow.value = 'Variable Gross GL';
    titleRow.font = { name: 'Arial', family: 4, size: 15, bold: true };
    titleRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }



    const DateBlock = worksheet.getCell("L2"); DateBlock.value = DateToday;
    DateBlock.font = { name: 'Arial', family: 4, size: 10 };
    DateBlock.alignment = { vertical: 'middle', horizontal: 'center' }
    worksheet.addRow([''])




    const DATE_EXTENSION = this.shared.datePipe.transform(
      new Date(),
      'MMddyyyy'
    );



    worksheet.addRow('')
    let Headings = [
      'Sl.no',
      'Store',
      'Account #',
      'Description',
      'Balance',
      'Control',
      'Date',
    ];


    const headerRow = worksheet.addRow(Headings);
    headerRow.font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFFFF' }, }
    headerRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'center' };
    headerRow.height = 22;
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern', pattern: 'solid', fgColor: { argb: '2a91f0' }, bgColor: { argb: 'FF0000FF' }
      }
      cell.border = { right: { style: 'thin' } }
      cell.alignment = { vertical: 'middle', horizontal: 'center' }
    });

    // //console.log(localarray);
    var count = 0
    for (const d of localarray) {
      count++
      d.contractdate = this.shared.datePipe.transform(d.Date, 'MM.dd.yyyy');

      var obj = [count,
        (d['As_dealername'] == '' ? '-' : (d['As_dealername'] == null ? '-' : (d['As_dealername']))),
        (d['ACCOUNT#'] == '' ? '-' : (d['ACCOUNT#'] == null ? '-' : (d['ACCOUNT#']))),
        (d['Account Description'] == '' ? '-' : (d['Account Description'] == null ? '-' : (d['Account Description']))),
        (d.Balance == '' ? '-' : (d.Balance == null ? '-' : (parseFloat(d.Balance)))),
        (d.Control == '' ? '-' : (d.Control == null ? '-' : d.Control)),
        (d.contractdate == '' ? '-' : (d.contractdate == null ? '-' : (d.contractdate))),


      ];


      const Data1 = worksheet.addRow(obj);

      // //console.log(Data1);

      Data1.font = { name: 'Arial', family: 4, size: 8 }
      Data1.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 }
      // Data1.getCell(1).alignment = {indent: 1,vertical: 'top', horizontal: 'left'}
      Data1.eachCell((cell, number) => {
        cell.border = { right: { style: 'thin' } }
        if (number == 5) {
          cell.numFmt = '$#,##0'

        }

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
      Data1.worksheet.columns.forEach((column: any, columnIndex: any) => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell: any) => {
          const length = cell.value ? cell.value.toString().length : 10;
          if (length > maxLength) {
            maxLength = length;
          }
        });
        column.width = maxLength < 10 ? 10 : maxLength + 2; // Set a minimum width of 10
      });

      // });
      // count++
    }
    worksheet.getColumn(1).width = 16;
    worksheet.getColumn(2).width = 16;
    worksheet.getColumn(3).width = 20;
    worksheet.getColumn(4).width = 20;
    worksheet.getColumn(5).width = 20;



    worksheet.addRow([]);


    workbook.xlsx.writeBuffer().then(buffer => {
      this.shared.exportToExcel(workbook, 'Variable Gross GL Account Level Details');
    });
  }
}