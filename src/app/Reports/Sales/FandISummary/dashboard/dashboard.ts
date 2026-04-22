import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Subscription } from 'rxjs';
import { environment } from '../../../../../environments/environment';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { CurrencyPipe } from '@angular/common';
import { NgbModalRef } from '@ng-bootstrap/ng-bootstrap';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { BorderStyle } from 'exceljs';
@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {

  FISummaryData: any = [];
  fiManagersname: any = '';
  spinnerLoader: boolean = false;
  NoData: any = '';
  groups: any = [1];
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;

  // Report Popup


  dealtype: any = ['Retail', 'Lease'];
  Keys: any = [
    { name: 'Counts', symbol: 'N', displayname: 'Counts (#)' },
    { name: 'Penetration', symbol: 'P', displayname: 'Penetration (%)' },
    { name: 'Income', symbol: 'D1', displayname: 'Income (Total $)' },
    { name: 'AvgIncome', symbol: 'D0', displayname: 'Income (Average $)' },
  ]

  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = 0;

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'S', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };

  DupFromDate: any = '';
  DupToDate: any = ''


  FromDate: any = '';
  ToDate: any = '';
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';


  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    Types: [
      { 'code': 'MTD', 'name': 'MTD' },
      { 'code': 'QTD', 'name': 'QTD' },
      { 'code': 'YTD', 'name': 'YTD' },
      { 'code': 'PYTD', 'name': 'PYTD' },
      { 'code': 'LY', 'name': 'Last Year' },
      { 'code': 'LM', 'name': 'Last Month' },
      { 'code': 'PM', 'name': 'Same Month PY' },
    ]
  }

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService,
  ) {
    this.shared.setTitle(this.comm.titleName + '-F & I Summary');



    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      if (localStorage.getItem('stime') != null) {
        let stime = localStorage.getItem('stime');
        if (stime != null && stime != '') {
          this.DateType = stime
          this.initializeDates(stime)
        }
      } else {
        this.DateType = 'MTD'
        this.initializeDates('MTD')
      }
      if (localStorage.getItem('flag') == 'V') {
        this.storeIds = [];
        console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.groupid
        JSON.parse(localStorage.getItem('userInfo')!).WidgetData.fistore.toString().indexOf(',') > 0 ?
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.fistore.split(',') :
          this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).WidgetData.fistore)
        this.FandIManagers()
        localStorage.setItem('flag', 'M')
      }
      else {
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
        //  this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',');

        JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.toString().indexOf(',') > 0 ?
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')[0] :
          this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids)

        // JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')[0]
        this.FandIManagers()
      }

      this.maxDate = new Date();
      this.minDate = new Date();
      this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
      this.maxDate.setDate(this.maxDate.getDate());
      this.setHeaderData()
      //   alert('HI')
      //
      this.getSalesData();

    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }

  }
  ngOnInit(): void {
  }
  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    console.log(this.FromDate, this.ToDate);
    localStorage.setItem('time', type);

    this.setDates(this.DateType)

  }

  getSalesData() {
    if (this.storeIds != undefined && this.storeIds.length > 0 && this.FromDate != '' && this.ToDate != '') {
      console.log(this.FromDate, this.ToDate, '................ From Date and Todate');

      this.shared.spinner.show();
      this.GetData();
    } else {
      // this.shared.spinner.hide();
    }
  }
  GetData() {
    // this.comm.redirectionFrom.flag = 'M'
    this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    const obj = {
      "store": this.storeIds.toString(),
      "startdate": this.FromDate,
      "enddate": this.ToDate,
      "dealtype": this.dealtype.toString(),
      "FIMGR": this.financeManagerId && this.financeManagerId.length > 0 ? this.financeManagerId.toString() : ''
    };
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetFandIMVPReportV1', obj)
      .subscribe(
        (res) => {
          this.FISummaryData = [];
          if (res && res.response && res.response.length > 0) {
            this.FISummaryData = res.response;
            console.log(res.response, '............');
            this.FISummaryData.some(function (x: any) {
              if (x.Penetration != undefined && x.Penetration != '' && x.Penetration != null) {
                x.Penetration = JSON.parse(x.Penetration);
              }
              if (x.Income != undefined && x.Income != '' && x.Income != null) {
                x.Income = JSON.parse(x.Income);
              }
              if (x.AvgIncome != undefined && x.AvgIncome != '' && x.AvgIncome != null) {
                x.AvgIncome = JSON.parse(x.AvgIncome);
              }
              if (x.Counts != undefined && x.Counts != '' && x.Counts != null) {
                x.Counts = JSON.parse(x.Counts);
                x.Counts.some(function (y: any) {
                  if (y.Products != undefined && y.Products != '' && y.Products != null) {
                    y.Products = JSON.parse(y.Products)
                  }
                })
              }
            });
            this.shared.spinner.hide();
            this.NoData = '';
          } else {
            this.shared.spinner.hide();
            this.NoData = 'No Data Found!!';
          }
        },
        (error) => {
          this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
          this.shared.spinner.hide();
          this.NoData = 'No Data Found';
        }
      );
  }
  financeManager: any = []
  financeManagerId: any = [];
  FandIManagers() {
    this.spinnerLoader=true
    const obj = {
      AS_ID: this.storeIds,
      type: 'F',
    }
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetEmployeesDev', obj).subscribe((res: any) => {
      this.spinnerLoader=false
      if (res.status == 200) {
        if (res.response && res.response.length > 0) {
          this.financeManager = res.response.filter((e: any) => e.FiName != 'Unknown');
          this.financeManagerId = this.financeManager.map(function (a: any) { return a.FiId; });
        }
      }
    })
  }

  employees(block: any, e: any) {
    if (block === 'FM') {
      const index = this.financeManagerId.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.financeManagerId.splice(index, 1);
      } else {
        this.financeManagerId.push(e);
      }
      if (this.financeManagerId.length == 1) {
        this.fiManagersname = this.financeManager.filter((val: any) => val.FiId == e)[0].FiName
      }
    }


    if (block === 'AllFM') {
      if (e === 0) {
        this.financeManagerId = this.financeManager.map(
          (fm: any) => fm.FiId
        );
      } else if (e === 1) {
        this.financeManagerId = [];
      }
    }
  }
  async openSalesModal(dealnumber: any, vin: any, storeid: any, stock: any, source: any, custno: any) {
    const module = await import('../../../../Layout/cdpdataview/deal/deal-module');
    const component = module.Deal;

    const modalRef = this.shared.ngbmodal.open(component, { size: 'xl', windowClass: 'connectedmodal' });
    modalRef.componentInstance.data = { dealno: dealnumber, vin: vin, storeid: storeid, stock: stock, source: source, custno: custno }; // Pass data to the modal component    
    modalRef.result.then((result) => {
      // this.topScroll()
      console.log(result); // Handle modal close result
    }, (reason) => {
      // this.topScroll()
      console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    });
  }
  expandDataArray: any = []
  expandorcollapse(cat: any) {
    const index = this.expandDataArray.findIndex((i: any) => i == cat);
    if (index >= 0) {
      this.expandDataArray.splice(index, 1);
    } else {
      this.expandDataArray.push(cat);
    }
  }
  getTotal(newvalue: any, usedvalue: any) {
    if (newvalue && usedvalue)
      return newvalue + usedvalue
    else if (newvalue)
      return newvalue
    else if (usedvalue)
      return usedvalue
    else
      return 0
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }
  popupReference!: NgbModalRef;
  popupvalues: any = { block: '', key: '', saletype: '', Category: '', prodname: '' }
  FISummaryDetails: any = []
  FISummaryDetailsNoData: any = ''
  detailpopup(tmp: any, data: any, block: any, key: any, prodName?: any) {
    this.popupvalues.block = block;
    this.popupvalues.key = key;
    this.popupvalues.saletype = data.Category
    this.popupvalues.Category = data.Category
    this.popupvalues.prodname = prodName ? prodName : ''
    this.FISummaryDetailsNoData = '';
    this.FISummaryDetails = []
    this.popupReference = this.shared.ngbmodal.open(tmp, { size: 'xl', backdrop: 'static', keyboard: true, centered: true, modalDialogClass: 'custom-modal' })
    this.GetFISumDetails(data, block, prodName)
  }
  closePopup() {
    if (this.popupReference) {
      this.popupReference.close();
    }
  }
  GetFISumDetails(data: any, block: any, ProductName: any) {
    console.log(data.Category);

    this.popupvalues.saletype = data.Category != 'Cash' && data.Category != 'Finance' && data.Category != 'Lease' ? '' : data.Category
    const obj = {
      "PT_Type": data.PT_ID,
      "store": this.storeIds.toString(),
      "startdate": this.FromDate,
      "enddate": this.ToDate,
      "dealtype": block == 'Total' ? '' : block,
      "saletype": data.Category == 'Total Units' ? 'Cash,Finance,Lease' : (data.Category != 'Cash' && data.Category != 'Finance' && data.Category != 'Lease' ? '' : data.Category),
      "ProdName": ProductName ? ProductName : ''

    }
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetFandIMVPReportV1Details', obj).subscribe(
      (res) => {
        if (res.response && res.response.length > 0) {
          this.FISummaryDetails = res.response
          this.FISummaryDetailsNoData = ''
        } else {
          this.FISummaryDetailsNoData = 'No Data Found'
        }
      },
      (error) => {
        this.FISummaryDetailsNoData = 'No Data Found'
      })
  }
  isDesc: boolean = false;
  column: string = 'CategoryName';
  sort(property: any, data: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
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
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit() {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'F & I Summary') {
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
      if (!obj || obj.title !== 'F & I Summary') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'F & I Summary') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'F & I Summary') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'F & I Summary') return;
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


  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
    // alert('Hello')

    this.FandIManagers()
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



  updatedDates(data: any) {
    // console.log(data);
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
  }

  setDates(type: any) {
    this.displaytime = '(' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ')';
    this.maxDate = new Date();
    this.minDate = new Date();
    this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    this.maxDate.setDate(this.maxDate.getDate());
    this.Dates.FromDate = this.FromDate;
    this.Dates.ToDate = this.ToDate;
    this.Dates.MinDate = this.minDate;
    this.Dates.MaxDate = this.maxDate;
    this.Dates.DateType = this.DateType;
    this.Dates.DisplayTime = this.displaytime;
  }


  viewDeal(dealData: any) {
    // const modalRef = this.shared.ngbmodal.open(DealRecapComponent, { size: 'md', windowClass: 'connectedmodal' });
    // modalRef.componentInstance.data = { dealno: dealData.ad_dealid, storeid: dealData.Store_Id, stock: dealData.Stock, vin: dealData.VIN, custno: dealData?.customernumber }; // Pass data to the modal component    
    // modalRef.result.then((result) => {
    //   console.log(result); // Handle modal close result
    // }, (reason) => {
    //   console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    // });
  }
  setHeaderData() {
    const headerdata = {
      title: 'F & I Summary',
      stores: this.storeIds,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groups,
      dealtype: this.dealtype,
      fimanagers: this.financeManagerId,
      datetype: this.DateType
    };
    this.shared.api.SetHeaderData({
      obj: headerdata,
    });
  }
  // Report popup Code
  activePopover: number = -1;
  togglePopover(popoverIndex: number) {
    if (this.activePopover === popoverIndex) {
      this.activePopover = -1;
    } else {
      this.activePopover = popoverIndex;
    }
  }

  FIManagerName: any = ''
  multipleorsingle(block: any, e: any) {
    if (block == 'FI') {
      if (e == 'All') {
        if (this.financeManagerId.length == this.financeManager.length) {
          this.financeManagerId = []
        } else {
          this.financeManagerId = this.financeManager.map(function (a: any) {
            return a.FiId;
          });
        }
      }
      else {
        const index = this.financeManagerId.findIndex((i: any) => i == e);
        if (index >= 0) {
          this.financeManagerId.splice(index, 1);
        } else {
          this.financeManagerId.push(e);
          if (this.financeManagerId.length == 1) {
            this.FIManagerName = this.financeManager.filter((val: any) => val.FiId == e)[0].FiName
          }
        }
      }
    }
    if (block == 'DL') {
      const index = this.dealtype.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.dealtype.splice(index, 1);
      } else {
        this.dealtype.push(e);
      }
    }
  }
  viewreport() {
    this.activePopover = -1
    if (this.storeIds.length == 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    }
    else if (this.dealtype == undefined || this.dealtype.length == 0) {
      this.toast.show('Please Select Atleast One Deal Type', 'warning', 'Warning');
    }
    else {
      this.setHeaderData();
      this.getSalesData()
    }
  }

  ExcelStoreNames: any = []
  managers: any = []
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
      const pdfFile = this.blobToFile(pdfBlob, 'F & I Summary.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'F & I Summary');
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
  generatePDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save('F & I Summary.pdf'); // ✅ only here
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
  // private formatNumber(val: any): string {
  //   if (val == null || val === '' || val === 0) return '-';
  //   return Math.round(Number(val)).toLocaleString('en-US');
  // }

  // private formatCurrency(val: any): string {
  //   if (val == null || val === '' || val === 0) return '-';
  //   return '$ ' + Math.round(Number(val)).toLocaleString('en-US');
  // }

  // private formatPercent(val: any): string {
  //   if (val == null || val === '') return '-';

  //   const num = Number(val);

  //   // If whole number → no decimals
  //   if (Number.isInteger(num)) {
  //     return num + '%';
  //   }

  //   // If decimal → show max 2 decimals (no trailing zeros)
  //   return parseFloat(num.toFixed(2)) + '%';
  // }

  formatValue(val: any, key: any): any {

    if (val === 0 || val === null || val === undefined) return '-';

    // D1 / D0
    if (key.symbol === 'D1' || key.symbol === 'D0') {
      return Number(val).toLocaleString(undefined, {
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
      });
    }

    // NUMBER
    if (key.symbol === 'N') {
      return Number(val).toLocaleString(undefined, {
        minimumFractionDigits: 0,
        maximumFractionDigits:
          key.name === 'Counts' ? 0 : 2
      });
    }

    // PERCENT
    if (key.symbol === 'P') {
      return Number(val).toFixed(1);
    }

    return val;
  }
  private createPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

    doc.setFontSize(12);
    doc.text('F & I Summary', 14, 10);

    let startY = 14;

    /* ================= HEADER (MATCH HTML) ================= */

    const monthText =
      `${this.shared.datePipe.transform(this.DupFromDate, 'MMMM')}` +
      (this.shared.datePipe.transform(this.DupFromDate, 'MMMM') !==
        this.shared.datePipe.transform(this.DupToDate, 'MMMM')
        ? ` - ${this.shared.datePipe.transform(this.DupToDate, 'MMMM')}`
        : '');

    const dateText =
      `${this.shared.datePipe.transform(this.DupFromDate, 'MM.dd.yyyy')} - ` +
      `${this.shared.datePipe.transform(this.DupToDate, 'MM.dd.yyyy')}`;

    const head = [
      [
        { content: monthText, styles: { halign: 'center' } },
        { content: this.storename, colSpan: 3, styles: { halign: 'center' } }
      ],
      [
        { content: dateText, styles: { halign: 'center' } },
        { content: 'New', styles: { halign: 'center' } },
        { content: 'Used', styles: { halign: 'center' } },
        { content: 'Total', styles: { halign: 'center' } }
      ]
    ];

    const body: any[] = [];

    /* ================= BODY (EXACT HTML LOOP) ================= */

    this.Keys.forEach((key: any) => {

      this.FISummaryData.forEach((FIInfo: any, i: number) => {

        // CATEGORY ROW (light blue)
        body.push({
          row: [key.displayname, '', '', ''],
          isCategory: true
        });

        FIInfo[key.name]?.forEach((subdata: any) => {

          // MAIN ROW
          body.push({
            row: [
              subdata.Category || '-',
              this.formatValue(subdata.New, key),
              this.formatValue(subdata.Used, key),
              this.formatValue(subdata.Total, key)
            ],
            isMain: true
          });

          // PRODUCTS (EXPANDED)
          if (subdata.Products && subdata.Products.length > 0) {

            subdata.Products.forEach((pd: any) => {

              body.push({
                row: [
                  '     ' + (pd.ProductName || '-'),
                  this.formatValue(pd.New, key),
                  this.formatValue(pd.Used, key),
                  this.formatValue(pd.Total, key)
                ],
                isSub: true
              });

            });
          }

        });

        // EMPTY ROW (spacing like HTML)
        body.push({
          row: ['', '', '', ''],
          isSpacer: true
        });

      });

    });

    /* ================= TABLE ================= */

    autoTable(doc, {
      startY,
      head: head as any,
      body: body.map(b => b.row),

      theme: 'grid',

      styles: {
        fontSize: 10,
        cellPadding: 2,
        halign: 'right',
        valign: 'middle',
        lineWidth: 0.1,
        textColor: [30, 30, 30]
      },

      headStyles: {
        fillColor: [5, 84, 239],
        textColor: [255, 255, 255],
        fontStyle: 'bold'
      },


      didParseCell: (data: any) => {

        const rowObj = body[data.row.index];
        /* ===== 2ND HEADER COLOR ===== */
        if (data.section === 'head' && data.row.index === 1) {
          data.cell.styles.fillColor = [69, 132, 255]; // #4584FF
          data.cell.styles.textColor = [255, 255, 255];
          data.cell.styles.halign = 'center';
        }


        if (data.section === 'body') {

          data.cell.styles.halign =
            data.column.index === 0 ? 'left' : 'right';


          /* ===== CATEGORY ROW ===== */
          if (rowObj?.isCategory) {
            data.cell.styles.fillColor = [217, 231, 255];
            data.cell.styles.fontStyle = 'bold';
          }
          /* SUB ROW */
          if (rowObj?.isSub && data.column.index === 0) {
            data.cell.styles.paddingLeft = 6;
          }
        }

      }
    });

    return doc;
  }

  getBorder() {
    return {
      top: { style: 'thin' as BorderStyle, color: { argb: 'FFCCCCCC' } },
      left: { style: 'thin' as BorderStyle, color: { argb: 'FFCCCCCC' } },
      bottom: { style: 'thin' as BorderStyle, color: { argb: 'FFCCCCCC' } },
      right: { style: 'thin' as BorderStyle, color: { argb: 'FFCCCCCC' } }
    };
  }
  exportToExcel(): void {

    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('F & I Summary');

    /* ================= TITLE ================= */
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['F & I Summary']);
    titleRow.font = { size: 14, bold: true };
    worksheet.mergeCells('A2:D2');

    worksheet.addRow('');

    /* ================= FILTERS ================= */

    const filterTitle = worksheet.addRow(['Report Filters']);
    filterTitle.font = { bold: true };

    const groupRow = worksheet.addRow([
      'Group:',
      this.comm.groupsandstores.find((v: any) => v.sg_id == this.groups)?.sg_name
    ]);

    const storeRow = worksheet.addRow([
      'Store:',
      this.storedisplayname?.toString().replaceAll(',', ', ') || '-'
    ]);

    const managerRow = worksheet.addRow([
      'F & I Managers:',
      this.financeManager
        .filter((item: any) => this.financeManagerId.includes(item.FiId))
        .map((item: any) => item.FiName)
        .toString()
    ]);

    const timeframeRow = worksheet.addRow([
      'Time Frame:',
      this.FromDate + ' to ' + this.ToDate
    ]);

    [groupRow, storeRow, managerRow, timeframeRow].forEach(row => {
      row.getCell(1).font = { bold: true };
    });

    worksheet.addRow('');

    /* ================= HEADER (AFTER FILTERS) ================= */

    const monthText =
      `${this.shared.datePipe.transform(this.DupFromDate, 'MMMM')}` +
      (this.shared.datePipe.transform(this.DupFromDate, 'MMMM') !==
        this.shared.datePipe.transform(this.DupToDate, 'MMMM')
        ? ` - ${this.shared.datePipe.transform(this.DupToDate, 'MMMM')}`
        : '');

    const dateText =
      `${this.shared.datePipe.transform(this.DupFromDate, 'MM.dd.yyyy')} - ` +
      `${this.shared.datePipe.transform(this.DupToDate, 'MM.dd.yyyy')}`;

    // HEADER ROW 1
    const header1 = worksheet.addRow([
      monthText,
      this.storename,
      '',
      ''
    ]);

    worksheet.mergeCells(`B${header1.number}:D${header1.number}`);

    header1.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0554EF' }
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center' };
      cell.border = this.getBorder();
    });

    // HEADER ROW 2
    const header2 = worksheet.addRow([
      dateText,
      'New',
      'Used',
      'Total'
    ]);

    header2.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4584FF' }
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center' };
      cell.border = this.getBorder();
    });


    /* ================= DATA ================= */

    for (const key of this.Keys) {

      const categoryRow = worksheet.addRow([key.displayname, '', '', '']);

      categoryRow.eachCell((cell, col) => {
        cell.font = { bold: true };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFD9E7FF' }
        };
        cell.alignment = { horizontal: col === 1 ? 'left' : 'right' };
        cell.border = this.getBorder();
      });

      for (const info of this.FISummaryData) {

        info[key.name]?.forEach((dept: any) => {

          const mainRow = worksheet.addRow([
            dept.Category || '-',
            this.formatValue(dept.New, key),
            this.formatValue(dept.Used, key),
            this.formatValue(dept.Total, key)
          ]);

          mainRow.eachCell((cell, col) => {
            cell.alignment = { horizontal: col === 1 ? 'left' : 'right' };
            cell.border = this.getBorder();

            if (typeof cell.value === 'number') {
              cell.numFmt = key.symbol == 'P' ? '#,##0.0' : '#,##0';
            }
          });

          if (dept.Products?.length) {
            dept.Products.forEach((pd: any) => {

              const subRow = worksheet.addRow([
                '     ' + (pd.ProductName || '-'),
                this.formatValue(pd.New, key),
                this.formatValue(pd.Used, key),
                this.formatValue(pd.Total, key)
              ]);

              subRow.eachCell((cell, col) => {
                if (col === 1) {
                  cell.alignment = { indent: 2 };
                } else {
                  cell.alignment = { horizontal: 'right' };
                }

                cell.border = this.getBorder();

                if (typeof cell.value === 'number') {
                  cell.numFmt = key.symbol == 'P' ? '#,##0.0' : '#,##0';
                }
              });

            });
          }

        });

        worksheet.addRow(['', '', '', '']);
      }
    }

    /* ================= COLUMN WIDTH ================= */

    worksheet.getColumn(1).width = 40;
    worksheet.getColumn(2).width = 15;
    worksheet.getColumn(3).width = 15;
    worksheet.getColumn(4).width = 15;

    /* ================= EXPORT ================= */

    this.shared.exportToExcel(workbook, 'F & I Summary');
  }
  excelDownload(): void {
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('F & I Summary Details');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 1,
        topLeftCell: 'A2',
        // showGridLines: false,
      },
    ];
    let bindingHeaders: any = []
    this.popupvalues.saletype == '' ? bindingHeaders = ['Date', 'ad_dealid', 'Stock', 'VIN', 'CustomerName', 'FIManager', 'ad_frontgross', 'ad_backgross', 'TotalGross'] :
      bindingHeaders = ['Date', 'ad_dealid', 'Stock', 'CustomerName', 'FIManager', 'ad_frontgross', 'ad_backgross', 'TotalGross']
    let headerRow: any = []
    this.popupvalues.saletype == '' ? headerRow = ['Date', 'Deal #', 'Stock', 'VIN', 'Customer Name', 'FI Manager', 'Sale', 'Cost', 'Total Gross'] :
      headerRow = ['Date', 'Deal #', 'Stock', 'Customer Name', 'FI Manager', 'Front Gross', 'Back Gross', 'Total Gross']
    const currencyFields: any = ['ad_frontgross', 'ad_backgross', 'TotalGross'];
    const excelHeader = worksheet.addRow(headerRow);
    excelHeader.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1F4E78' }
      };
      cell.alignment = { horizontal: 'center' };
    });
    for (const info of this.FISummaryDetails) {
      const rowData = bindingHeaders.map((key: any) => {
        const val = info[key];
        return val === 0 || val == null ? '-' : key == 'Date' ? this.shared.datePipe.transform(val, 'MM.dd.yyyy') : val
      });
      const dealerRow = worksheet.addRow(rowData);
      dealerRow.font = { bold: true };
      bindingHeaders.forEach((key: any, index: any) => {
        const cell = dealerRow.getCell(index + 1);
        if (currencyFields.includes(key) && typeof cell.value === 'number') {
          cell.numFmt = '"$"#,##0.00';
          cell.alignment = { horizontal: 'right' };
        } else if (!isNaN(Number(cell.value))) {
          cell.alignment = { horizontal: 'right' };
        }
      });
    }
    worksheet.columns.forEach((column: any) => {
      let maxLength = 15;
      column.eachCell({ includeEmpty: true }, (cell: any) => {
        let columnLength = 0;
        if (cell.value != null) {
          const cellText = cell.value.toString();
          columnLength = cellText.length;
        }
        maxLength = Math.max(maxLength, columnLength);
      });
      column.width = maxLength + 2;
    });
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      this.shared.exportToExcel(workbook, 'F & I Summary Details');

    });
  }
  //   window.print();
  // }
  // generatePDF() {
  //   this.shared.spinner.show();
  //   const printContents = document.getElementById('SellingGross')!.innerHTML;
  //   const iframe = document.createElement('iframe');
  //   iframe.style.position = 'absolute';
  //   iframe.style.width = '0px';
  //   iframe.style.height = '0px';
  //   iframe.style.border = 'none';
  //   document.body.appendChild(iframe);
  //   const doc = iframe.contentWindow?.document;
  //   if (!doc) {
  //     console.error('Failed to create iframe document');
  //     return;
  //   }
  //   doc.open();
  //   doc.write(`
  //   <html>
  //           <head>
  //             <title>F & I Summary</title>
  //                <style>
  //                @font-face {
  //                 font-family: 'GothamBookRegular';
  //                 src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
  //                      url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
  //               }
  //               @font-face {
  //                 font-family: 'Roboto';
  //                 src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
  //                      url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
  //               }
  //               @font-face {
  //                 font-family: 'RobotoBold';
  //                 src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
  //                      url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
  //               }
  //               .negative {
  //                 color: red;
  //               }
  //               .HideItems{
  //                 display:none;
  //                }            
  // .backgrossbtn {
  //                 float: left;
  //                 font-size: 11px;
  //                 background-color: #2d91f1;
  //                 color: white;
  //                 padding: 4px;
  //                 border: 1px solid #2d91f1;
  //                 cursor: pointer;
  //                 border-radius: 3px;
  //             }
  //               .Selectedrow {
  //                 background-color: #d3ecff !important;
  //                 color: #000 !important;
  //             }
  //               .justify-content-between {
  //                 justify-content: space-between !important;
  //             }
  //             .d-flex {
  //                 display: flex !important;
  //             }
  //               .bg-white {
  //                 background: #ffffff !important;
  //             }
  //               .negative {
  //                 color: red;
  //               }
  //               .negative {
  //                 color: red;
  //               }
  //               .performance-scorecard .table>:not(:first-child) {
  //                 border-top: 0px solid #ffa51a
  //               }
  //               .performance-scorecard .table {
  //                 text-align: center;
  //                 text-transform: capitalize;
  //                 border: transparent;
  //                 width: 100%;
  //               }
  //               .performance-scorecard .table th,
  //               .performance-scorecard .table td{
  //                 white-space: nowrap;
  //                 vertical-align: top;
  //               }
  //               .performance-scorecard .table th:first-child,
  //               .performance-scorecard .table td:first-child {
  //                 position: sticky;
  //                 left: 0;
  //                 z-index: 1;
  //                 background-color: #337ab7;
  //               }
  //               .performance-scorecard .table tr:nth-child(odd) td:first-child,
  //               .performance-scorecard .table tr:nth-child(odd) td:nth-child(2) {
  //                 // background-color: #e9ecef;
  //                 background-color: #ffffff;
  //               }
  //               .performance-scorecard .table tr:nth-child(even) td:first-child,
  //               .performance-scorecard .table tr:nth-child(even) td:nth-child(2) {
  //                 background-color: #ffffff;
  //               }
  //               .performance-scorecard .table tr:nth-child(odd) {
  //                 // background-color: #e9ecef;
  //                 background-color: #ffffff;
  //               }
  //               .performance-scorecard .table tr:nth-child(even) {
  //                 background-color: #ffffff;
  //               }
  //               .performance-scorecard .table .spacer {
  //                 // width: 50px !important;
  //                 background-color: #cfd6de !important;
  //                 border-left: 1px solid #cfd6de !important;
  //                 border-bottom: 1px solid #cfd6de !important;
  //                 border-top: 1px solid #cfd6de !important;
  //                }
  //               .performance-scorecard .table .hidden {
  //                 display: none !important;
  //               }
  //               .performance-scorecard .table .bdr-rt {
  //                 border-right: 1px solid #abd0ec;
  //               }
  //               .performance-scorecard .table thead {
  //                 position: sticky;
  //                 top: 0;
  //                 z-index: 99;
  //                 font-family: 'FaktPro-Bold';
  //                 font-size: 0.8rem;
  //               }
  //               .performance-scorecard .table thead th {
  //                 padding: 5px 10px;
  //                 margin: 0px;
  //               }
  //               .performance-scorecard .table thead .bdr-btm {
  //                 border-bottom: #005fa3;
  //               }
  //               .performance-scorecard .table thead tr:nth-child(1) {
  //                 background-color: #fff !important;
  //                 color: #000;
  //                 text-transform: uppercase;
  //                 border-bottom: #cfd6de;
  //               }
  //               .performance-scorecard .table thead tr:nth-child(2) {
  //                 background-color: #fff !important;
  //                 color: #000;
  //                 text-transform: uppercase;
  //                 border-bottom: #cfd6de;
  //                 box-shadow: inset 0 1px 0 0 #cfd6de;
  //               }
  //               .performance-scorecard .table thead tr:nth-child(2) th :nth-child(1) {
  //                 background-color: #337ab7 !important;
  //                 color: #fff;
  //               }
  //               .performance-scorecard .table thead tr:nth-child(3) {
  //                 background-color: #fff !important;
  //                 color: #000;
  //                 text-transform: uppercase;
  //                 border-bottom: #cfd6de;
  //                 box-shadow: inset 0 1px 0 0 #cfd6de;
  //               }
  //               .performance-scorecard .table thead tr:nth-child(3) th :nth-child(1) {
  //                 background-color: #337ab7 !important;
  //                 color: #fff;
  //               }
  //               .performance-scorecard .table tbody {
  //                 font-family: 'FaktPro-Normal';
  //                 font-size: .9rem;
  //               }
  //               .performance-scorecard .table tbody td {
  //                 padding: 2px 10px;
  //                 margin: 0px;
  //                 border: 1px solid #cfd6de
  //               }
  //               .performance-scorecard .table tbody tr {
  //                 border-bottom: 1px solid #37a6f8;
  //                 border-left: 1px solid #37a6f8
  //               }
  //               .performance-scorecard .table tbody td:first-child {
  //                 text-align: start;
  //                 box-shadow: inset -1px 0 0 0 #cfd6de;
  //               }
  //               .performance-scorecard .table tbody tr td:not(:first-child) {
  //                 text-align: right !important;
  //               }
  //               .performance-scorecard .table tbody .sub-title {
  //                 font-size: .8rem !important;
  //               }
  //               .performance-scorecard .table tbody .sub-subtitle{
  //                 font-size: .7rem !important;
  //               }
  //               .performance-scorecard .table tbody .text-bold {
  //                 font-family: 'FaktPro-Bold';
  //               }
  //               .performance-scorecard .table tbody .darkred-bg {
  //                 background-color: #282828 !important;
  //                 color: #fff;
  //               }
  //               .performance-scorecard .table tbody .lightblue-bg {
  //                 background-color: #646e7a !important;
  //                 color: #fff;
  //               }
  //               .performance-scorecard .table tbody .gold-bg {
  //                 background-color: #ffa51a;
  //                 color: #fff;
  //               }
  //                </style>
  //           </head>
  //       <body id='content'>
  //       ${printContents}
  //       </body>
  //         </html>`);
  //   doc.close();
  //   const div = doc.getElementById('content');
  //   const options = {
  //     logging: true,
  //     allowTaint: false,
  //     useCORS: true,
  //   };
  //   if (!div) {
  //     console.error('Element not found');
  //     return;
  //   }
  //   html2canvas(div, options)
  //     .then((canvas) => {
  //       let imgWidth = 285;
  //       let pageHeight = 206;
  //       let imgHeight = (canvas.height * imgWidth) / canvas.width;
  //       let heightLeft = imgHeight;
  //       const contentDataURL = canvas.toDataURL('image/png');
  //       let pdfData = new jsPDF('l', 'mm', 'a4', true);
  //       let position = 5;
  //       function addExtraDataToPage(pdf: any, extraData: any, positionY: any) {
  //         pdf.text(extraData, 10, positionY);
  //       }
  //       function addPageAndImage(pdf: any, contentDataURL: any, position: any) {
  //         pdf.addPage();
  //         pdf.addImage(
  //           contentDataURL,
  //           'PNG',
  //           5,
  //           position,
  //           imgWidth,
  //           imgHeight,
  //           undefined,
  //           'FAST'
  //         );
  //       }
  //       pdfData.addImage(
  //         contentDataURL,
  //         'PNG',
  //         5,
  //         position,
  //         imgWidth,
  //         imgHeight,
  //         undefined,
  //         'FAST'
  //       );
  //       addExtraDataToPage(pdfData, '', position + imgHeight + 10);
  //       heightLeft -= pageHeight;
  //       while (heightLeft >= 0) {
  //         position = heightLeft - imgHeight;
  //         addPageAndImage(pdfData, contentDataURL, position);
  //         addExtraDataToPage(pdfData, '', position + imgHeight + 10);
  //         heightLeft -= pageHeight;
  //       }
  //       return pdfData;
  //     })
  //     .then((doc) => {
  //       doc.save('sellingreport.pdf');
  //       // popupWin.close();
  //       this.shared.spinner.hide();
  //     });
  // }
  // sendEmailData(Email: any, notes: any, from: any) {
  //   this.shared.spinner.show();
  //   const printContents = document.getElementById('SellingGross')!.innerHTML;
  //   const iframe = document.createElement('iframe');
  //   // Make the iframe invisible
  //   iframe.style.position = 'absolute';
  //   iframe.style.width = '0px';
  //   iframe.style.height = '0px';
  //   iframe.style.border = 'none';
  //   document.body.appendChild(iframe);
  //   const doc = iframe.contentWindow?.document;
  //   if (!doc) {
  //     console.error('Failed to create iframe document');
  //     return;
  //   }
  //   doc.open();
  //   doc.write(`
  //         <html>
  //             <head>
  //                 <title>F & I Summary</title>
  //                 <style>
  //                 @font-face {
  //                   font-family: 'GothamBookRegular';
  //                   src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
  //                        url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
  //                 }
  //                 @font-face {
  //                   font-family: 'Roboto';
  //                   src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
  //                        url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
  //                 }
  //                 @font-face {
  //                   font-family: 'RobotoBold';
  //                   src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
  //                        url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
  //                 }
  //                 .negative {
  //                   color: red;
  //                 }
  //                 .HideItems{
  //                   display:none;
  //                  }            
  // .backgrossbtn {
  //                   float: left;
  //                   font-size: 11px;
  //                   background-color: #2d91f1;
  //                   color: white;
  //                   padding: 4px;
  //                   border: 1px solid #2d91f1;
  //                   cursor: pointer;
  //                   border-radius: 3px;
  //               }
  //                 .Selectedrow {
  //                   background-color: #d3ecff !important;
  //                   color: #000 !important;
  //               }
  //                 .justify-content-between {
  //                   justify-content: space-between !important;
  //               }
  //               .d-flex {
  //                   display: flex !important;
  //               }
  //                 .bg-white {
  //                   background: #ffffff !important;
  //               }
  //                 .negative {
  //                   color: red;
  //                 }
  //                 .negative {
  //                   color: red;
  //                 }
  //                 .performance-scorecard .table>:not(:first-child) {
  //                   border-top: 0px solid #ffa51a
  //                 }
  //                 .performance-scorecard .table {
  //                   text-align: center;
  //                   text-transform: capitalize;
  //                   border: transparent;
  //                   width: 100%;
  //                 }
  //                 .performance-scorecard .table th,
  //                 .performance-scorecard .table td{
  //                   white-space: nowrap;
  //                   vertical-align: top;
  //                 }
  //                 .performance-scorecard .table th:first-child,
  //                 .performance-scorecard .table td:first-child {
  //                   position: sticky;
  //                   left: 0;
  //                   z-index: 1;
  //                   background-color: #337ab7;
  //                 }
  //                 .performance-scorecard .table tr:nth-child(odd) td:first-child,
  //                 .performance-scorecard .table tr:nth-child(odd) td:nth-child(2) {
  //                   // background-color: #e9ecef;
  //                   background-color: #ffffff;
  //                 }
  //                 .performance-scorecard .table tr:nth-child(even) td:first-child,
  //                 .performance-scorecard .table tr:nth-child(even) td:nth-child(2) {
  //                   background-color: #ffffff;
  //                 }
  //                 .performance-scorecard .table tr:nth-child(odd) {
  //                   // background-color: #e9ecef;
  //                   background-color: #ffffff;
  //                 }
  //                 .performance-scorecard .table tr:nth-child(even) {
  //                   background-color: #ffffff;
  //                 }
  //                 .performance-scorecard .table .spacer {
  //                   // width: 50px !important;
  //                   background-color: #cfd6de !important;
  //                   border-left: 1px solid #cfd6de !important;
  //                   border-bottom: 1px solid #cfd6de !important;
  //                   border-top: 1px solid #cfd6de !important;
  //                  }
  //                 .performance-scorecard .table .hidden {
  //                   display: none !important;
  //                 }
  //                 .performance-scorecard .table .bdr-rt {
  //                   border-right: 1px solid #abd0ec;
  //                 }
  //                 .performance-scorecard .table thead {
  //                   position: sticky;
  //                   top: 0;
  //                   z-index: 99;
  //                   font-family: 'FaktPro-Bold';
  //                   font-size: 0.8rem;
  //                 }
  //                 .performance-scorecard .table thead th {
  //                   padding: 5px 10px;
  //                   margin: 0px;
  //                 }
  //                 .performance-scorecard .table thead .bdr-btm {
  //                   border-bottom: #005fa3;
  //                 }
  //                 .performance-scorecard .table thead tr:nth-child(1) {
  //                   background-color: #fff !important;
  //                   color: #000;
  //                   text-transform: uppercase;
  //                   border-bottom: #cfd6de;
  //                 }
  //                 .performance-scorecard .table thead tr:nth-child(2) {
  //                   background-color: #fff !important;
  //                   color: #000;
  //                   text-transform: uppercase;
  //                   border-bottom: #cfd6de;
  //                   box-shadow: inset 0 1px 0 0 #cfd6de;
  //                 }
  //                 .performance-scorecard .table thead tr:nth-child(2) th :nth-child(1) {
  //                   background-color: #337ab7 !important;
  //                   color: #fff;
  //                 }
  //                 .performance-scorecard .table thead tr:nth-child(3) {
  //                   background-color: #fff !important;
  //                   color: #000;
  //                   text-transform: uppercase;
  //                   border-bottom: #cfd6de;
  //                   box-shadow: inset 0 1px 0 0 #cfd6de;
  //                 }
  //                 .performance-scorecard .table thead tr:nth-child(3) th :nth-child(1) {
  //                   background-color: #337ab7 !important;
  //                   color: #fff;
  //                 }
  //                 .performance-scorecard .table tbody {
  //                   font-family: 'FaktPro-Normal';
  //                   font-size: .9rem;
  //                 }
  //                 .performance-scorecard .table tbody td {
  //                   padding: 2px 10px;
  //                   margin: 0px;
  //                   border: 1px solid #cfd6de
  //                 }
  //                 .performance-scorecard .table tbody tr {
  //                   border-bottom: 1px solid #37a6f8;
  //                   border-left: 1px solid #37a6f8
  //                 }
  //                 .performance-scorecard .table tbody td:first-child {
  //                   text-align: start;
  //                   box-shadow: inset -1px 0 0 0 #cfd6de;
  //                 }
  //                 .performance-scorecard .table tbody tr td:not(:first-child) {
  //                   text-align: right !important;
  //                 }
  //                 .performance-scorecard .table tbody .sub-title {
  //                   font-size: .8rem !important;
  //                 }
  //                 .performance-scorecard .table tbody .sub-subtitle{
  //                   font-size: .7rem !important;
  //                 }
  //                 .performance-scorecard .table tbody .text-bold {
  //                   font-family: 'FaktPro-Bold';
  //                 }
  //                 .performance-scorecard .table tbody .darkred-bg {
  //                   background-color: #282828 !important;
  //                   color: #fff;
  //                 }
  //                 .performance-scorecard .table tbody .lightblue-bg {
  //                   background-color: #646e7a !important;
  //                   color: #fff;
  //                 }
  //                 .performance-scorecard .table tbody .gold-bg {
  //                   background-color: #ffa51a;
  //                   color: #fff;
  //                 }
  //                 </style>
  //             </head>
  //             <body id='content'>
  //                 ${printContents}
  //             </body>
  //         </html>
  //     `);
  //   doc.close();
  //   const div = doc.getElementById('content');
  //   if (!div) {
  //     console.error('Element not found');
  //     return;
  //   }
  //   const options = {
  //     logging: true,
  //     allowTaint: false,
  //     useCORS: true,
  //   };
  //   html2canvas(div, options)
  //     .then((canvas) => {
  //       let imgWidth = 285;
  //       let pageHeight = 206;
  //       let imgHeight = (canvas.height * imgWidth) / canvas.width;
  //       let heightLeft = imgHeight;
  //       const contentDataURL = canvas.toDataURL('image/png');
  //       let pdfData = new jsPDF('l', 'mm', 'a4', true);
  //       let position = 5;
  //       pdfData.addImage(
  //         contentDataURL,
  //         'PNG',
  //         5,
  //         position,
  //         imgWidth,
  //         imgHeight,
  //         undefined,
  //         'FAST'
  //       );
  //       heightLeft -= pageHeight;
  //       while (heightLeft >= 0) {
  //         position = heightLeft - imgHeight;
  //         pdfData.addPage();
  //         pdfData.addImage(
  //           contentDataURL,
  //           'PNG',
  //           5,
  //           position,
  //           imgWidth,
  //           imgHeight,
  //           undefined,
  //           'FAST'
  //         );
  //         heightLeft -= pageHeight;
  //       }
  //       const pdfBlob = pdfData.output('blob');
  //       const pdfFile = this.blobToFile(pdfBlob, 'SalesGross.pdf');
  //       const formData = new FormData();
  //       formData.append('to_email', Email);
  //       formData.append('subject', 'F & I Summary');
  //       formData.append('file', pdfFile);
  //       formData.append('notes', notes);
  //       formData.append('from', from);
  //       this.shared.api.postmethod(this.comm.routeEndpoint + 'mail', formData).subscribe(
  //         (res: any) => {
  //           console.log('Response:', res);
  //           if (res.status === 200) {
  //             // alert(res.response);
  //             this.toast.show(res.response,'success','Success');
  //           } else {
  //             alert('Invalid Details');
  //           }
  //         },
  //         (error) => {
  //           console.error('Error:', error);
  //         }
  //       );
  //     })
  //     .catch((error) => {
  //       console.error('html2canvas error:', error);
  //     })
  //     .finally(() => {
  //       this.shared.spinner.hide();
  //       // popupWin.close();
  //     });
  // }
  // public blobToFile = (theBlob: Blob, fileName: string): File => {
  //   return new File([theBlob], fileName, {
  //     lastModified: new Date().getTime(),
  //     type: theBlob.type,
  //   });
  // };
}