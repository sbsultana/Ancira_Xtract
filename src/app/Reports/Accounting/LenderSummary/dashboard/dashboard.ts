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
import { NgbActiveModal, NgbModal, NgbModalRef } from '@ng-bootstrap/ng-bootstrap';
import { Title } from '@angular/platform-browser';
import { CommonModule, CurrencyPipe, DatePipe, NgStyle } from '@angular/common';
import * as FileSaver from 'file-saver';
import * as ExcelJS from 'exceljs';
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
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  imports: [CommonModule, SharedModule, BsDatepickerModule, Stores, FormsModule, DateRangePicker, ReactiveFormsModule],
  standalone: true,
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  LendersData: any = []
  LendersDataDetails: any = []
  LendersIndividual: any = [];
  LendersTotal: any = []
  DupFromDate: any = '';
  DupToDate: any = ''
  Lenders: any = []
  FromDate: any = '';
  ToDate: any = '';
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';
  NoData: boolean = false;

  //  storeIds: any = []
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };

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


  storeIds!: any;
  dateType: any = 'MTD';
  groups: any = 0;

  CompleteComponentState: boolean = true;
  solutionurl: any;
  LogCount = 1;
  header: any = [
    {
      type: 'Bar',
      storeIds: this.storeIds,
      groups: this.groups,
    },
  ];
  popup: any = [{ type: 'Popup' }];

  constructor(public shared: Sharedservice, public setdates: Setdates, private comm: common, private toast: ToastService,
    public spinner: NgxSpinnerService, private Api: Api, private ngbmodal: NgbModal, private datepipe: DatePipe
  ) {
    this.solutionurl = this.shared.api;
    this.shared.setTitle(this.comm.titleName + '-Lender Summary');
    if (typeof window !== 'undefined') {
      if (localStorage.getItem('flag') == 'V') {
        this.storeIds = [];
        console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.groupid
        JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store.indexOf(',') > 0 ?
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store.split(',') :
          this.storeIds.push(JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store)
        localStorage.setItem('flag', 'M')
      } else {
        if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
          this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
          this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
        }
      }
      if (this.shared.common.groupsandstores.length > 0) {
        this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
        this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
        this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
        this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
        // // console.log(this.stores, this.groupsArray, 'Stores and Groups');
        this.getStoresandGroupsValues()
        // this.StoresData(this.ngChanges)
      }

      if (localStorage.getItem('stime') != null) {
        let stime = localStorage.getItem('stime');
        if (stime != null && stime != '') {
          this.initializeDates(stime)
          this.DateType = stime
        }
      } else {
        this.initializeDates('MTD')
        this.DateType = 'MTD'
      }
      localStorage.setItem('stime', 'MTD')


      this.shared.setTitle(this.shared.common.titleName + '-Lender Summary');
      // if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: 'Lender Summary',
        stores: this.storeIds,
        datetype: this.DateType,
        fromdate: this.FromDate,
        todate: this.ToDate,
        groups: this.groups,
        count: 0,
      };
      this.shared.api.SetHeaderData({
        obj: data,
      });
      this.header = [
        {
          type: 'Bar',
          storeIds: this.storeIds,
          fromdate: this.FromDate,
          todate: this.ToDate,
          groups: this.groups,
        },
      ];
      this.setDates(this.DateType);
      this.GetLenders();
      // }
    }
  }

  ngOnInit(): void { }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
  }


  openDetails(Item: any, ParentItem: any, cat: any) { }
  Favreports: any = [];

  StoresData(data: any) {

    console.log(data, 'Data');

    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
  }

  isDesc: boolean = false;
  column: string = '';
  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // // console.log(property)
    this.LendersData.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
  GetLenders() {
    console.log(this.FromDate, this.ToDate);
    this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    this.LendersData = [];
    this.NoData = false;
    this.spinner.show();
    const obj = {
      "fromDate": this.FromDate,
      "toDate": this.ToDate,
      'Stores': this.storeIds.toString()
    };
    this.Api.postmethod(this.comm.routeEndpoint + 'LenderSummary', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          this.LendersData = [];
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.LendersTotal = res.response.filter(
                (e: any) => e.Location == 'REPORT TOTAL'
              );
              this.LendersIndividual = res.response.filter(
                (i: any) => i.Location != 'REPORT TOTAL'
              );
              this.NoData = false;
              this.LendersIndividual.some(function (x: any) {
                if (
                  x.Data != undefined &&
                  x.Data != '' &&
                  x.Data != null
                ) {
                  x.Data = JSON.parse(x.Data);
                  x.Data = x.Data.reduce((r: any, { SaleType }: any) => {
                    if (!r.some((o: any) => o.SaleType == SaleType)) {
                      r.push({
                        SaleType,
                        subData: x.Data.filter((v: any) => v.SaleType == SaleType),
                      });
                    }
                    return r;
                  }, []);
                }
              });
              this.LendersIndividual.push(this.LendersTotal[0]);
              this.LendersData = this.LendersIndividual;
              console.log(this.LendersData);
              this.spinner.hide();
            } else {
              this.spinner.hide();
              this.NoData = true;
            }
          } else {
            this.spinner.hide();
            this.NoData = true;
          }
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
  popupReference!: NgbModalRef;
  parentItem: any = []
  childItem: any = [];
  reference: any = ''
  lenderpopup(tmp: any, parent: any, child: any, ref: any) {
    this.parentItem = parent;
    this.childItem = child
    this.reference = ref
    console.log(this.parentItem, this.childItem, this.reference)
    this.popupReference = this.ngbmodal.open(tmp, { size: 'xl', backdrop: 'static', keyboard: true, centered: true, modalDialogClass: 'custom-modal' })
    this.GetLendersDetails()
  }
  loader: boolean = false
  GetLendersDetails() {
    this.loader = true
    this.LendersDataDetails = []
    const obj = {
      "fromDate": this.FromDate,
      "toDate": this.ToDate,
      'Stores': this.parentItem.StoreID,
      "FinSource": this.childItem['Finance Source']
    };
    this.Api.postmethod(this.comm.routeEndpoint + 'GetLenderSummaryDetails', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          this.loader = false;
          this.LendersDataDetails = [];
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.LendersDataDetails = res.response
              console.log(this.LendersDataDetails);
              this.spinner.hide();
            } else {
              this.spinner.hide();
            }
          } else {
            this.spinner.hide();
          }
        } else {
          this.loader = false;
          this.toast.show(res.status, 'danger', 'Error');
          this.spinner.hide();
        }
      },
      (error) => {
        this.toast.show('502 Bad Gate Way Error', 'danger', 'Error');
        this.spinner.hide();
      }
    );
  }
  closePopup() {
    if (this.popupReference) {
      this.popupReference.close();
    }
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }


  SPRstate: any;
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Lender Summary') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          // this.groupId = this.ngChanges.groups;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          // this.storeIds = this.ngChanges.storeIds;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          // // console.log(this.stores, this.groupsArray, 'Stores and Groups');
          this.getStoresandGroupsValues()
        }
      }
    })
    this.shared.api.GetReportOpening().subscribe((res) => {
      // // console.log(res);
      if (res.obj.Module == 'Lender Summary') {
        document.getElementById('report')?.click();
      }
    });
    this.shared.api.GetReports().subscribe((data) => {
      if (data.obj.Reference == 'Lender Summary') {
        if (data.obj.header == undefined) {
          this.storeIds = data.obj.storeValues;
          this.groups = data.obj.groups;
          if (data.obj.FromDate != undefined && data.obj.ToDate != undefined) {
            this.FromDate = data.obj.FromDate;
            this.ToDate = data.obj.ToDate;
            this.storeIds = data.obj.storeValues;
            this.dateType = data.obj.dateType;
            this.GetLenders()
          } else {
            this.FromDate = data.obj.FromDate;
            this.ToDate = data.obj.ToDate;
            this.storeIds = data.obj.storeValues;
            this.dateType = data.obj.dateType;
            this.GetLenders()
          }
        } else {
          if (data.obj.header == 'Yes') {
            this.storeIds = '1';
            // console.log(this.storeIds);
            this.GetLenders()
          }
        }
        const headerdata = {
          title: 'Lender Summary',

          stores: this.storeIds,
          datetype: this.dateType,
          fromdate: this.FromDate,
          todate: this.ToDate,
          groups: this.groups,
        };
        this.shared.api.SetHeaderData({
          obj: headerdata,
        });
        this.header = [
          {
            type: 'Bar',
            storeIds: this.storeIds,
            fromdate: this.FromDate,
            todate: this.ToDate,
            groups: this.groups,
          },
        ];
      }
    });
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Lender Summary') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Lender Summary') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Lender Summary') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Lender Summary') return;
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




  openCustomPicker(datepicker: any) {
    this.DateType = 'C';
    localStorage.setItem('time', this.DateType);
    datepicker.toggle();
  }


  dateRangeCreated($event: any) {
    if ($event) {
      this.FromDate = this.shared.datePipe.transform($event[0], 'MM-dd-yyyy');
      this.ToDate = this.shared.datePipe.transform($event[1], 'MM-dd-yyyy');
      this.bsRangeValue = [this.FromDate, this.ToDate];
      if (this.DateType === 'C') this.custom = true;
    }
  }

  updatedDates(data: any) {
    // console.log(data);
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
  }




  // comments code

  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0;
  @ViewChild('scrollcent') scrollcent!: ElementRef;
  updateVerticalScroll(event: any): void {
    this.scrollCurrentposition = event.target.scrollTop;
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop / (event.target.scrollHeight - scrollDemo.clientHeight)) * 100
    );
  }






  //////////////////////  REPORT CODE /////////////////////////////////////////////
  // === State used in HTML ===
  Bar: boolean = true;
  activePopover: number = -1;

  custom: boolean = false;
  // FromDate: any;
  // ToDate: any;
  bsRangeValue!: Date[];



  // stores: any[] = [];
  selectedstorevalues: any[] = [];
  storeName: string = '';
  // groupName: string = '';
  // groups: any[] = [];
  selectedGroups: any[] = [];
  AllStores: boolean = true;
  AllGroups: boolean = true;

  storeorgroup: any = ['G'];
  // retailorlease: any = [];
  // saleType: any = 'Retail,Lease,Wholesale,Misc,Fleet,Demo,Special Order,Rental,Dealer Trade';

  // dealStatus: any = ['Booked', 'Finalized', 'Delivered'];

  // constructor(private pipe: DatePipe) {}

  // === HTML Interactions ===
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }

  // Store selection
  allstores() {
    this.AllStores = !this.AllStores;
    this.selectedstorevalues = this.AllStores ? this.stores.map((s: { ID: any; }) => s.ID) : [];
  }


  individualStores(e: any) {

    const index = this.selectedstorevalues.findIndex((i: any) => i == e.ID);
    if (index >= 0) {
      this.selectedstorevalues.splice(index, 1);
      this.AllStores = false;
      if (this.selectedstorevalues.length == 1) {
        this.storeorgroup = 'S'
      }
      else {
        this.storeorgroup = 'G'
      }

    } else {
      this.selectedstorevalues.push(e.ID);
      if (this.selectedstorevalues.length == 1) {
        this.storeorgroup = 'S'
      }
      else {
        this.storeorgroup = 'G'
      }
      if (this.selectedstorevalues.length == this.stores.length) {
        this.AllStores = true;
      } else {
        this.AllStores = false;
      }
    }
    if (this.selectedstorevalues.length == 1) {
      this.storeName = this.stores.filter((val: any) => val.ID == this.selectedstorevalues.toString())[0].storename
    }
  }
  // Group selection
  allgroups() {
    this.AllGroups = !this.AllGroups;
    this.selectedGroups = this.AllGroups ? this.groups.map((g: { sg_id: any; }) => g.sg_id) : [];
  }

  individualgroups(g: any) {
    this.selectedGroups = [g.sg_id];
    this.groupName = g.sg_name;
    this.storeorgroup = g.sg_id === 1 ? 'S' : 'G';
  }




  storeorgroups(_block: any, val: string) {
    this.storeorgroup = [val];
  }


  setDates(type: any) {
    // localStorage.setItem('time', type);
    // this.datevaluetype=
    // console.log(type);

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
  // dateRangeCreated($event: any) {
  //   if ($event) {
  //     this.FromDate = this.shared.datePipe.transform($event[0], 'MM-dd-yyyy');
  //     this.ToDate = this.shared.datePipe.transform($event[1], 'MM-dd-yyyy');
  //     if (this.DateType === 'C') this.custom = true;
  //   }
  // }

  close() {
    this.shared.ngbmodal.dismissAll();
  }

  // Final Apply
  viewreport() {
    this.activePopover = -1;
    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      return;
    }
    else {
      const data = {
        Reference: 'Lender Summary',
        FromDate: this.FromDate,
        ToDate: this.ToDate,
        storeValues: this.storeIds.toString(),
        dateType: this.DateType,
        groups: this.selectedGroups.toString(),
      };
      this.shared.api.SetReports({
        obj: data,
      });
      this.close();
      this.GetLenders()
    }
  }

  generatePDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save('Lender Summary.pdf'); // ✅ only here
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
      const pdfFile = this.blobToFile(pdfBlob, 'Lender Summary.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Lender Summary');
      formData.append('file', pdfFile);
      formData.append('notes', notes);
      formData.append('from', from);

      this.Api.postmethod(this.comm.routeEndpoint + 'mail', formData)
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
    doc.text('Lender Summary', 14, 12);

    /* ================= DATES ================= */
    const monthText = this.DupFromDate
      ? this.datepipe.transform(this.DupFromDate, 'MMMM yyyy')
      : '';

    const fromDate = this.DupFromDate
      ? this.datepipe.transform(this.DupFromDate, 'MM.dd.yyyy')
      : '';

    const toDate = this.DupToDate
      ? this.datepipe.transform(this.DupToDate, 'MM.dd.yyyy')
      : '';

    /* ================= HELPER ================= */
    const formatValue = (val: any) => {
      if (val === null || val === undefined || val === '' || Number(val) === 0) {
        return '-';
      }
      return Number(val);
    };

    /* ================= HEADERS (2 ROWS) ================= */
    const headers = [
      [
        monthText || '',
        '',
        '',
        '',
        '',
        ''
      ],
      [
        `${fromDate} - ${toDate}`,
        'Deal Count',
        'Contract Value',
        'Profit on Financing',
        'Front Gross',
        'Back Gross'
      ]
    ];

    /* ================= BODY ================= */
    const body: any[] = [];

    this.LendersData.forEach((lvl1: any) => {

      const isTotal = lvl1.Location === 'REPORT TOTAL';

      /* ===== LEVEL 1 ===== */
      body.push([
        lvl1.Location || '-',
        formatValue(lvl1.DealCount),
        formatValue(lvl1.ContractValue),
        formatValue(lvl1.ProfitonFinancing),
        formatValue(lvl1.FrontGross),
        formatValue(lvl1.BackGross)
      ]);

      /* ===== LEVEL 2 ===== */
      if (Array.isArray(lvl1.Data)) {
        lvl1.Data.forEach((lvl2: any) => {

          body.push([
            '   ' + (lvl2.SaleType || '-'),
            '',
            '',
            '',
            '',
            ''
          ]);

          /* ===== LEVEL 3 ===== */
          if (Array.isArray(lvl2.subData)) {
            lvl2.subData.forEach((lvl3: any) => {

              body.push([
                '      ' + (lvl3['Finance Source'] || lvl3.FIN || '-'),
                formatValue(lvl3.DealCount),
                formatValue(lvl3.ContractValue),
                formatValue(lvl3.ProfitonFinancing),
                formatValue(lvl3.FrontGross),
                formatValue(lvl3.BackGross)
              ]);

            });
          }

        });
      }

    });

    /* ================= TABLE ================= */
    autoTable(doc, {
      startY: 18,
      head: headers,
      body: body,
      theme: 'grid',

      styles: {
        fontSize: 8,
        cellPadding: 2,
        lineWidth: 0.1,
        textColor: [20, 20, 20],
        valign: 'middle'
      },

      headStyles: {
        textColor: [255, 255, 255],
        halign: 'center',
        fontStyle: 'bold'
      },

      didParseCell: (data: any) => {

        const col = data.column.index;

        /* ===== HEADER COLORS (2 ROWS) ===== */
        if (data.section === 'head') {

          // Row 1 → Dark Blue
          if (data.row.index === 0) {
            data.cell.styles.fillColor = [5, 84, 239]; // #0554EF
          }

          // Row 2 → Light Blue
          if (data.row.index === 1) {
            data.cell.styles.fillColor = [69, 132, 255]; // #4584FF
          }

          // Align first column left
          if (col === 0) {
            data.cell.styles.halign = 'left';
          }

          return;
        }

        /* ===== ALIGNMENT ===== */
        if (col === 0) {
          data.cell.styles.halign = 'left';
        } else {
          data.cell.styles.halign = 'right';
        }

        /* ===== EMPTY VALUES ===== */
        if (data.cell.raw === '-' || data.cell.raw === null || data.cell.raw === undefined) {
          data.cell.text = ['-'];
          data.cell.styles.halign = 'center';
          return;
        }

        /* ===== NUMBER FORMATTING ===== */
        if (typeof data.cell.raw === 'number') {

          const val = Math.round(data.cell.raw);

          // Deal Count
          if (col === 1) {
            data.cell.text = [val.toLocaleString()];
          }

          // Currency columns
          if ([2, 3, 4, 5].includes(col)) {
            data.cell.text = [`$ ${val.toLocaleString()}`];

            if (val < 0) {
              data.cell.styles.textColor = [255, 0, 0];
            }
          }
        }

        /* ===== ALT ROW COLOR ===== */
        if (data.section === 'body' && data.row.index % 2 === 0) {
          data.cell.styles.fillColor = [245, 247, 250];
        }

        /* ===== LEVEL 1 COLOR ===== */
        if (data.row.raw[0] && !data.row.raw[0].startsWith(' ')) {
          data.cell.styles.fillColor = [217, 231, 255]; // Excel light blue
        }

        /* ===== TOTAL ROW ===== */
        if (data.row.raw[0] === 'REPORT TOTAL') {
          data.cell.styles.fillColor = [141, 180, 255];
          data.cell.styles.fontStyle = 'bold';
        }
      }
    });

    return doc;
  }
  // Excel Filter Headers

  getSelectedStoreNames(): string {
    if (!this.storeIds || this.storeIds.length === 0) return '';

    const ids = this.storeIds.toString().split(',');

    const selectedStores = this.stores.filter((s: any) =>
      ids.includes(s.ID.toString())
    );

    return selectedStores.map((s: any) => s.storename).join(', ');
  }
  getReportFilters(): { title: string; filters: any[] } {
    return {
      title: 'Lender Summary',
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
          label: 'Timeframe',
          value: this.DateType === 'C'
            ? `${this.FromDate} to ${this.ToDate}`   // Custom range
            : `${this.displaytime} (${this.FromDate} to ${this.ToDate})` // MTD/YTD etc
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
  exportToExcel(): void {

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Lender Summary');

    /* ================= HELPER ================= */
    const formatValue = (val: any) => {
      if (val === null || val === undefined || val === '' || Number(val) === 0) {
        return '-';
      }
      return Number(val);
    };

    /* ================= FILTERS ================= */
    const filterRowCount = this.addExcelFiltersSection(worksheet);
    const headerStartRow = filterRowCount + 1;

    /* ================= DATES ================= */
    const monthText = this.DupFromDate
      ? this.datepipe.transform(this.DupFromDate, 'MMMM yyyy')
      : '';

    const fromDate = this.DupFromDate
      ? this.datepipe.transform(this.DupFromDate, 'MM.dd.yyyy')
      : '';

    const toDate = this.DupToDate
      ? this.datepipe.transform(this.DupToDate, 'MM.dd.yyyy')
      : '';

    /* ================= HEADER ================= */
    worksheet.addRow([monthText, '', '', '', '', '']);

    worksheet.addRow([
      `${fromDate} - ${toDate}`,
      'Deal Count',
      'Contract Value',
      'Profit on Financing',
      'Front Gross',
      'Back Gross'
    ]);

    /* ================= HEADER STYLE ================= */
    [headerStartRow, headerStartRow + 1].forEach(r => {
      worksheet.getRow(r).eachCell((cell: any) => {

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
        cell.border = this.getBorder();
      });
    });

    /* ================= FORMAT FUNCTION ================= */
    const formatRow = (row: any, isTotal = false, indent = 0) => {

      row.eachCell((cell: any, colNumber: number) => {

        cell.border = this.getBorder();
        cell.font = { name: 'Calibri', bold: isTotal };

        if (colNumber === 1) {
          cell.alignment = {
            horizontal: 'left',
            vertical: 'middle',
            indent: indent
          };
          return;
        }

        if (cell.value === '-' || cell.value === null || cell.value === undefined) {
          cell.alignment = { horizontal: 'center' };
          return;
        }

        let num = Number(cell.value);

        if (!isNaN(num)) {

          // Deal Count
          if (colNumber === 2) {
            cell.numFmt = '#,##0';
          }

          // Currency columns
          if ([3, 4, 5, 6].includes(colNumber)) {
            cell.numFmt = '"$" * #,##0; "$" * -#,##0';
          }

          if (num < 0) {
            cell.font = { color: { argb: 'FF0000' }, name: 'Calibri' };
          }
        }

        cell.alignment = { horizontal: 'right' };
      });

      if (isTotal) {
        row.eachCell((cell: any) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '8DB4FF' }
          };
        });
      }
    };

    /* ================= DATA ================= */
    let altIndex = 0;

    this.LendersData.forEach((lvl1: any) => {

      const isTotal = lvl1.Location === 'REPORT TOTAL';

      /* ===== LEVEL 1 ===== */
      const row1 = worksheet.addRow([
        lvl1.Location || '-',
        formatValue(lvl1.DealCount),
        formatValue(lvl1.ContractValue),
        formatValue(lvl1.ProfitonFinancing),
        formatValue(lvl1.FrontGross),
        formatValue(lvl1.BackGross)
      ]);

      formatRow(row1, isTotal, 0);

      row1.eachCell((cell: any) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'D9E7FF' }
        };
      });

      /* ===== LEVEL 2 ===== */
      if (Array.isArray(lvl1.Data)) {
        lvl1.Data.forEach((lvl2: any) => {

          altIndex++;

          const row2 = worksheet.addRow([
            lvl2.SaleType || '-',
            '',
            '',
            '',
            '',
            ''
          ]);

          formatRow(row2, false, 1);

          if (altIndex % 2 === 0) {
            row2.eachCell((cell: any) => {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'F5F9FF' }
              };
            });
          }

          /* ===== LEVEL 3 ===== */
          if (Array.isArray(lvl2.subData)) {
            lvl2.subData.forEach((lvl3: any) => {

              altIndex++;

              const row3 = worksheet.addRow([
                lvl3['Finance Source'] || lvl3.FIN || '-',
                formatValue(lvl3.DealCount),
                formatValue(lvl3.ContractValue),
                formatValue(lvl3.ProfitonFinancing),
                formatValue(lvl3.FrontGross),
                formatValue(lvl3.BackGross)
              ]);

              formatRow(row3, false, 2);

              if (altIndex % 2 === 0) {
                row3.eachCell((cell: any) => {
                  cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'F5F9FF' }
                  };
                });
              }

            });
          }

        });
      }

    });

    /* ================= FINAL SETTINGS ================= */
    worksheet.views = [
      { state: 'frozen', xSplit: 1, ySplit: headerStartRow + 1 }
    ];

    worksheet.columns = [
      { width: 45 },
      { width: 15 },
      { width: 20 },
      { width: 22 },
      { width: 18 },
      { width: 18 }
    ];

    /* ================= DOWNLOAD ================= */
    workbook.xlsx.writeBuffer().then((data: any) => {

      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });

      FileSaver.saveAs(blob, 'Lender Summary_Report.xlsx');
    });
  }
  getBorder() {
    return {
      top: { style: 'thin' as 'thin' },
      left: { style: 'thin' as 'thin' },
      bottom: { style: 'thin' as 'thin' },
      right: { style: 'thin' as 'thin' }
    };
  }
}