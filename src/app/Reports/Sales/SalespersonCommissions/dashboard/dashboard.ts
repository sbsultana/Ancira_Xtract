import { Component, } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Subscription } from 'rxjs';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';

import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';


@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores, DateRangePicker],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  Appointmentdata: any = [];
  NoData: any = false;

  SalespersonCommData: any = []

  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;


  FromDate: any = '';
  ToDate: any = '';
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'C';
  displaytime: any = '';
  DupFromDate: any = '';
  DupToDate: any = ''
  bsRangeValue!: Date[];

  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    Types: []
  }

  storeIds: any = '0';
  groups: any = 1;
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  stores: any = [];


  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'Y',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  constructor(public shared: Sharedservice, public setdates: Setdates, private toast: ToastService,
  ) {
    localStorage.setItem('time', 'C');
    this.shared.setTitle(this.shared.common.titleName + '- Salesperson Commissions');
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',');
      let today = new Date();
      const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
      const tomorrow = new Date();
      tomorrow.setDate(today.getDate() + 1);
      this.FromDate = this.shared.datePipe.transform(monthStart, 'MM-dd-yyyy');
      this.ToDate = this.shared.datePipe.transform(tomorrow, 'MM-dd-yyyy');
      this.setDates(this.DateType)
      this.setHeaderData();
      this.GetSalespersonCommissions()
    }
  }

  setHeaderData() {
    const data = {
      title: 'Salesperson Commissions',
      stores: this.storeIds,
      groups: this.groupId,
      fromdate: this.FromDate,
      todate: this.ToDate,
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }

  isDesc: boolean = false;
  column: string = 'CategoryName';

  datetype() {
    if (this.DateType == 'PM') {
      return 'SP';
    }
    else if (this.DateType == 'C') {
      return 'C'
    }
    return this.DateType;
  }

  GetSalespersonCommissions() {
    this.shared.spinner.show();
    let obj = {
      fromdate: this.shared.datePipe.transform(this.FromDate, 'MM-dd-yyyy'),
      todate: this.shared.datePipe.transform(this.ToDate, 'MM-dd-yyyy'),
      store: this.storeIds.toString(),
    };
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetSalesPersonCommCalc', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          this.SalespersonCommData = res.response;
          this.shared.spinner.hide();
          this.NoData = true;
          if (this.SalespersonCommData.length > 0) {
            this.NoData = false;
          } else {
            this.NoData = true;
          }
        } else {
          this.shared.spinner.hide();
          this.toast.show('Invalid Details', 'danger', 'Error');
        }
      },
      (error) => { });
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  ngAfterViewInit() {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Salesperson Commissions') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (this.email != undefined) {
        if (res.obj.title == 'Salesperson Commissions') {
          if (res.obj.stateEmailPdf == true) {
            //    this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
        }
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (this.print != undefined) {
        if (res.obj.title == 'Salesperson Commissions') {
          if (res.obj.statePrint == true) {
            //    this.GetPrintData();
          }
        }
      }
    });
    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'Salesperson Commissions') {
          if (res.obj.statePDF == true) {
            //   this.generatePDF();
          }
        }
      }
    });
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res) => {
      if (this.excel != undefined) {
        if (res.obj.title == 'Salesperson Commissions') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });
  }

  ngOnDestroy() {
    if (this.excel != undefined) {
      this.excel.unsubscribe()
    }
    if (this.Pdf != undefined) {
      this.Pdf.unsubscribe()
    }
    if (this.print != undefined) {
      this.print.unsubscribe()
    }
    if (this.email != undefined) {
      this.email.unsubscribe()
    }
  }

  updatedDates(data: any) {
    console.log(data);
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
    this.bsRangeValue = [new Date(this.FromDate), new Date(this.ToDate)];

  }

  setDates(type: any) {
    this.DateType == 'C' ? this.displaytime = ' (  ' + this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy') + ' - ' + this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy') + ' ) ' :
      this.displaytime = 'Time Frame (' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ')';
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
    console.log(this.FromDate, this.ToDate, this.DateType, this.displaytime, '..............');
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

  openMenu() {
    // const DetailsSPC = this.ngbmodal.open(
    //   SalespersoncommissionsDetailsComponent,
    //   {
    //     size: 'xl',
    //     backdrop: 'static',
    //   }
    // );
    // DetailsSPC.componentInstance.SPC_Details = {
    //   Salesperson: Saleperson,
    //   FromDate: FromDate,
    //   ToDate: ToDate,
    //   StoreName: this.StoreValues.toString(),
    // };
    // //console.log(DetailsSPC);
  }

  activePopover: number = -1;
  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }
  viewreport() {
    this.activePopover = -1
    this.setHeaderData();
    this.GetSalespersonCommissions()
  }
  ExcelStoreNames: any = [];

  exportToExcel() { }


}