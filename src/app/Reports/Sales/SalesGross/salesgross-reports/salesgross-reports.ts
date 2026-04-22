import { Component, HostListener, Input, SimpleChanges } from '@angular/core';

import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Stores } from '../../../../CommonFilters/stores/stores';

import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { ToastContainer } from '../../../../Layout/toast-container/toast-container';
import { NgbModule } from '@ng-bootstrap/ng-bootstrap';
@Component({
  selector: 'app-salesgross-reports',
  imports: [SharedModule, DateRangePicker, Stores, ToastContainer,NgbModule],
  templateUrl: './salesgross-reports.html',
  styleUrl: './salesgross-reports.scss'
})
export class SalesgrossReports {

  ngChanges: any = []

  // Report Popup
  @Input() header: any;
  groupId: any = [1];
  storeIds: any = [1]
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  otherStoresArray: any = [];
  otherStoreIds: any = [];
  DefaultLoad: any = ''

  groupName: any = '';
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores,
    'storeids': this.storeIds, 'type': 'M', 'others': 'Y',
    'groupName': this.groupName, 'storename': this.storename, storecount: null,
    'storedisplayname': this.storedisplayname, 'DefaultLoad': this.DefaultLoad,
    otherStoresArray: this.otherStoresArray, otherStoreIds: this.otherStoreIds
  };
  Performance: string = 'Load';

  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';
  FromDate: any = '';
  ToDate: any = '';
  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    Types: [
      { 'code': 'MTD', 'name': 'MTD' },
      { 'code': 'YTD', 'name': 'YTD' },
      { 'code': 'PYTD', 'name': 'PYTD' },
      { 'code': 'LM', 'name': 'Last Month' },
      { 'code': 'PM', 'name': 'Same Month PY' },
    ]
  }

  selectedDataGrouping: any = [];
  dataGrouping: any = [
    { "ARG_ID": 1, "ARG_LABEL": "Store", "ARG_SEQ": 1, "columnname": "store", "Active": 'Y' },
    { "ARG_ID": 2, "ARG_LABEL": "New/Used", "ARG_SEQ": 2, "columnname": "ad_dealtype", "Active": 'Y' },
    { "ARG_ID": 3, "ARG_LABEL": "Sale Type", "ARG_SEQ": 3, "columnname": "ad_dealtype2", "Active": 'Y' },
    { "ARG_ID": 15, "ARG_LABEL": "Funding Type", "ARG_SEQ": 4, "columnname": "", "Active": 'N' },
    { "ARG_ID": 16, "ARG_LABEL": "Lender", "ARG_SEQ": 5, "columnname": "ad_Lender", "Active": 'Y' },
    { "ARG_ID": 17, "ARG_LABEL": "Deal Status", "ARG_SEQ": 6, "columnname": "ad_dealstatus", "Active": 'Y' },
    { "ARG_ID": 18, "ARG_LABEL": "Purchase/Trade", "ARG_SEQ": 7, "columnname": "", "Active": 'N' },
    { "ARG_ID": 19, "ARG_LABEL": "Year", "ARG_SEQ": 8, "columnname": "ad_Year", "Active": 'Y' },
    { "ARG_ID": 20, "ARG_LABEL": "Make", "ARG_SEQ": 9, "columnname": "ad_make", "Active": 'Y' },
    { "ARG_ID": 21, "ARG_LABEL": "Model", "ARG_SEQ": 10, "columnname": "ad_model", "Active": 'Y' },
    { "ARG_ID": 22, "ARG_LABEL": "Car/Truck", "ARG_SEQ": 11, "columnname": "", "Active": 'N' },
    { "ARG_ID": 23, "ARG_LABEL": "Trim", "ARG_SEQ": 12, "columnname": "ad_trim", "Active": 'Y' },
    { "ARG_ID": 24, "ARG_LABEL": "Style", "ARG_SEQ": 13, "columnname": "ad_styleid", "Active": 'N' },
    { "ARG_ID": 25, "ARG_LABEL": "Month", "ARG_SEQ": 14, "columnname": "ad_Month", "Active": 'Y' },
    { "ARG_ID": 4, "ARG_LABEL": "Salesperson", "ARG_SEQ": 15, "columnname": "salesperson", "Active": 'Y' },
    { "ARG_ID": 5, "ARG_LABEL": "F&I Manager", "ARG_SEQ": 16, "columnname": "fimanager", "Active": 'Y' },
    { "ARG_ID": 6, "ARG_LABEL": "Sales Manager", "ARG_SEQ": 17, "columnname": "salesmanager", "Active": 'Y' },
    { "ARG_ID": 7, "ARG_LABEL": "Desk Manager", "ARG_SEQ": 18, "columnname": "", "Active": 'N' },
    { "ARG_ID": 8, "ARG_LABEL": "Closer", "ARG_SEQ": 19, "columnname": "", "Active": 'N' },
    { "ARG_ID": 9, "ARG_LABEL": "BDC Representative", "ARG_SEQ": 20, "columnname": "", "Active": 'N' },
    { "ARG_ID": 10, "ARG_LABEL": "Customer ZIP", "ARG_SEQ": 21, "columnname": "ad_custzip", "Active": 'Y' },
    { "ARG_ID": 11, "ARG_LABEL": "Customer State", "ARG_SEQ": 22, "columnname": "ad_custstate", "Active": 'Y' },
    { "ARG_ID": 12, "ARG_LABEL": "Customer Age", "ARG_SEQ": 23, "columnname": "ad_custAge", "Active": 'Y' },
    { "ARG_ID": 13, "ARG_LABEL": "Sales Team", "ARG_SEQ": 24, "columnname": "", "Active": 'N' },
    { "ARG_ID": 14, "ARG_LABEL": "Sales Group", "ARG_SEQ": 25, "columnname": "", "Active": 'N' },
  ]
  ProductDeals: any = 'No'

  TeamsArray: any = [
    { name: 'Sales Persons', code: 'SP' },
    { name: 'Sales Managers', code: 'SM' },
    { name: 'F&I Managers', code: 'FM' },
  ]


  Teams = 'SP'
  salesPersons: any = [];
  salesManagers: any = [];
  financeManager: any = [];
  selectedSalespersonvalues: any = [];
  selectedSalesManagersvalues: any = [];
  selectedFiManagersvalues: any = [];
  acquisitionsource: any = []
  Acquisition: any = []
  overallSelectedpeople: any = 0
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe , .reportpeople-card');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  ngOnChanges(changes: SimpleChanges) {
    console.log(changes, 'Report');
    this.ngChanges = changes['header'].currentValue[0];
    this.FromDate = this.ngChanges.fromdate;
    this.ToDate = this.ngChanges.todate;
    this.DateType = this.ngChanges.datevaluetype;
    this.groupId = this.ngChanges.groups
    this.setDates(this.ngChanges.datevaluetype)
    // this.StoresData(this.ngChanges)
    this.neworused = []
    this.neworused = this.ngChanges.dealType;
    this.retailorlease = [];
    this.retailorlease = this.ngChanges.saleType
    this.dealstatus = [];
    this.dealstatus = this.ngChanges.dealStatus
    this.reporttotal = this.ngChanges.toporbottom
    this.selectedDataGrouping = [...this.ngChanges.dataGroupings]
    this.ProductDeals = this.ngChanges.ProductDeals
    this.DefaultLoad = this.ngChanges.DefaultLoad
    this.otherStoreIds = this.ngChanges.otherstoreids

  }
  constructor(private shared: Sharedservice, private datesSrvc: Setdates, private toast: ToastService) { }

  ngOnInit() {
    this.selectedDataGrouping = []
    this.selectedDataGrouping.push(this.dataGrouping[0])
    this.selectedDataGrouping.push(this.dataGrouping[1])

    this.getGroups();
    this.acquisitionsrc();

  }


  SelectedData(val: any) {
    const index = this.selectedDataGrouping.findIndex((i: any) => i == val);
    if (index >= 0) {
      this.selectedDataGrouping.splice(index, 1);
    } else {
      if (this.selectedDataGrouping.length >= 3) {
        this.toast.show('Select up to 3 Filters only to Group your data', 'warning', 'Warning');
      } else {
        this.selectedDataGrouping.push(val);
      }
    }
  }
  acquisitionsrc() {
    const obj = {}
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetAcquisitionSourceList', obj).subscribe((res: any) => {
      if (res.status == 200) {
        this.acquisitionsource = res.response
        if (this.ngChanges.as[0] == "All") {
          this.Acquisition = this.acquisitionsource.map(function (a: any) {
            return a.ad_acquisition_source;
          });
        } else {
          this.Acquisition = []
          this.Acquisition = this.ngChanges.as
        }
      }
    })
  }
  teamsselection(e: any) {
    this.Teams = e;
  }
  Deals(e: any) {
    this.ProductDeals = e
  }
  employees(block: any, e: any) {

    if (block == 'SP') {
      const index = this.selectedSalespersonvalues.findIndex(
        (i: any) => i == e
      );
      if (index >= 0) {
        this.selectedSalespersonvalues.splice(index, 1);
      } else {
        this.selectedSalespersonvalues.push(e);
      }
    }
    if (block == 'SM') {
      const index = this.selectedSalesManagersvalues.findIndex(
        (i: any) => i == e
      );
      if (index >= 0) {
        this.selectedSalesManagersvalues.splice(index, 1);
      } else {
        this.selectedSalesManagersvalues.push(e);
      }
    }
    if (block == 'FM') {
      const index = this.selectedFiManagersvalues.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.selectedFiManagersvalues.splice(index, 1);
      } else {
        this.selectedFiManagersvalues.push(e);
      }
    }
    if (block == 'AllSP') {

      if (e == 1) {
        this.selectedSalespersonvalues = [];
      } else if (e == 0) {
        this.selectedSalespersonvalues = this.salesPersons.map(function (
          a: any
        ) {
          return a.SPID;
        });
      }
    }
    if (block == 'AllSM') {
      if (e == 1) {
        this.selectedSalesManagersvalues = [];
      } else if (e == 0) {
        this.selectedSalesManagersvalues = this.salesManagers.map(function (
          a: any
        ) {
          return a.SmId;
        });
      }
    }
    if (block == 'AllFM') {
      if (e == 1) {
        this.selectedFiManagersvalues = [];
      } else if (e == 0) {
        this.selectedFiManagersvalues = this.financeManager.map(function (
          a: any
        ) {
          return a.FiId;
        });
      }
    }
    this.overallSelectedpeople = this.selectedFiManagersvalues.length + this.selectedSalesManagersvalues.length + this.selectedSalespersonvalues.length

  }
  spinnerLoaderteams: boolean = false
  spinnerLoader:boolean = false;
  getEmployees(val: any,strids:any, barorpopup?: any) {
    this.salesManagers = []
    this.salesPersons = []
    this.financeManager = []
    const obj = {
      AS_ID: strids,
      type: val,
    };
    this.spinnerLoaderteams = true;
    this.spinnerLoader = true;
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetEmployeesDev', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          this.spinnerLoader = false;
          if (val == 'SP') {
            this.salesPersons = res.response.filter((e: any) => e.SPName != 'Unknown' && e.SPName != '');
            this.selectedSalespersonvalues = this.salesPersons.map(function (a: any) { return a.SPID; });
            if (barorpopup == 'Bar') {
              if (this.ngChanges.sp == '0') {
                this.selectedSalespersonvalues = this.salesPersons.map(function (a: any) { return a.SPID; });
              }
              if (this.ngChanges.sp == '') {
                this.selectedSalespersonvalues = []
              }
              if (this.ngChanges.sp != '' && this.salesManagers && this.salesManagers.length > 0) {
                let spids = this.ngChanges.sp.split(',');
                this.selectedSalespersonvalues = spids;
              }
            }

          }

          if (val == 'F') {
            this.financeManager = res.response.filter((e: any) => e.FiName != 'Unknown');
            this.selectedFiManagersvalues = this.financeManager.map(function (a: any) { return a.FiId; });
            if (barorpopup == 'Bar') {
              if (this.ngChanges.fm == '0') {
                this.selectedFiManagersvalues = this.financeManager.map(function (a: any) { return a.FiId; });
              }
              if (this.ngChanges.fm == '') {
                this.selectedFiManagersvalues = []
              }
              if (this.ngChanges.fm != '' && this.financeManager && this.financeManager.length > 0) {
                let fiids = this.ngChanges.fm.split(',');
                this.selectedFiManagersvalues = fiids;
              }
            }
          }
          if (val == 'M') {
            this.salesManagers = res.response.filter((e: any) => e.SmName != 'Unknown');
            this.selectedSalesManagersvalues = this.salesManagers.map(function (a: any) { return a.SmId; });
            if (barorpopup == 'Bar') {
              if (this.ngChanges.sm == '0') {
                this.selectedSalesManagersvalues = this.salesManagers.map(function (a: any) { return a.SmId; });
              }
              if (this.ngChanges.sm == '') {
                this.selectedSalesManagersvalues = []
              }
              if (this.ngChanges.sm != '' && this.salesManagers && this.salesManagers.length > 0) {
                let smids = this.ngChanges.sm.split(',');
                this.selectedSalesManagersvalues = smids;
              }
            }

            this.spinnerLoaderteams = false;
          }
          this.overallSelectedpeople = this.selectedFiManagersvalues.length + this.selectedSalesManagersvalues.length + this.selectedSalespersonvalues.length

        } else {
          this.spinnerLoader = false;
          this.toast.show('Invalid Details.', 'danger', 'Error');
        }
      },
      (error) => {
      }
    );
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
  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {

      if (this.shared.common.pageName == 'Sales Gross') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.groupId = this.ngChanges.groups;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
           this.otherStoresArray = this.shared.common.OtherStoresData && this.shared.common.OtherStoresData.length > 0 ? this.shared.common.OtherStoresData[0].Stores : []
          this.storeIds = this.ngChanges.stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()

        }
      }
    })

  }
  scrollPosition = 0;

  getScrollPosition(event: any): void {
    this.scrollPosition = event.target.scrollLeft;
  }
  getGroups() {
    if (this.shared.common.groupsandstores != undefined) {
      if (this.shared.common.groupsandstores.length > 0) {
        this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
        this.groupId = this.ngChanges.groups
        this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
         this.otherStoresArray = this.shared.common.OtherStoresData && this.shared.common.OtherStoresData.length > 0 ? this.shared.common.OtherStoresData[0].Stores : []

        this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
        this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
        this.getStoresandGroupsValues()
        // this.StoresData(this.ngChanges)
      }
    }
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
    this.storesFilterData.otherStoreIds = this.otherStoreIds;
    this.storesFilterData.otherStoresArray = this.otherStoresArray;
    this.storesFilterData = {
      groupsArray: this.groupsArray,
      groupId: this.groupId,
      storesArray: this.stores,
      storeids: this.storeIds,
      groupName: this.groupName,
      storename: this.storename,
      storecount: this.storecount,
      storedisplayname: this.storedisplayname,
      'type': 'M', 'others': 'Y', 'DefaultLoad': this.DefaultLoad, otherStoresArray: this.otherStoresArray, otherStoreIds: this.otherStoreIds
    };
    console.log(this.storesFilterData, 'Store FIlter Data');
    let allstrids = [];
    allstrids = [...this.storeIds]
 
  }
  StoresData(data: any) {

    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
    this.otherStoreIds = data.otherStoreIds;
    let allstrids = [];
    allstrids = [...this.storeIds,...this.otherStoreIds]
    this.getEmployees('SP', allstrids,'');
    this.getEmployees('F', allstrids,'');
    this.getEmployees('M', allstrids,'');

  }
  setDates(type: any) {
    // this.datevaluetype=
    if (type != 'C') {
      this.displaytime = '( ' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ' )';
    }
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
  updatedDates(data: any) {
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
    localStorage.setItem('time', this.DateType);

  }
  neworused: any = ['New', 'Used'];
  retailorlease: any = ['Retail ', 'Lease ', 'Misc '];
  dealTypeData: any = ['All', 'Retail and Lease', 'Retail', 'Lease', 'Wholesale', 'Misc', 'Fleet', 'Demo', 'Special Order', 'Rental', 'Dealer Trade']

  includeorexclude: any = [''];
  dealstatus: any = ['Delivered', 'Capped', 'Finalized'];
  selecttarget: any = ['F'];
  reporttotal: any = 'T';
  includecharges: any = [''];
  Transactorgl: any = ['T'];
  multipleorsingle(block: any, e: any) {
    if (block == 'NU') {
      const index = this.neworused.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.neworused.splice(index, 1);
      } else {
        this.neworused.push(e);
      }
    }
    if (block == 'RL') {
      const index = this.retailorlease.findIndex((i: any) => i == e);
      if (index >= 0) {
        if (e == 'All') {
          // this.retailorlease.splice(index, 1);
          this.retailorlease = []
        } else {
          this.retailorlease.splice(index, 1);
          let allindex = this.retailorlease.findIndex((i: any) => i == 'All');
          if (allindex >= 0) {
            this.retailorlease.splice(allindex, 1);
          }
        }
      } else {
        if (e == 'All') {
          const dealdata = JSON.stringify(this.dealTypeData)
          this.retailorlease = JSON.parse(dealdata)
        } else {
          this.retailorlease.push(e);
          if (this.retailorlease.length == this.dealTypeData.length - 1) {
            const dealdata = JSON.stringify(this.dealTypeData)
            this.retailorlease = JSON.parse(dealdata)
          }
        }
      }
    }
    if (block == 'PH') {
      this.includeorexclude = [];
      this.includeorexclude.push(e);
    }
    if (block == 'DS') {
      const index = this.dealstatus.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.dealstatus.splice(index, 1);
      } else {
        this.dealstatus.push(e);
      }
    }
    if (block == 'ST') {
      const index = this.selecttarget.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.selecttarget.splice(index, 1);
      } else {
        this.selecttarget.push(e);
      }
    }
    if (block == 'RT') {
      this.reporttotal = e;
    }
    if (block == 'IC') {
      this.includecharges = [];
      this.includecharges.push(e);
    }
    if (block == 'SRC') {
      this.Transactorgl = [];
      this.Transactorgl.push(e);
    }
    if (block == 'AS') {
      if (e == 'All') {
        if (this.Acquisition.length == this.acquisitionsource.length) {
          this.Acquisition = []
        } else {
          this.Acquisition = this.acquisitionsource.map(function (a: any) {
            return a.ad_acquisition_source;
          });
        }
      }
      else {
        const index = this.Acquisition.findIndex((i: any) => i == e);
        if (index >= 0) {
          this.Acquisition.splice(index, 1);
        } else {
          this.Acquisition.push(e);
        }
      }
    }
  }
  viewreport() {
    this.activePopover = -1
    let peoples = this.selectedFiManagersvalues.length + this.selectedSalesManagersvalues.length + this.selectedSalespersonvalues.length
    let aqusrc: any = ['All']
    if ((!this.storeIds || this.storeIds.length === 0) && (!this.otherStoreIds || this.otherStoreIds.length === 0)) {
      this.toast.show(
        'Please Select Atleast One Store',
        'warning',
        'Warning'
      );
    }
    else if (this.selectedDataGrouping.length == 0) {
      this.toast.show(
        'Please Select Atleast One Grouping',
        'warning',
        'Warning'
      );
    } else if (this.dealstatus.length == 0) {
      this.toast.show(
        'Please Select Atleast One Deal Status',
        'warning',
        'Warning'
      );

    }
    else if (this.retailorlease.length == 0) {
      this.toast.show(
        'Please Select Atleast One Deal Type',
        'warning',
        'Warning'
      );

    }
    else if (peoples == 0) {
      this.toast.show(
        'Please Select Atleast One From People Filter',
        'warning',
        'Warning'
      );

    }
    else if (this.Acquisition.length == 0) {
      this.toast.show('Please Select Atleast One Acquisition Source', 'warning', 'Warning');
    }
    else {
      const data = {
        Reference: 'Sales Gross',
        FromDate: this.shared.datePipe.transform(this.FromDate, 'MM-dd-yyyy'),
        ToDate: this.shared.datePipe.transform(this.ToDate, 'MM-dd-yyyy'),
        storeValues: this.storeIds,
        Spvalues:
          this.selectedSalespersonvalues.length == this.salesPersons.length
            ? '0'
            : this.selectedSalespersonvalues.toString(),
        SMvalues:
          this.selectedSalesManagersvalues.length == this.salesManagers.length
            ? '0'
            : this.selectedSalesManagersvalues.toString(),
        FIvalues:
          this.selectedFiManagersvalues.length == this.financeManager.length
            ? '0'
            : this.selectedFiManagersvalues.toString(),
        dealType: this.neworused,
        saleType: this.retailorlease,
        dealStatus: this.dealstatus,
        dataGroupingvalues: this.selectedDataGrouping,
        dateType: this.DateType,
        reportTotal: this.reporttotal,
        acquisition: this.Acquisition.length == this.acquisitionsource.length ? aqusrc : this.Acquisition,
        groups: this.groupId,
        ProductDeals: this.ProductDeals,
        otherstoreids: this.otherStoreIds
      };
      this.shared.api.SetReports({
        obj: data,
      });
      this.close();
      // }
      // }
    }
  }

  close() {
    this.Performance = 'Unload';
  }
}
