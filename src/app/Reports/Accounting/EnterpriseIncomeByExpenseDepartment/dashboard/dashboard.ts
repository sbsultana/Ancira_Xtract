import { Component, HostListener, Injector } from '@angular/core';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Subscription } from 'rxjs';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { DatePipe, formatDate } from '@angular/common';
import { Api } from '../../../../Core/Providers/Api/api';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { NgxSpinnerService } from 'ngx-spinner';
import { Title } from '@angular/platform-browser';
import { common } from '../../../../common';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { Router } from '@angular/router';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { EnterpriseincomebyexpensedepartmentDetails } from '../enterpriseincomebyexpensedepartment-details/enterpriseincomebyexpensedepartment-details';
import { Workbook } from 'exceljs';
import FileSaver from 'file-saver';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  Current_Date: any;
  Report: any = '';
  subscriptionPrint!: Subscription;
  subscriptionPdf!: Subscription;
  subscriptionExcel!: Subscription;
  subscriptionReport!: Subscription;
  FSData: any = [];
  date: any;
  ExpenseTrendByStoreKeys: any = [];
  AllDatakeys: any = [];
  ExpenseTrendByStore_Excel: any;
  XpenseTrendByStoreKeys: any = [];
  ExpenseTrendByStore: any;
  ExpenseTrendByStore_ExcelMonth: any;
  XpenseTrendByStoreKeysMonth: any = [];
  ExpenseTrendByStoreKeysMonth: any = [];
  AllDatakeysMonth: any = [];
  IncomeSummaryData: any;
  ServiceTrendTotalsData: any = [];
  // Filter: any;
  // StoreName: any;
  SubFilter: any;
  SelectedTab: any = [];
  Month: any = '';
  // stores: any;
  PreviousMonths: any = '13';
  selectedstorevalues: any = [];
  selectedstorename: any;

  StoreName: any = 'All Stores';
  Filter: any = [];
  ShowHideGP: any = 'Show';
  ShowHideBudget: any = 'Show';

  filters: string[] = [];
  selectedFilters: string[] = [];
  selectedLabel: string = '( All )';
  activePopover: number | null = null;
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month',
    maxDate: new Date()
  };
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  selectDate: Date = new Date();
  currentMonth: any = '';
  StoreValues: any = 2;
  groups: any = 1;
  PresentDayDate: string;
  header: any = [
    {
      type: 'Bar',
      StoreValues: this.StoreValues,
      Month: this.Month,
      Filter: this.Filter,
      ShowHideGP: this.ShowHideGP,
      ShowHideBudget: this.ShowHideBudget,
      groups: this.groups
    },
  ];
  stores: any = [];
  popup: any = [{ type: 'Popup' }];
  CurrentRoute: any;
  SetTitle: any;
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = '0';
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  constructor(
    private datepipe: DatePipe,
    public apiSrvc: Api,
    private ngbmodal: NgbModal,
    private ngbmodalActive: NgbActiveModal,
    private spinner: NgxSpinnerService,
    private title: Title,
    private comm: common,
    private toast: ToastService,
    private router: Router,
    private injector: Injector,
    public shared: Sharedservice
  ) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')
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
    this.CurrentRoute = this.router.url.substring(1);
    switch (this.CurrentRoute) {
      // case 'FixedIncomeByExpense':
      //   this.filters = ['Service', 'Parts'];
      //   this.SetTitle = 'Fixed Income / Expense';
      //   break;
      // case 'VariableIncomeByExpense':
      //   this.filters = ['New', 'Used'];
      //   this.SetTitle = 'Variable Income / Expense';
      //   break;
      case 'EnterpriseExpense':
        this.filters = ['New', 'Used', 'Service', 'Parts', 'Collision'];
        this.SetTitle = 'Enterprise Expense';
        break;
      default:
        this.filters = [];
        this.SetTitle = 'Income / Expense';
        break;
    }
    this.selectedFilters = [...this.filters];
    const lastMonth = new Date();
    let today = new Date();
    if (today.getDate() < 5) {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth() - 1));
    } else {
      this.date = new Date(lastMonth.setMonth(lastMonth.getMonth()));
    }
    this.Month = this.date



    this.title.setTitle(this.comm.titleName + `-${this.SetTitle}`);
    if (localStorage.getItem('Fav') != 'Y') {
      const data = {
        title: this.SetTitle,
        path1: '',
        path2: '',
        path3: '',
        Month: this.date,
        stores: this.storeIds,
        filter: this.Filter,
        ShowHideGP: this.ShowHideGP,
        ShowHideBudget: this.ShowHideBudget,
        PreviousMonths: this.PreviousMonths,
        groups: this.groups,
      };
      this.apiSrvc.SetHeaderData({
        obj: data,
      });

      this.currentMonth = this.Month;
      this.selectDate = this.Month
      this.GetDataByMonths(this.currentMonth, this.selectedFilters);
    }
    // });
    const format = 'ddMMyyyy';
    const locale = 'en-US';
    const myDate = new Date();
    const formattedDate = formatDate(myDate, format, locale);
    this.PresentDayDate = formattedDate;
  }
  roleId: any;
  ngOnInit(): void {
    this.roleId = localStorage.getItem('roleId');
    console.log('role Id', this.roleId);
  }

  StoreNamesHeadings: any = [];
  MonthsHeadings: any = [];

  Scrollpercent: any = 0;
  updateVerticalScroll(event: any): void {
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );
  }
  selectedType: 'Department' | 'Vendor' = 'Department';
  onTypeChange() {
    console.log(this.selectedType);
  }
  togglePopover(index: number) {
    this.activePopover = this.activePopover === index ? null : index;
  }

  isSelected(filter: string): boolean {
    return this.selectedFilters.includes(filter);
  }

  toggleFilter(filter: string) {
    const index = this.selectedFilters.indexOf(filter);
    if (index >= 0) {
      this.selectedFilters.splice(index, 1);
    } else {
      this.selectedFilters.push(filter);
    }
    this.updateLabel();
  }

  updateLabel() {
    if (!this.selectedFilters || this.selectedFilters.length === 0) {
      this.selectedLabel = '( Select )';
    } else if (this.selectedFilters.length === this.filters.length) {
      this.selectedLabel = '( All )';
    } else if (this.selectedFilters.length === 1) {
      this.selectedLabel = `( ${this.selectedFilters[0]} )`;
    } else {
      this.selectedLabel = `( Selected ${this.selectedFilters.length} )`;
    }
  }

  applyDateChange() {
    if (!this.storeIds || this.storeIds.length === 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      return;
    }
    if (!this.selectedFilters || this.selectedFilters.length === 0) {
      this.toast.show('Please Select Atleast One Department', 'warning', 'Warning');
      return;
    }
    this.currentMonth = this.formatMonth(this.selectDate);
    this.GetDataByMonths(this.currentMonth, this.selectedFilters);
    this.activePopover = null;
    this.isLoading = true;
  }

  formatMonth(date: Date): string {
    const year = date.getFullYear();
    const month = ('0' + (date.getMonth() + 1)).slice(-2);
    return `${year}-${month}`;
  }


  storeName: any;
  isLoading = true;
  NoData = false;
  IncomeSummaryDataKeys: any = [];
  ServiceTrendingSubKeys: any = [];
  ServiceTrendingXlm: any = [];
  GetDataByMonths(date: any, filters: any) {
    this.IncomeSummaryData = [];

    this.spinner.show();
    const DateToday = this.datepipe.transform(new Date(this.date), 'yyyy-MM-dd');
    this.Month = DateToday
    const obj = {
      DEPARTMENT: filters.toString(),
      AS_IDS: this.storeIds.toString(),
      DATE: this.shared.datePipe.transform(this.currentMonth,'yyyy-MM')+'-10',
    };
    console.log(obj);

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetEnterpriseIncomeExpenseTrendReport', obj)
      .subscribe(
        (res: any) => {
          this.isLoading = false;
          if (res?.status === 200 && Array.isArray(res.response) && res.response.length) {

            this.IncomeSummaryData = res.response;
            console.log('IncomeSummaryData', this.IncomeSummaryData);

            const incomeSummaryKeys = Object.keys(res.response[0] || {}).slice(6);

            const lastTwoValues = incomeSummaryKeys.slice(-2);
            const remainingValues = incomeSummaryKeys.slice(0, -2);

            this.IncomeSummaryDataKeys = [...lastTwoValues, ...remainingValues];

            this.IncomeSummaryDataKeys = this.IncomeSummaryDataKeys.filter((dealership: string) =>
              dealership !== 'TOTAL' &&
              dealership !== 'AVG' &&
              dealership !== 'SEQ' &&
              (this.ShowHideGP === 'Show' || (this.ShowHideGP === 'Hide' && !dealership.endsWith('_GP'))) &&
              (this.ShowHideBudget === 'Show' ||
                (this.ShowHideBudget === 'Hide' &&
                  !dealership.endsWith('_BDGT') &&
                  !dealership.endsWith('_VAR')))
            );

            this.IncomeSummaryData.forEach((row: any) => {
              row.SubData = [];
              row.data2sign = '+';
            });

            console.log('Keys ', this.IncomeSummaryDataKeys);

            this.NoData = this.IncomeSummaryDataKeys.length > 6;

          } else {
            this.IncomeSummaryData = [];
            this.IncomeSummaryDataKeys = [];
            this.NoData = false;
          }
          this.spinner.hide();
        },
        (error) => {
          console.error(error);
          this.isLoading = false;
          this.IncomeSummaryData = [];
          this.IncomeSummaryDataKeys = [];
          this.NoData = false;
          this.spinner.hide();
        }
      );

  }

  formatKey(key: string): string {
    if (key.endsWith('_GP')) {
      return '% of GP';
    } else if (key.endsWith('_BDGT')) {
      return 'Budget';
    } else if (key.endsWith('_VAR')) {
      return 'Variance';
    }
    return key;
  }

  isKeyVisible(key: string): boolean {
    return key !== '-1';
  }
  FixedExpenseLayerOne: any;
  FixedExpenseLayerOneKeys: any;
  GetFixedExpenseLayerOne(STR: any, index: number, ParentCode: any): void {
    this.spinner.show();
    console.log(STR);
    // STR.expanded = !STR.expanded;
    this.FixedExpenseLayerOne = [];
    this.FixedExpenseLayerOneKeys = [];
    const DateToday = this.datepipe.transform(
      new Date(this.currentMonth),
      'dd-MMM-yyyy'
    );
    const obj = {
      LEVEL: '1',
      AS_IDs: this.storeIds.toString(),
      STORENAMES: '',
      DEPARTMENT: this.selectedFilters.toString(),
      DATE: '10-'+this.shared.datePipe.transform(this.currentMonth,'MMM-yyyy'),
      SUBTYPEDETAIL: STR.DISPLAY_LABLE,
      ACCTTYPEDETAIL: '',
      LABLECODE: STR.PARENTLABLECODE,
      ACCTNUM: '',
      ACCTDESC: '',
      Control: '',
      Category: this.selectedType
    };
    console.log(obj);

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetEnterpriseExpenseTrendDetailByLevelDeptV1', obj)
      .subscribe(
        (x: any) => {
          if (x.status === 200) {
            this.FixedExpenseLayerOne = x.response;
            console.log('FixedExpenseLayerOne', this.FixedExpenseLayerOne);

            if (x.response[0] && typeof x.response[0] === 'object') {
              const FixedExpenseLayerOneKeys = Object.keys(x.response[0]).slice(
                3
              );
              const lastTwoValues = FixedExpenseLayerOneKeys.slice(-2);
              const remainingValues = FixedExpenseLayerOneKeys.slice(0, -2);
              this.FixedExpenseLayerOneKeys =
                lastTwoValues.concat(remainingValues);
            }

            if (
              this.FixedExpenseLayerOne &&
              this.FixedExpenseLayerOne.length > 0
            ) {
              this.FixedExpenseLayerOne.forEach((item: any) => {
                item.SubLayerData = [];
                item.dataLayer2sign = '+';
              });

              this.IncomeSummaryData.forEach((val: any) => {
                val.SubData = [];
                val.data2sign = '+';
                this.FixedExpenseLayerOne.forEach((ele: any) => {
                  if (val.DISPLAY_LABLE === ele.Level1Head && val.PARENTLABLECODE === ParentCode) {
                    val.SubData.push(ele);
                    val.data2sign = '-';
                  }
                });
              });

              console.log('Final Data', this.IncomeSummaryData);
              this.FixedExpenseLayerOneKeys =
                this.FixedExpenseLayerOneKeys.filter(
                  (key: string) => key !== 'TOTAL' && key !== 'AVG' && key !== 'SEQ'
                );
              console.log('Layer One Keys', this.FixedExpenseLayerOneKeys);

              // this.NoData = false;
            } else {
              this.NoData = true;
            }
            this.spinner.hide();
          }
        },
        (error: any) => {
          console.error(error);
          this.spinner.hide();
        }
      );
  }

  isNumber(value: any): boolean {
    return (
      value !== null && value !== undefined && value !== 0 && !isNaN(value)
    );
  }
  FixedExpenseLayerTwo: any[] = [];
  FixedExpenseLayerTwoKeys: string[] = [];
  GetFixedExpenseLayerTwo(STR: any, Layer: any, index: number): void {
    // STR.expanded = !STR.expanded;
    this.spinner.show();
    this.FixedExpenseLayerTwo = [];
    this.FixedExpenseLayerTwoKeys = [];
    const DateToday = this.datepipe.transform(
      new Date(this.currentMonth),
      'dd-MMM-yyyy'
    );
    const obj = {
      LEVEL: '2',
      AS_IDs: this.storeIds.toString(),
      STORENAMES: '',
      DEPARTMENT: this.selectedFilters.toString(),
      DATE: '10-'+this.shared.datePipe.transform(this.currentMonth,'MMM-yyyy'),
      SUBTYPEDETAIL: STR.DISPLAY_LABLE,
      ACCTTYPEDETAIL: '',
      LABLECODE: STR.PARENTLABLECODE,
      ACCTNUM: Layer.AcctNum,
      ACCTDESC: Layer.Desc,
      Control: '',
    };
    console.log(obj);

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetEnterpriseExpenseTrendDetailByLevelDept', obj)
      .subscribe(
        (x: any) => {
          if (x.status === 200) {
            this.FixedExpenseLayerTwo = x.response;
            console.log('FixedExpenseLayerTwo', this.FixedExpenseLayerTwo);

            if (x.response[0] && typeof x.response[0] === 'object') {
              const FixedExpenseLayerTwoKeys = Object.keys(x.response[0]).slice(
                5
              );
              const lastTwoValues = FixedExpenseLayerTwoKeys.slice(-2);
              const remainingValues = FixedExpenseLayerTwoKeys.slice(0, -2);
              this.FixedExpenseLayerTwoKeys =
                lastTwoValues.concat(remainingValues);
            }

            this.FixedExpenseLayerOne.forEach((val: any) => {
              val.SubLayerData = [];
              val.dataLayer2sign = '+';
              this.FixedExpenseLayerTwo.forEach((ele: any) => {
                if (val.Desc === ele.Level2Head) {
                  val.SubLayerData.push(ele);
                  val.dataLayer2sign = '-';
                }
              });
            });

            console.log('Final Data', this.IncomeSummaryData);
            this.FixedExpenseLayerTwoKeys =
              this.FixedExpenseLayerTwoKeys.filter(
                (key: string) => key !== 'TOTAL' && key !== 'AVG' && key !== 'SEQ'
              );
            console.log('Layer Two Keys', this.FixedExpenseLayerTwoKeys);

            // this.NoData = this.FixedExpenseLayerTwo.length === 0;
          }
          this.spinner.hide();
        },
        (error: any) => {
          console.error(error);
          this.spinner.hide();
        }
      );
  }
  isPercentage(value: any): boolean {
    // Assuming your logic to determine if the value is a percentage
    return typeof value === 'string' && value.includes('%');
  }

  expandorcollapse(ind: any, e: any, ref: any, Item: any, parentData: any) {
    let id = (e.target as Element).id;
    console.log(ref, Item);
    if (id == 'D_' + ind) {
      console.log(id);
      if (ref == '-') {
        this.IncomeSummaryData[ind].data2sign = '+';
        this.IncomeSummaryData[ind].SubData = [];
      }
      if (ref == '+') {
        this.IncomeSummaryData[ind].data2sign = '-';
        this.GetFixedExpenseLayerOne(Item, ind, Item.PARENTLABLECODE);
      }
    }
    if (id == 'DN_' + ind) {
      if (ref == '-') {
        this.FixedExpenseLayerOne[ind].dataLayer2sign = '+';
        this.FixedExpenseLayerOne[ind].SubLayerData = [];
      }
      if (ref == '+') {
        this.FixedExpenseLayerOne[ind].dataLayer2sign = '-';
        this.GetFixedExpenseLayerTwo(parentData, Item, ind);
        this.IncomeSummaryData[ind].data2sign = '-';
      }
    }
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  SFstate: any;
  StoreCodes: any;
  block: any = '';
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == this.SetTitle) {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.apiSrvc.GetReportOpening().subscribe((res) => {
      console.log(res);
      if (res.obj.Module == this.SetTitle) {
        document.getElementById('report')?.click();
      }
    });
    this.apiSrvc.GetReports().subscribe((data) => {
      //console.log(data)
      if (data.obj.Reference == this.SetTitle) {
        if (data.obj.header == undefined) {
          this.date = data.obj.month;
          this.Month = data.obj.month;
          this.StoreValues = data.obj.storeValues;
          this.StoreName = data.obj.Sname;
          this.Filter = data.obj.filter;
          this.ShowHideGP = data.obj.ShowHideGP;
          this.ShowHideBudget = data.obj.ShowHideBudget
          this.SubFilter = data.obj.subfilters;
          this.StoreCodes = data.obj.storecode;
          this.groups = data.obj.groups;
          this.PreviousMonths = data.obj.PreviousMonths;
          this.index = '';
          this.Scrollpercent = 0;
        } else {
          if (data.obj.header == 'Yes') {
            this.StoreValues = data.obj.storeValues;
          }
        }
        if (this.StoreValues != '') {
          this.GetDataByMonths(this.currentMonth, this.selectedFilters);
        } else {
          this.NoData = true;
          this.IncomeSummaryData = [];
        }
        const headerdata = {
          title: this.SetTitle,
          path1: '',
          path2: '',
          path3: '',
          Month: this.Month,
          filter: this.Filter,
          ShowHideGP: this.ShowHideGP,
          ShowHideBudget: this.ShowHideBudget,
          stores: this.StoreValues,
          storecode: this.StoreCodes,
          Sname: this.storeName,
          PreviousMonths: this.PreviousMonths,
          groups: this.groups,
        };
        this.apiSrvc.SetHeaderData({
          obj: headerdata,
        });
        this.header = [
          {
            type: 'Bar',
            StoreValues: this.StoreValues,
            Month: new Date(this.Month),
            Filter: this.Filter,
            ShowHideGP: this.ShowHideGP,
            ShowHideBudget: this.ShowHideBudget,
            groups: this.groups
          },
        ];
        console.log(headerdata);
      }
    });
    this.apiSrvc.getExportToExcelAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      this.SFstate = res.obj.state;
      if (res.obj.title == this.SetTitle) {
        if (res.obj.state == true) {
          this.exportAsXLSX();
        }
      }
    });

    this.apiSrvc.getExportToPrintAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (res.obj.title == this.SetTitle) {
        if (res.obj.state == true) {
          this.GetPrintData();
        }
      }
    });
    this.apiSrvc.getExportToPDFAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (res.obj.title == this.SetTitle) {
        if (res.obj.state == true) {
          this.generatePDF();
        }
      }
    });
    this.apiSrvc.getExportToEmailPDFAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (res.obj.title == this.SetTitle) {
        if (res.obj.state == true) {
          // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
        }
      }
    });
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

    // this.setHeaderData();
    // this.GetData();

  }
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
  }
  reportOpen(temp: any) {
    this.ngbmodalActive = this.ngbmodal.open(temp, {
      size: 'xl',
      backdrop: 'static',
    });
  }
  SubSelectedTab1: any = [];

  closeReport() {
    this.Report = '';
  }
  getStores() {
    this.selectedstorevalues = [];
    this.stores = JSON.parse(localStorage.getItem('Stores')!);
  }
  openMainStoresDetails(ObjectOne: any, StoreName: any, Refer: any) {
    this.index = '';
    console.log(ObjectOne);
    const DetailsET = this.ngbmodal.open(EnterpriseincomebyexpensedepartmentDetails, {
      size: 'xl',
      backdrop: 'static',
      // injector: Injector.create({
      //   providers: [
      //     { provide: CurrencyPipe, useClass: CurrencyPipe }
      //   ],
      //   parent: this.injector
      // })
    });
    DetailsET.componentInstance.DetailsET = {
      LAYERONE: ObjectOne,
      DEPARTMENT: this.filters.toString(),
      STORES: this.storeIds.toString(),
      LatestDate: this.currentMonth,
      STORENAME: StoreName,
      REFERENCE: Refer,
    };
  }
  openStoresDetails(
    ObjectOne: any,
    ObjectTwo: any,
    ObjectThree: any,
    StoreName: any,
    Refer: any
  ) {
    this.index = '';
    console.log(ObjectOne);
    const DetailsET = this.ngbmodal.open(EnterpriseincomebyexpensedepartmentDetails, {
      size: 'xl',
      backdrop: 'static',
      // injector: Injector.create({
      //   providers: [
      //     { provide: CurrencyPipe, useClass: CurrencyPipe }
      //   ],
      //   parent: this.injector
      // })
    });
    DetailsET.componentInstance.DetailsET = {
      LAYERONE: ObjectOne,
      LAYERTWO: ObjectTwo,
      LAYERTHREE: ObjectThree,
      DEPARTMENT: this.filters.toString(),
      STORES: this.storeIds.toString(),
      LatestDate: this.currentMonth,
      STORENAME: StoreName,
      REFERENCE: Refer,
    };
  }
  ValueFormat: any;
  openStoreGraph(
    Object: any,
    StoreName: any,
    date: any,
    item: any,
    SummaryType: any
  ) {
    console.log(Object, StoreName, date, item, SummaryType);
    const DetailsSF = this.ngbmodal.open({
      size: 'xl',
      backdrop: 'static',
      // injector: Injector.create({
      //   providers: [
      //     { provide: CurrencyPipe, useClass: CurrencyPipe }
      //   ],
      //   parent: this.injector
      // })
    });
    DetailsSF.componentInstance.ETgraphdetails = {
      ITEM: item,
      TYPE: item.LABLEVAL,
      NAME: item.LABLE,
      SUMMARYTYPE: SummaryType,
    };
  }

  index = '';
  commentobj: any = {};
  close(data: any) {
    console.log(data);
    this.index = '';
  }
  addcmt(data: any) {
    if (data == 'A') {
      this.index = '';
      const DetailsSF = this.ngbmodal.open({
        size: 'xl',
        backdrop: 'static',
      });
      // myObject['skillItem2'] = 15;
      this.commentobj['state'] = 0;
      (DetailsSF.componentInstance.SFComments = this.commentobj),
        DetailsSF.result.then(
          (data) => {
            console.log(data);
          },
          (reason) => {
            console.log(reason);

            if (reason == 'O') {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
            } else {
              this.commentobj['state'] = 1;
              this.index = this.commentobj['indexval'];
            }
          }
        );
    }
  }

  // commentopen(Object:any, storename:any, ref:any, item:any, i:any, mod_name:any) {
  //   console.log(Object, storename, ref, item, mod_name, 'abcdefgh');
  //   this.index = i;
  //   this.commentobj = {
  //     TYPE: item.LABLEVal == undefined ? item.LABLEVAL : item.LABLEVal,
  //     NAME: item.LABLE,
  //     STORES: this.selectedstorevalues,
  //     LatestDate: this.Month,
  //     STORENAME: this.selectedstorename,
  //     Month: this.Month,
  //     ModuleId: '72',
  //     ModuleRef: mod_name,
  //     state: 1,
  //     indexval: i,
  //   };

  // }

  commentopen(item: any, i: any) {
    console.log('Selected Row  : ', item);
    this.index = i.toString();
    this.commentobj = {
      TYPE: item.LABLE1,
      NAME: item.LABLE1,
      STORES: item.SNo,
      STORENAME: item.LABLE1,
      Month: '',
      ModuleId: '72',
      ModuleRef: 'ET',
      state: 1,
      indexval: i,
    };
  }
  ExcelStoreNames: any = [];
  exportAsXLSX(): void {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Enterprise Income Expense(Dept)');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 13, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A14', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow([this.SetTitle]);
    titleRow.eachCell((cell, number) => {
      cell.alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');

    worksheet.addRow('');
    const DateToday = this.datepipe.transform(
      new Date(),
      'MM.dd.yyyy h:mm:ss a'
    );
    worksheet.addRow([DateToday]).font = {
      name: 'Arial',
      family: 4,
      size: 9,
    };
    const IncomeSummaryData = this.IncomeSummaryData.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const MonthDate = this.datepipe.transform(this.currentMonth, 'MMM yyyy');
    const Header = [MonthDate, 'Total', 'Average'];
    for (let i = 0; i < this.IncomeSummaryDataKeys.length; i++) {
      Header.push(this.formatKey(this.IncomeSummaryDataKeys[i]));
    }
    const ReportFilter = worksheet.addRow(['Report Filters :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const SummaryType = worksheet.addRow(['Summary Type :']);
    const summarytype = worksheet.getCell('B6');
    summarytype.value = 'Store Summary';
    summarytype.font = { name: 'Arial', family: 4, size: 9 };
    summarytype.alignment = { vertical: 'top', horizontal: 'left' };
    SummaryType.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const PresentMonthDate = this.datepipe.transform(
      this.currentMonth,
      'MM - dd - yyyy'
    );
    const DateMonth = worksheet.addRow(['Month :']);
    const datemonth = worksheet.getCell('B7');
    datemonth.value = PresentMonthDate;
    datemonth.font = { name: 'Arial', family: 4, size: 9 };
    datemonth.alignment = { vertical: 'top', horizontal: 'left' };
    DateMonth.getCell(1).font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    const Groups = worksheet.getCell('A8');
    Groups.value = 'Group :';
    Groups.font = { name: 'Arial', family: 4, size: 9, bold: true };
    const groups = worksheet.getCell('B8');
    groups.value = this.groupName;
    groups.font = { name: 'Arial', family: 4, size: 9 };
    groups.alignment = {
      vertical: 'middle',
      horizontal: 'left',
      wrapText: true,
    };
    let storeValue = '';

    const selectedStoreIds: string[] =
      this.storeIds && this.storeIds.length
        ? this.storeIds.map((id: any) => id.toString())
        : [];

    const allStores: any[] = Array.isArray(this.stores) ? this.stores : [];

    // ✅ Bind ONLY selected store names
    storeValue = allStores
      .filter((s: any) => selectedStoreIds.includes(s.ID.toString()))
      .map((s: any) => s.storename.trim())
      .filter(Boolean)
      .join(', ');

    // ✅ Final fallback (safety)
    if (!storeValue && selectedStoreIds.length) {
      storeValue = selectedStoreIds.join(', ');
    }
    worksheet.mergeCells('B9', 'K11');
    const Stores = worksheet.getCell('A9');
    Stores.value = 'Stores :';
    const stores = worksheet.getCell('B9');
    stores.value = storeValue;
    stores.font = { name: 'Arial', family: 4, size: 9 };
    stores.alignment = {
      vertical: 'top',
      horizontal: 'left',
      wrapText: true,
    };
    Stores.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
    };
    worksheet.addRow('');
    const headerRow = worksheet.addRow(Header);
    console.log(headerRow);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = {
      indent: 1,
      vertical: 'middle',
      horizontal: 'center',
    };
    headerRow.height = 30;
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '337ab7' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = {
        vertical: 'top',
        horizontal: 'center',
        wrapText: true,
      };
    });
    IncomeSummaryData.forEach((d: any) => {
      let Obj = [d.DISPLAY_LABLE, d.TOTAL, d.AVG];

      this.IncomeSummaryDataKeys.forEach((e: any) => {
        Obj.push(d[e] === '' || d[e] === null ? '-' : parseInt(d[e], 10));
      });

      const row = worksheet.addRow(Obj);

      // Cell alignment for the first column
      row.getCell(1).alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };

      if (row.number % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e5e5e5' },
            bgColor: { argb: 'FF0000FF' },
          };
        });
      }
      if (d.DISPLAYHEAD_FLAG === 1) {
        row.eachCell((cell, colNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '94b6d1' },
          };
        });
        row.eachCell((cell, colNumber) => {
          if (colNumber === 1) {
            cell.font = {
              name: 'Arial',
              family: 4,
              size: 9,
              bold: true,
              color: { argb: 'FFFFFF' },
            };
          } else {
            cell.value = null;
          }
          cell.font = {
            name: 'Arial',
            family: 4,
            size: 9,
            bold: true,
            color: { argb: 'FFFFFF' },
          };
        });
      }
      if (d.ISHEAD_TOTAL === 'Y') {
        row.eachCell((cell) => {
          cell.alignment = {
            indent: 2,
            vertical: 'top',
            horizontal: 'right',
          };
          cell.font = {
            name: 'Arial',
            family: 4,
            size: 8,
            bold: true,
          };
        });
      }
      if (d.ISHEAD_TOTAL === 'N') {
        row.eachCell((cell) => {
          cell.alignment = {
            indent: 3,
            vertical: 'top',
            horizontal: 'right',
          };
        });
      } else {
        // Default alignment and font properties for all cells
        row.eachCell((cell) => {
          cell.border = { right: { style: 'thin' } };
          cell.alignment = {
            vertical: 'top',
            horizontal: 'right',
            indent: 1,
          };
        });
      }

      // Font settings for all cells
      row.font = { name: 'Arial', family: 4, size: 8 };

      // Border and alignment settings for all cells
      row.eachCell((cell) => {
        cell.border = { right: { style: 'thin' } };
        cell.alignment = {
          vertical: 'top',
          horizontal: 'right',
          indent: 1,
        };
      });

      // Conditional number formatting based on DISPLAY_LABEL
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber >= 5 && (colNumber - 5) % 2 === 0) {
          // Apply percentage formatting to columns 5, 7, 9, 11, etc.
          cell.numFmt = '0%';
        } else if (colNumber >= 5 && (colNumber - 5) % 2 !== 0) {
          // Apply currency formatting to columns 6, 8, 10, 12, etc.
          cell.numFmt = '$#,##0';
        }
      });
      if (d.DISPLAY_LABLE && d.DISPLAY_LABLE.includes('%')) {
        row.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
          if (rowNumber > 1) {
            cell.numFmt = '0%';
          }
        });
      } else {
        row.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
          if (rowNumber > 1) {
            cell.numFmt = '$#,##0';
          }
        });
      }
    });
    worksheet.getColumn(1).width = 30;
    worksheet.getColumn(1).alignment = {
      indent: 1,
      vertical: 'middle',
      horizontal: 'left',
    };

    for (let i = 2; i <= 28; i++) {
      worksheet.getColumn(i).width = 15;
    }
    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, this.SetTitle + EXCEL_EXTENSION);
    });
  }



  generatePDF() {
    this.spinner.show();
    const printContents = document.getElementById('fixedexpensetrend')!.innerHTML;
    const iframe = document.createElement('iframe');
    iframe.style.position = 'absolute';
    iframe.style.width = '0px';
    iframe.style.height = '0px';
    iframe.style.border = 'none';

    document.body.appendChild(iframe);

    const doc = iframe.contentWindow?.document;
    if (!doc) {
      console.error('Failed to create iframe document');
      return;
    }

    doc.open();
    doc.write(`
    <html>
    <head>
    <title>${this.SetTitle}</title>
    <style>
    @font-face {
     font-family: 'GothamBookRegular';
     src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
          url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
   }
   @font-face {
     font-family: 'Roboto';
     src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
          url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
   }
   @font-face {
     font-family: 'RobotoBold';
     src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
          url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
   }

   .performance-scorecard{
     display: flex;
     flex-direction: column;
     max-width: fit-content;
   }
   .performance-scorecard .table>:not(:first-child){
     border-top:0px solid #ffa51a
   } 
   .performance-scorecard .table{
     text-align: center;
     text-transform: capitalize;
     border: transparent;
     width: min-content;
   }   
   .performance-scorecard .table th,  .performance-scorecard .table td {white-space: nowrap;vertical-align: top;}
   .performance-scorecard .table th:first-child, th:first-child table  td:first-child{
     position: sticky;
     left: 0;
     z-index: 1;
     background-color: #337ab7;
   } 
   .performance-scorecard .table tr:nth-child(odd) td:first-child,
   .performance-scorecard .table tr:nth-child(odd) td:nth-child(2){
     background-color: #e9ecef;
   }
   .performance-scorecard .table tr:nth-child(even) td:first-child,
   .performance-scorecard .table tr:nth-child(even) d:nth-child(2){
     background-color: #ffffff;
   }
   .performance-scorecard .table tr:nth-child(odd){
     background-color: #e9ecef;
 
   } 
   .performance-scorecard .table tr:nth-child(even){
     background-color: #ffffff;
 } 
 .performance-scorecard .table .spacer {
   // width: 50px !important;
   background-color: #cfd6de  !important;
   border-left: 1px solid #cfd6de !important;
   border-bottom: 1px solid #cfd6de !important;
   border-top: 1px solid #cfd6de !important;
} 
.performance-scorecard .table .bdr-rt{border-right: 1px solid #abd0ec;}
.performance-scorecard .table thead{
 position: sticky;
 top: 0;
 z-index: 99;
 font-family: 'FaktPro-Bold';
 font-size: 0.8rem;
} 
.performance-scorecard .table thead th{padding: 5px 10px;margin: 0px;width: 200px;white-space: normal; } 
.performance-scorecard .table thead tr:nth-child(1){ background-color: #337ab7 !important;color: #fff;text-transform: uppercase; border-bottom: #cfd6de;box-shadow: inset 0 1px 0 0 #cfd6de;}    
.performance-scorecard .table thead tr:nth-child(3) { background-color: #337ab7 !important;color: #fff;text-transform: uppercase; border-bottom: #cfd6de;box-shadow: inset 0 1px 0 0 #cfd6de;} 
.performance-scorecard .table tbody{
 font-family: 'FaktPro-Normal';
 font-size: .9rem;
}
.performance-scorecard .table tbody td:not(:first-child){text-align: center;}
.performance-scorecard .table tbody td{padding:2px 10px;margin: 0px; border: 1px solid #cfd6de }
.performance-scorecard .table tbody tr{border-bottom: 1px solid #cfd6de;border-left: 1px solid #cfd6de}
.performance-scorecard .table tbody td:first-child{text-align: start;box-shadow:inset -1px 0 0 0 #cfd6de ;}
.performance-scorecard .table tbody .text-bold{font-family: 'FaktPro-Bold';}
.performance-scorecard .table tbody .darkred-bg{ background-color: #282828 !important;color: #fff; }
.performance-scorecard .table tbody .lightblue-bg{ background-color: #94b6d1 !important;color: transparent !important;}
.performance-scorecard .table tbody .lightblueCol-bg{ background-color: #94b6d1 !important;color: #000 !important;padding-left: 1rem !important;}
.performance-scorecard .table tbody .gold-bg{ background-color: #ffa51a;color: #fff;}
.performance-scorecard .table tbody .extra-padding{
 padding-left: 2rem;
}
.performance-scorecard .table tbody .padding{
 padding-left: 1rem;
}
.performance-scorecard .table tbody .LayerOne{
 padding-left: 3.5rem;
} 
.performance-scorecard .table tbody .LayerTwo{
 padding-left: 5rem;
}
.divBold {
 font-family: "RobotoBold" !important;
} 
 </style>
  </head>
  <body id='content'>
        ${printContents}
        </body>
          </html>`);
    doc.close();

    const div = doc.getElementById('content');
    const options = {
      logging: true,
      allowTaint: false,
      useCORS: true,
    };
    if (!div) {
      console.error('Element not found');
      return;
    }
    html2canvas(div, options)
      .then((canvas) => {
        let imgWidth = 286;
        let pageHeight = 208;
        let imgHeight = (canvas.height * imgWidth) / canvas.width;
        let heightLeft = imgHeight;
        const contentDataURL = canvas.toDataURL('image/png');
        let pdfData = new jsPDF('l', 'mm', 'a4', true);
        let position = 5;

        function addExtraDataToPage(pdf: any, extraData: any, positionY: any) {
          pdf.text(extraData, 10, positionY);
        }

        function addPageAndImage(pdf: any, contentDataURL: any, position: any) {
          pdf.addPage();
          pdf.addImage(
            contentDataURL,
            'PNG',
            5,
            position,
            imgWidth,
            imgHeight,
            undefined,
            'FAST'
          );
        }

        pdfData.addImage(
          contentDataURL,
          'PNG',
          5,
          position,
          imgWidth,
          imgHeight,
          undefined,
          'FAST'
        );
        addExtraDataToPage(pdfData, '', position + imgHeight + 10);
        heightLeft -= pageHeight;

        while (heightLeft >= 0) {
          position = heightLeft - imgHeight;
          addPageAndImage(pdfData, contentDataURL, position);
          addExtraDataToPage(pdfData, '', position + imgHeight + 10);
          heightLeft -= pageHeight;
        }

        return pdfData;
      })
      .then((doc) => {
        doc.save(`${this.SetTitle}.pdf`);
        // popupWin.close();
        this.spinner.hide();
      });
  }
  sendEmailData(Email: any, notes: any, from: any) {
    this.spinner.show();
    const printContents = document.getElementById('fixedexpensetrend')!.innerHTML;
    const iframe = document.createElement('iframe');

    // Make the iframe invisible
    iframe.style.position = 'absolute';
    iframe.style.width = '0px';
    iframe.style.height = '0px';
    iframe.style.border = 'none';

    document.body.appendChild(iframe);

    const doc = iframe.contentWindow?.document;
    if (!doc) {
      console.error('Failed to create iframe document');
      return;
    }

    doc.open();
    doc.write(`
        <html>
            <head>
            <title>${this.SetTitle}</title>
            <style>
            @font-face {
             font-family: 'GothamBookRegular';
             src: url('assets/fonts/Gotham\ Book\ Regular.otf') format('otf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                  url('assets/fonts/Gotham\ Book\ Regular.otf') format('opentype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
           }
           @font-face {
             font-family: 'Roboto';
             src: url('assets/fonts/Roboto-Regular.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                  url('assets/fonts/Roboto-Regular.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
           }
           @font-face {
             font-family: 'RobotoBold';
             src: url('assets/fonts/Roboto-Bold.ttf') format('ttf'), /* Chrome 6+, Firefox 3.6+, IE 9+, Safari 5.1+ */
                  url('assets/fonts/Roboto-Bold.ttf') format('truetype'); /* Chrome 4+, Firefox 3.5, Opera 10+, Safari 3—5 */
           }

           .performance-scorecard{
             display: flex;
             flex-direction: column;
             max-width: fit-content;
           }
           .performance-scorecard .table>:not(:first-child){
             border-top:0px solid #ffa51a
           } 
           .performance-scorecard .table{
             text-align: center;
             text-transform: capitalize;
             border: transparent;
             width: min-content;
           }   
           .performance-scorecard .table th,  .performance-scorecard .table td {white-space: nowrap;vertical-align: top;}
           .performance-scorecard .table th:first-child, th:first-child table  td:first-child{
             position: sticky;
             left: 0;
             z-index: 1;
             background-color: #337ab7;
           } 
           .performance-scorecard .table tr:nth-child(odd) td:first-child,
           .performance-scorecard .table tr:nth-child(odd) td:nth-child(2){
             background-color: #e9ecef;
           }
           .performance-scorecard .table tr:nth-child(even) td:first-child,
           .performance-scorecard .table tr:nth-child(even) d:nth-child(2){
             background-color: #ffffff;
           }
           .performance-scorecard .table tr:nth-child(odd){
             background-color: #e9ecef;
         
           } 
           .performance-scorecard .table tr:nth-child(even){
             background-color: #ffffff;
         } 
         .performance-scorecard .table .spacer {
           // width: 50px !important;
           background-color: #cfd6de  !important;
           border-left: 1px solid #cfd6de !important;
           border-bottom: 1px solid #cfd6de !important;
           border-top: 1px solid #cfd6de !important;
       } 
       .performance-scorecard .table .bdr-rt{border-right: 1px solid #abd0ec;}
       .performance-scorecard .table thead{
         position: sticky;
         top: 0;
         z-index: 99;
         font-family: 'FaktPro-Bold';
         font-size: 0.8rem;
       } 
       .performance-scorecard .table thead th{padding: 5px 10px;margin: 0px;width: 200px;white-space: normal; } 
       .performance-scorecard .table thead tr:nth-child(1){ background-color: #337ab7 !important;color: #fff;text-transform: uppercase; border-bottom: #cfd6de;box-shadow: inset 0 1px 0 0 #cfd6de;}    
       .performance-scorecard .table thead tr:nth-child(3) { background-color: #337ab7 !important;color: #fff;text-transform: uppercase; border-bottom: #cfd6de;box-shadow: inset 0 1px 0 0 #cfd6de;} 
       .performance-scorecard .table tbody{
         font-family: 'FaktPro-Normal';
         font-size: .9rem;
       }
       .performance-scorecard .table tbody td:not(:first-child){text-align: center;}
       .performance-scorecard .table tbody td{padding:2px 10px;margin: 0px; border: 1px solid #cfd6de }
       .performance-scorecard .table tbody tr{border-bottom: 1px solid #cfd6de;border-left: 1px solid #cfd6de}
       .performance-scorecard .table tbody td:first-child{text-align: start;box-shadow:inset -1px 0 0 0 #cfd6de ;}
       .performance-scorecard .table tbody .text-bold{font-family: 'FaktPro-Bold';}
       .performance-scorecard .table tbody .darkred-bg{ background-color: #282828 !important;color: #fff; }
       .performance-scorecard .table tbody .lightblue-bg{ background-color: #94b6d1 !important;color: transparent !important;}
       .performance-scorecard .table tbody .lightblueCol-bg{ background-color: #94b6d1 !important;color: #000 !important;padding-left: 1rem !important;}
       .performance-scorecard .table tbody .gold-bg{ background-color: #ffa51a;color: #fff;}
       .performance-scorecard .table tbody .extra-padding{
         padding-left: 2rem;
       }
       .performance-scorecard .table tbody .padding{
         padding-left: 1rem;
       }
       .performance-scorecard .table tbody .LayerOne{
         padding-left: 3.5rem;
       } 
       .performance-scorecard .table tbody .LayerTwo{
         padding-left: 5rem;
       }
       .divBold {
         font-family: "RobotoBold" !important;
       } 
         </style>
            </head>
            <body id='content'>
                ${printContents}
            </body>
        </html>
    `);

    doc.close();

    const div = doc.getElementById('content');
    if (!div) {
      console.error('Element not found');
      return;
    }

    const options = {
      logging: true,
      allowTaint: false,
      useCORS: true,
      scale: 1 // Adjust scale to fit the page better
    };

    html2canvas(div, options)
      .then((canvas) => {
        let imgWidth = 285;
        let pageHeight = 204;
        let imgHeight = (canvas.height * imgWidth) / canvas.width;
        let heightLeft = imgHeight;
        const contentDataURL = canvas.toDataURL('image/png');
        let pdfData = new jsPDF('l', 'mm', 'a4', true);
        let position = 5;
        pdfData.addImage(
          contentDataURL,
          'PNG',
          5,
          position,
          imgWidth,
          imgHeight,
          undefined,
          'FAST'
        );
        heightLeft -= pageHeight;
        while (heightLeft >= 0) {
          position = heightLeft - imgHeight;
          pdfData.addPage();
          pdfData.addImage(
            contentDataURL,
            'PNG',
            5,
            position,
            imgWidth,
            imgHeight,
            undefined,
            'FAST'
          );
          heightLeft -= pageHeight;
        }

        const pdfBlob = pdfData.output('blob');
        const pdfFile = this.blobToFile(pdfBlob, `${this.SetTitle}.pdf`);
        const formData = new FormData();
        formData.append('to_email', Email);
        formData.append('subject', this.SetTitle);
        formData.append('file', pdfFile);
        formData.append('notes', notes);
        formData.append('from', from);
        this.apiSrvc.postmethod(this.comm.routeEndpoint + 'mail', formData).subscribe(
          (res: any) => {
            console.log('Response:', res);
            if (res.status === 200) {

              this.toast.success(res.response)
            } else {

              this.toast.show('Invalid Details.', 'danger', 'Error');
            }
          },
          (error) => {
            console.error('Error:', error);
          }
        );
      })
      .catch((error) => {
        console.error('html2canvas error:', error);
      })
      .finally(() => {
        this.spinner.hide();
        // popupWin.close();
      });
  }
  public blobToFile = (theBlob: Blob, fileName: string): File => {
    return new File([theBlob], fileName, {
      lastModified: new Date().getTime(),
      type: theBlob.type,
    });
  };

  GetPrintData() {
    window.print();
  }

}
