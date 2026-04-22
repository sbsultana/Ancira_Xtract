import { Component, HostListener, Injector } from '@angular/core';
import { Api } from '../../../../Core/Providers/Api/api';
import { common } from '../../../../common';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { FormsModule } from '@angular/forms';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { NgxSpinnerService } from 'ngx-spinner';
import { CurrencyPipe, DatePipe, formatDate } from '@angular/common';
import { Title } from '@angular/platform-browser';
import { Subscription } from 'rxjs';
import { environment } from '../../../../../environments/environment';
import * as ExcelJS from 'exceljs';
import FileSaver from 'file-saver';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
const EXCEL_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
import numeral from 'numeral';
import { BsDatepickerConfig, BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Router } from '@angular/router';
import { EnterpriseincomebyexpenseDetails } from '../enterpriseincomebyexpense-details/enterpriseincomebyexpense-details';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  Current_Date: any;
  Report: any = '';

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
  ShowHideGP: any = 'Hide';
  ShowHideBudget: any = 'Hide';

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
      case 'FixedIncomeByExpense':
        this.filters = ['Service', 'Parts', 'Details'];
        this.SetTitle = 'Fixed Income / Expense';
        break;
      case 'VariableIncomeByExpense':
        this.filters = ['New', 'Used'];
        this.SetTitle = 'Variable Income / Expense';
        break;
      case 'EnterpriseIncomeByExpense':
        this.filters = ['New', 'Used', 'Service', 'Parts', 'Details'];
        this.SetTitle = 'Enterprise Income / Expense';
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
    // this.date= this.Month



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

  showhideGP(value: string) {
    this.ShowHideGP = value;
    this.activePopover = null;
  }
  showhideBudget(value: string) {
    this.ShowHideBudget = value;
    this.activePopover = null;
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
      DATE: this.shared.datePipe.transform(this.currentMonth, 'yyyy-MM') + '-10',
    };
    console.log(obj);

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetEnterpriseIncomeExpenseReport', obj)
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
  isPercent(label: string): boolean {
    return ['% of GP', 'Selling Gross %']
      .some(x => label?.includes(x));
  }
  isNetNegative(label: string, value: any): boolean {
    return ['Net Profit (Loss)', 'Operating Net'].includes(label)
      && !this.inTheGreen(value);
  }

  FixedExpenseLayerOne: any;
  FixedExpenseLayerOneKeys: any;
  GetFixedExpenseLayerOne(STR: any, index: number, ParentCode: any): void {
    this.spinner.show();
    console.log(STR);
    // STR.expanded = !STR.expanded;
    this.FixedExpenseLayerOne = [];
    this.FixedExpenseLayerOneKeys = [];
    const obj = {
      LEVEL: '1',
      AS_IDs: this.storeIds.toString(),
      STORENAMES: '',
      DEPARTMENT: this.selectedFilters.toString(),
      DATE: '10-' + this.shared.datePipe.transform(this.currentMonth, 'MMM-yyyy'),
      SUBTYPEDETAIL: STR.DISPLAY_LABLE,
      ACCTTYPEDETAIL: '',
      LABLECODE: STR.PARENTLABLECODE,
      ACCTNUM: '',
      ACCTDESC: '',
      Control: '',
    };
    console.log(obj);

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetEnterpriseIncomeExpenseDetailByLevel', obj)
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
    const obj = {
      LEVEL: '2',
      AS_IDs: this.storeIds.toString(),
      STORENAMES: '',
      DEPARTMENT: this.selectedFilters.toString(),
      DATE: '10-' + this.shared.datePipe.transform(this.currentMonth, 'MMM-yyyy'),
      SUBTYPEDETAIL: STR.DISPLAY_LABLE,
      ACCTTYPEDETAIL: '',
      LABLECODE: STR.PARENTLABLECODE,
      ACCTNUM: Layer.AcctNum,
      ACCTDESC: Layer.Desc,
      Control: '',
    };
    console.log(obj);

    this.apiSrvc
      .postmethod(this.comm.routeEndpoint + 'GetEnterpriseIncomeExpenseDetailByLevel', obj)
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

  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
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
    this.excel = this.apiSrvc.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== this.SetTitle) return;
      this.SFstate = obj.state;
      if (obj.state) {
        this.exportAsXLSX();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== this.SetTitle) return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== this.SetTitle) return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== this.SetTitle) return;
      if (obj.stateEmailPdf) {
        this.sendEmailData(obj.Email, obj.notes, obj.from);
      }
    });
  }

  ngOnDestroy(): void {
    this.excel?.unsubscribe();
    this.pdf?.unsubscribe();
    this.print?.unsubscribe();
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
    const DetailsET = this.ngbmodal.open(EnterpriseincomebyexpenseDetails, {
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
    const DetailsET = this.ngbmodal.open(EnterpriseincomebyexpenseDetails, {
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

  generatePDF() {
    if (this.storeIds.length >= 16) {
      this.toast.show('Please select only up to 15 stores', 'warning', 'Warning');
      return;
    }

    this.shared.spinner.show();

    try {
      const doc = this.createPDF();

      const safeTitle = (this.SetTitle || 'Report');

      doc.save(`${safeTitle}.pdf`);

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
    if (this.storeIds.length >= 16) {
      this.toast.show('Please select only up to 15 stores', 'warning', 'Warning');
      return;
    }
    this.spinner.show();
    const safeTitle = (this.SetTitle || 'Report');
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
      const pdfFile = this.blobToFile(pdfBlob, `${safeTitle}.pdf`);
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', safeTitle);
      formData.append('file', pdfFile);
      formData.append('notes', notes);
      formData.append('from', from);

      this.apiSrvc.postmethod(this.comm.routeEndpoint + 'mail', formData)
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
    if (this.storeIds.length >= 16) {
      this.toast.show('Please select only up to 15 stores', 'warning', 'Warning');
      return;
    }
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

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a2' });

    /* ================= TITLE ================= */
    doc.setFontSize(14);
    doc.text(this.SetTitle, 14, 12);

    let startY = 18;

    /* ================= FILTERS (same as Excel) ================= */
    // const filtersObj = this.getReportFilters();

    doc.setFontSize(10);
    // filtersObj.filters.forEach((f: any) => {
    //   doc.text(`${f.label}: ${f.value}`, 14, startY);
    //   startY += 5;
    // });

    // startY += 3;

    /* ================= HEADER ================= */
    const MonthDate = this.datepipe.transform(this.currentMonth, 'MMM yyyy');

    const head = [[
      MonthDate,
      'Total',
      'Average',
      ...this.IncomeSummaryDataKeys.map((k: string) => this.formatKey(k))
    ]];

    /* ================= BODY ================= */
    const body: any[] = [];

    this.IncomeSummaryData.forEach((d: any) => {

      let row = [
        d.DISPLAY_LABLE,
        d.TOTAL,
        d.AVG
      ];

      this.IncomeSummaryDataKeys.forEach((k: any) => {
        row.push(d[k] === '' || d[k] === null ? '-' : Number(d[k]));
      });

      body.push({
        rowData: d,
        values: row
      });
    });

    /* ================= TABLE ================= */
    autoTable(doc, {
      startY: startY,
      head: head,
      body: body.map(r => r.values),
      theme: 'grid',

      styles: {
        fontSize: 10,
        cellPadding: 2,
        lineWidth: 0.1,
        lineColor: [200, 200, 200],
        textColor: [20, 20, 20], // 👈 slightly richer black
        valign: 'middle'
      },

      /* ✅ COLUMN WIDTH CONTROL */
      columnStyles: {
        0: { cellWidth: 50 }, // 1st column (Label)
        1: { cellWidth: 30 }, // 2nd column
        // remaining columns → auto (no need to define)
      },

      didParseCell: (data: any) => {

        const rowObj = body[data.row.index];
        const rowData = rowObj?.rowData;

        /* ================= HEADER ================= */
        if (data.section === 'head') {
          data.cell.styles.fillColor = [5, 84, 239]; // Excel blue
          data.cell.styles.textColor = [255, 255, 255];
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.halign = 'center';
          return;
        }

        if (!rowData) return;

        const isHeaderRow = rowData.DISPLAYHEAD_FLAG === 1;
        const isTotalRow = rowData.ISHEAD_TOTAL === 'Y';

        /* ================= ALIGN ================= */
        data.cell.styles.halign =
          data.column.index === 0 ? 'left' : 'right';

        /* ================= EMPTY LOGIC ================= */
        const valEmpty =
          data.cell.raw === null ||
          data.cell.raw === undefined ||
          data.cell.raw === '' ||
          data.cell.raw === 0;

        if (isHeaderRow) {
          if (valEmpty) {
            data.cell.text = [''];
            return;
          }
        } else {
          if (valEmpty) {
            data.cell.text = ['-'];
            data.cell.styles.halign = 'center';
            return;
          }
        }

        /* ================= NUMBER FORMAT ================= */
        if (data.column.index > 0) {

          const rawValue = body[data.row.index]?.values[data.column.index];

          if (rawValue === '-' || rawValue === null || rawValue === undefined || rawValue === '') {
            return;
          }

          let val = Number(rawValue);

          if (!isNaN(val)) {

            /* ✅ % VALUES (keep decimals) */
            if (rowData.DISPLAY_LABLE?.includes('% of GP') ||
              rowData.DISPLAY_LABLE?.includes('Selling Gross %')) {

              const formatted = val.toLocaleString(undefined, {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
              });

              data.cell.text = [`${formatted}%`];
              data.cell.styles.halign = 'right';

            } else {

              /* ✅ $ VALUES (ROUND FIGURE) */
              const rounded = Math.round(val);

              const formatted = rounded.toLocaleString(undefined, {
                minimumFractionDigits: 0,
                maximumFractionDigits: 0
              });

              data.cell.text = [`$ ${formatted}`];

              if (rounded < 0) {
                data.cell.styles.textColor = [220, 53, 69];
              }
            }
          }
        }
        /* ================= SECTION HEADER ROW ================= */
        if (isHeaderRow) {
          Object.values(data.row.cells).forEach((cell: any) => {
            cell.styles.fillColor = [217, 231, 255]; // light blue
            cell.styles.fontStyle = 'bold';
          });
        }

        /* ================= TOTAL ROW ================= */
        if (isTotalRow) {
          Object.values(data.row.cells).forEach((cell: any) => {
            cell.styles.fontStyle = 'bold';
            cell.styles.fillColor = null;
          });
        }

        /* ================= ZEBRA ================= */
        if (!isHeaderRow && data.row.index % 2 === 0) {
          Object.values(data.row.cells).forEach((cell: any) => {
            cell.styles.fillColor = [245, 247, 250];
          });
        }
      }
    });

    return doc;
  }

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
      title: this.SetTitle, // dynamic title from route
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
          label: 'Department',
          value: this.selectedFilters && this.selectedFilters.length
            ? this.selectedFilters.join(', ')
            : 'All'
        },
        {
          label: 'Month',
          value: this.datepipe.transform(this.selectDate, 'MMMM yyyy')
        },
        {
          label: '% of GP',
          value: this.ShowHideGP === 'Show' ? 'Yes' : 'No'
        },
        {
          label: 'Budget',
          value: this.ShowHideBudget === 'Show' ? 'Yes' : 'No'
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
  ExcelStoreNames: any = [];
  exportAsXLSX(): void {

    const workbook = new ExcelJS.Workbook();
    let SetTitle = this.SetTitle.replace(/[\/()]/g, '').trim();
    const worksheet = workbook.addWorksheet(SetTitle);

    /* ================= FILTER SECTION ================= */
    const filterRowCount = this.addExcelFiltersSection(worksheet);

    /* ================= FREEZE ================= */
    worksheet.views = [{
      state: 'frozen',
      ySplit: filterRowCount + 1
    }];

    const IncomeSummaryData = this.IncomeSummaryData.map((e: any) => ({ ...e }));

    const MonthDate = this.datepipe.transform(this.currentMonth, 'MMM yyyy');

    const Header = [MonthDate, 'Total', 'Average'];
    for (let i = 0; i < this.IncomeSummaryDataKeys.length; i++) {
      Header.push(this.formatKey(this.IncomeSummaryDataKeys[i]));
    }

    /* ================= HEADER (LIGHT BLUE) ================= */
    const headerRow = worksheet.addRow(Header);

    headerRow.eachCell((cell: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0554EF' }

      };

      cell.font = {
        bold: true,
        name: 'Calibri',
        color: { argb: 'FFFFFFFF' }

      };

      cell.alignment = { horizontal: 'center', vertical: 'middle' };

      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    /* ================= DATA ================= */
    IncomeSummaryData.forEach((d: any) => {

      let Obj = [d.DISPLAY_LABLE, d.TOTAL, d.AVG];

      this.IncomeSummaryDataKeys.forEach((e: any) => {
        Obj.push(d[e] === '' || d[e] === null ? '-' : Number(d[e]));
      });

      const row = worksheet.addRow(Obj);

      row.eachCell((cell: any, colNumber: number) => {

        cell.font = { name: 'Calibri', size: 11 };

        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };

        // Alignment
        if (colNumber === 1) {
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        } else {
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        }


        // Accounting format (NO brackets, only minus)
        if (colNumber >= 2 && typeof cell.value === 'number') {

          /* ✅ % ROW */
          if (d.DISPLAY_LABLE?.includes('% of GP') ||
            d.DISPLAY_LABLE?.includes('Selling Gross %')) {

            cell.numFmt = '0.00"%"';

          } else {

            /* ✅ $ CURRENCY FORMAT */
            cell.numFmt = '_("$"* #,##0_);[Red]_("$"* -#,##0_);_("$"* "-"??_);_(@_)';

            if (cell.value < 0) {
              cell.font = {
                ...cell.font,
                color: { argb: 'FFFF0000' }
              };
            }
          }
        }

        // ❗ Skip '-' for header rows
        if (d.DISPLAYHEAD_FLAG !== 1) {
          if (!cell.value || cell.value === '-') {
            cell.value = '-';
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
          }
        } else {
          // Header rows → keep empty
          if (!cell.value || cell.value === '-') {
            cell.value = '';
          }
        }
      });

      /* Alternate Row Color */
      if (row.number % 2 === 0) {
        row.eachCell((cell: any) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF5F7FA' }
          };
        });
      }

      /* ===== HEAD ROW ===== */
      if (d.DISPLAYHEAD_FLAG === 1) {
        row.eachCell((cell: any) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'D9E7FF' }
          };
          cell.font = {
            bold: true,
            color: { argb: 'FF000000' }
          };
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'D9E7FF' }
          };
        });
      }

      /* ===== TOTAL ROW ===== */
      if (d.ISHEAD_TOTAL === 'Y') {
        row.eachCell((cell: any) => {
          cell.font = {
            name: 'Calibri',
            size: 11,
            bold: true
          };

          cell.fill = undefined;
          cell.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
          };
        });
      }
    });

    /* ================= COLUMN WIDTH ================= */
    worksheet.getColumn(1).width = 30;

    for (let i = 2; i <= Header.length; i++) {
      worksheet.getColumn(i).width = 32;
    }

    /* ================= EXPORT ================= */
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      FileSaver.saveAs(blob, this.SetTitle + EXCEL_EXTENSION);
    });
  }

  GetPrintData() {
    window.print();
  }

}
