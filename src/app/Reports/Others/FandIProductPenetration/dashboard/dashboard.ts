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
import { CurrencyPipe } from '@angular/common';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {
  Report: any = '';
  FromDate: any = '';
  ToDate: any = '';
  TotalReport: any = 'T';
  StoreIds = 2;
  Headings: any = [];
  salesPersonId: any = '0';
  salesManagerId: any = '0';
  financeManagerId: any = '0';
  responcestatus: any = '';
  path1: any = 'Store_Name';
  path2: any = 'FIMANAGER';
  Filter: any = 'FIMANAGER';
  path1name: any = 'Store';
  path2name: any = 'F&I Manager';

  path1id: any = 40;
  path2id: any = 26;

  groups: any = 1;
  bsValue = new Date();
  bsRangeValue!: Date[];
  toporbottom: any = ['T'];

  @ViewChild('content', { static: false }) content!: ElementRef;
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;
  Pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  storeName: any;
  storeIds: any = '';
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  stores: any = []
  selectedProductType: any = []
  selectedProductCategory: any = []

  //  storeIds: any = '';
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'N',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname
  };
  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService,
  ) {
    localStorage.setItem('time', 'MTD');
    this.initializeDates('MTD')

    this.shared.setTitle(this.comm.titleName + '-F & I Product Penetration');
    // this.StoreIds=environment.stores .map(function (a:any) {
    //   return a.AS_ID;
    // }).toString();
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      // this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.ustores.split(',')
      // this.storeIds = '1,2';
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
    }


    if (localStorage.getItem('Fav') != 'Y') {

      this.getDataGroupings()

      this.setDates(this.DateType);
      this.setHeaderData()
      this.getFandIPPData();
      this.getProductType()
      this.getProductCategory()
    } else {

    }
    if (this.storeIds != '') {
      this.getEmployees();
    }
  }

  ngOnInit(): void {
    this.Headings = [
      'Service Contract',
      'Gap',
      'VTR',
      'Insurance',
      'Chemical',
      'Prepaid Maintenance',
    ];
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
  setHeaderData() {
    const data = {
      title: 'F & I Product Penetration',
      path1: 'Store',
      path2: 'F&I Manager',
      path3: '',
      path1id: 40,
      path2id: '26',
      stores: this.storeIds,
      toporbottom: this.TotalReport,
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groups,
      financemanagers: this.financeManagerId,


    };
    this.shared.api.SetHeaderData({ obj: data });
  }
  multipleorsingle(block: any, e: any) {

    if (block == 'RT') {
      this.TotalReport = e;
    }

  }


  Scrollpercent: any = 0;
  updateVerticalScroll(event: any): void {
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );

    // //console.log(
    //   event.target.scrollTop,
    //   event.document,
    //   ((document.documentElement.scrollTop + document.body.scrollTop) /
    //     (document.documentElement.scrollHeight -
    //       document.documentElement.clientHeight)) *
    //     100
    // );

    // var el = document.getElementById('scrollcent');
    // var top = el.offsetTop;
    // var height = el.offsetHeight;
    // var scrollTop = document.body.scrollTop;
    // var difference = el.scrollTop - top;

    // var percent = difference / height;
    // //console.log(el.scrollTop, el.offsetHeight, percent, scrollTop);
  }


  selBlock: any;
  screenheight: any = 0; divheight: any = 0; trposition: any = 0;

  index = '';
  commentobj: any = {};



  closeReport() {
    this.Report = '';
  }

  FandIPPData: any = [];
  TotalData: any;
  NoData: boolean = false;
  NodataFound: boolean = false;

  producttype: any = []
  getProductType() {
    const obj = {}
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetProductType', obj).subscribe((res: any) => {
      if (res.status == 200) {
        if (res.response) {
          this.producttype = res.response
          this.selectedProductType = [...res.response]
        }
      }
    })
  }

  productcategory: any = []
  getProductCategory() {
    const obj = {}
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetProductCategory', obj).subscribe((res: any) => {
      if (res.status == 200) {
        if (res.response) {
          this.productcategory = res.response;
          this.selectedProductCategory = [...res.response]
        }
      }
    })
  }

  SelectedPTorPC(data: any, block: any) {
    if (block == 'PT') {
      const index = this.selectedProductType.findIndex((i: any) => i == data);
      if (index >= 0) {
        this.selectedProductType.splice(index, 1);
        if (this.selectedProductType.length == 0) {

          this.toast.show('Please select atleast on Product Type.', 'danger', 'Error');
        }
      } else {
        this.selectedProductType.push(data);
      }
    }
    if (block == 'PC') {
      const index = this.selectedProductCategory.findIndex((i: any) => i == data);
      if (index >= 0) {
        this.selectedProductCategory.splice(index, 1);
        if (this.selectedProductCategory.length == 0) {

          this.toast.show('Please select atleast on Product Category.', 'danger', 'Error');
        }
      } else {
        this.selectedProductCategory.push(data);
      }
    }


  }
  getFandIPPData() {
    this.responcestatus = '';
    this.shared.spinner.show();
    this.GetData();
    this.GetTotalData();
  }
  GetData() {
    this.TotalData = [];

    this.shared.spinner.show();
    const obj = {
      StartDate: this.FromDate.replace(/-/g, '/'),
      EndDate: this.ToDate.replace(/-/g, '/'),
      Stores: this.storeIds.toString(),
      Type: "D",
      SalesPerson: this.salesPersonId,
      SalesManager: this.salesManagerId,
      FinanceManager: this.financeManager.length == this.financeManagerId.length ? 0 : this.financeManagerId.toString(),
      ProductType: this.selectedProductType && this.selectedProductType.length > 0 ? this.selectedProductType.map(function (a: any) { return a.TID; }).toString() : 0,
      ProductCategory: this.selectedProductCategory && this.selectedProductCategory.length > 0 ? this.selectedProductCategory.map(function (a: any) { return a.CID; }).toString() : 0,
      var1: this.selectedDataGrouping && this.selectedDataGrouping.length > 0 ? this.selectedDataGrouping[0].columnname : '',
      var2: this.selectedDataGrouping && this.selectedDataGrouping.length == 2 ? this.selectedDataGrouping[1].columnname : '',
    };
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetFandIProductPenetrationByCategoriesNew', obj)
      .subscribe(
        (res: { message: any; status: number; response: any[]; }) => {

          this.shared.spinner.hide();
          if (res.status == 200) {
            this.FandIPPData = [];
            this.FandIPPData = res.response.sort((a: any, b: any) => a.data1 > b.data1 ? 1 : -1);;
            // const index = this.FandIPPData.findIndex(
            //   (i:any) => i.StoreName == 'Total'
            // );
            // if (index >= 0) {
            //   this.TotalData.push(this.FandIPPData[index]);
            //   this.FandIPPData.splice(index, 1);
            // this.FandIPPData.push(this.TotalData[0]);

            // }
            if (this.FandIPPData != undefined) {
              if (this.FandIPPData.length > 0) {
                this.responcestatus = this.responcestatus + 'I';
                this.FandIPPData.some(function (x: any) {
                  x.Dealer = '-';

                  if (x.productdata != undefined) {
                    x.productdata = JSON.parse(x.productdata);
                    var heading = x.productdata.map(function (a: any) {
                      return a.ProductCategory;
                    });
                  }
                  if (x.data2 != undefined) {
                    x.data2 = JSON.parse(x.data2);

                  }
                });


                if (this.FandIPPData[0].productdata != undefined) {
                  this.Headings = this.FandIPPData[0].productdata.map(function (
                    a: any
                  ) {
                    return a.ProductCategory;
                  });
                }
                this.combineIndividualandTotal();
              }
              else {
                // this.shared.spinner.hide();
                // this.NoData = true;
              }
            } else {

              //  this.toast.error('Empty Response', '');
              // this.shared.spinner.hide();
              // this.NoData = true;
            }
          } else {

            // this.toast.error(res.status, '');

            this.shared.spinner.hide();
            // this.NoData = true;
          }
        },
        (error: any) => {
          // this.toast.error('502 Bad Gate Way Error', '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }
  TotalFandIPPData: any = []
  GetTotalData() {
    this.TotalFandIPPData = [];
    const obj = {
      StartDate: this.FromDate.replace(/-/g, '/'),
      EndDate: this.ToDate.replace(/-/g, '/'),
      Stores: this.storeIds.toString(),
      Type: "T",
      SalesPerson: this.salesPersonId,
      SalesManager: this.salesManagerId,
      FinanceManager: this.financeManagerId.toString(),
      // var1: this.path1 == 'Store_Name' ? '' : this.path1,
      // var2: this.path2
      ProductType: this.selectedProductType && this.selectedProductType.length > 0 ? this.selectedProductType.map(function (a: any) { return a.TID; }).toString() : 0,
      ProductCategory: this.selectedProductCategory && this.selectedProductCategory.length > 0 ? this.selectedProductCategory.map(function (a: any) { return a.CID; }).toString() : 0,
      var1: this.selectedDataGrouping && this.selectedDataGrouping.length == 1 ? (this.selectedDataGrouping[0].columnname == 'Store_Name' ? '' : this.selectedDataGrouping[0].columnname) : (this.selectedDataGrouping[0]?.columnname == 'Store_Name' ? '' : this.selectedDataGrouping[0]?.columnname),
      var2: this.selectedDataGrouping && this.selectedDataGrouping.length == 2 ? this.selectedDataGrouping[1].columnname : '',
    };
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetFandIProductPenetrationByCategoriesNew', obj)
      .subscribe(
        (totalres: { status: number; response: any; }) => {
          // console.log(totalres);

          if (totalres.status == 200) {
            this.TotalFandIPPData = totalres.response
            if (this.TotalFandIPPData != undefined) {
              if (this.TotalFandIPPData.length > 0) {
                this.responcestatus = this.responcestatus + 'T';
                this.TotalFandIPPData.some(function (x: any) {
                  x.data1 = 'REPORT TOTALS'
                  if (x.productdata != undefined) {
                    x.productdata = JSON.parse(x.productdata);
                    var heading = x.productdata.map(function (a: any) {
                      return a.ProductCategory;
                    });
                  }
                });

                if (this.TotalFandIPPData[0].productdata != undefined) {
                  this.Headings = this.TotalFandIPPData[0].productdata.map(function (
                    a: any
                  ) {
                    return a.ProductCategory;
                  });
                }
                this.combineIndividualandTotal();
              }
              else {
                // this.shared.spinner.hide();
                // this.NoData = true;
              }
            } else {
              // this.toast.error('Empty Response','');
              // this.shared.spinner.hide();
              // this.NoData = true;
            }
          } else {
            // this.toast.error(totalres.status, '');
            this.shared.spinner.hide();
            // this.NoData = true;
          }
        },
        (error: any) => {
          // this.toast.error('502 Bad Gate Way Error', '');
          this.shared.spinner.hide();
          // this.NoData = true;
        }
      );
  }

  combineIndividualandTotal() {
    if (this.responcestatus == 'IT' || this.responcestatus == 'TI') {
      if (this.TotalReport == 'B') {
        this.FandIPPData.push(this.TotalFandIPPData[0]);
      } else {
        this.FandIPPData.unshift(this.TotalFandIPPData[0]);
      }
      console.log(this.FandIPPData);

      this.shared.spinner.hide();
    }
    // else if (this.responcestatus == 'T') {
    //   this.FandIPPData = this.TotalFandIPPData;
    // }
    else if (this.responcestatus == 'I') {
      this.FandIPPData = this.FandIPPData;
    } else {
      // this.NoData = true;
    }

    if (this.FandIPPData.length < 1) {
      this.NoData = true;
      this.FandIPPData = [];
    } else {
      this.NoData = false;
    }
    this.shared.spinner.hide()
    // //console.log(this.FandIPPData);

  }

  isDesc: boolean = false;
  column: string = 'CategoryName';

  sort(property: any) {
    console.log(property);

    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;

    this.FandIPPData.sort(function (a: any, b: any) {
      // console.log(a.productdata[index].Products[0][property],b.productdata[index].Products[0][property]);      
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });

  }

  nestedsort(property: any, index: any) {
    console.log(property, index);

    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;

    this.FandIPPData.sort(function (a: any, b: any) {
      // console.log(a.productdata[index].Products[0][property],b.productdata[index].Products[0][property]);      
      if (a.productdata[index].Products[0][property] < b.productdata[index].Products[0][property]) {
        return -1 * direction;
      } else if (a.productdata[index].Products[0][property] > b.productdata[index].Products[0][property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });

  }
  // GetFandIPPData() {
  //   this.TotalData = [];
  //   this.spinner.show();
  //   const obj = {
  //     StartDate: this.FromDate.replace(/-/g, '/'),
  //     EndDate: this.ToDate.replace(/-/g, '/'),
  //     Stores: this.StoreIds,
  //   };
  //   //console.log(obj);
  //   this.apiSrvc
  //     .postmethod(this.comm.routeEndpoint+'GetFandIProductPenetrationByCategories', obj)
  //     .subscribe(
  //       (res) => {
  //         if (res.status == 200) {
  //           this.FandIPPData = res.response;
  //           // const index = this.FandIPPData.findIndex(
  //           //   (i:any) => i.StoreName == 'Total'
  //           // );
  //           // if (index >= 0) {
  //           //   this.TotalData.push(this.FandIPPData[index]);
  //           //   this.FandIPPData.splice(index, 1);
  //           // this.FandIPPData.push(this.TotalData[0]);

  //           // }
  //           if (this.FandIPPData != undefined) {
  //             if (this.FandIPPData.length > 0) {
  //               this.FandIPPData.some(function (x: any) {
  //                 if (x.productdata != undefined) {
  //                   x.productdata = JSON.parse(x.productdata);
  //                   var heading = x.productdata.map(function (a: any) {
  //                     return a.ProductCategory;
  //                   });
  //                 }
  //               });

  //               if (this.FandIPPData[0].productdata != undefined) {
  //                 this.Headings = this.FandIPPData[0].productdata.map(function (
  //                   a: any
  //                 ) {
  //                   return a.ProductCategory;
  //                 });
  //               }
  //             }
  //           }

  //           //console.log('Report Total', this.FandIPPData);
  //           //console.log('F & I Data', this.FandIPPData, this.Headings);
  //           this.spinner.hide();
  //           this.NodataFound = true;
  //           if (this.FandIPPData.length > 0) {
  //             this.NoData = false;
  //           } else {
  //             this.NoData = true;
  //           }
  //         } else {
  //          
  //         }
  //       },
  //       (error) => {
  //         //console.log(error);
  //       }
  //     );
  // }
  fandippdetails: any;
  opendetails(layer1: any, layer2: any, layer3: any, layer4: any, block: any, index: any) {
    //console.log(layer1, layer2, layer3, layer4);
    this.fandippdetails = {
      LAYER1: layer1,
      LAYER2: layer2,
      LAYER3: layer3,
      LAYER4: this.path1 == 'FIMANAGER' ? layer1 : layer4,
      STARTDATE: this.FromDate.replace(/-/g, '/'),
      ENDDATE: this.ToDate.replace(/-/g, '/'),
      LAYER5: (this.path2 == 'Store_Name' && block == 2) ? layer4.store_id : '',
      Grouping1: this.path1name == 'All Dealerships' ? 'Stores' : this.path1name,
      Grouping2: this.path2name
      //  FI_ID:this.path1=='FIMANAGER' : 
    };
    this.getCategoryDealLevel()
    // this.ngbmodalActive = this.ngbmodal.open(temp, {
    //   size: 'sm',
    //   backdrop: 'static',
    // });
  }

  // Detailsclose(event: any) {
  //   this.ngbmodalActive.dismiss();
  // }
  FIPPstate: any;
  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'F & I Product Penetration') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: { obj: { state: boolean; title: string; }; }) => {
      if (this.excel != undefined) {
        this.FIPPstate = res.obj.state;
        if (res.obj.title == 'F & I Product Penetration') {
          if (res.obj.state == true) {
            this.exportToExcel();
          }
        }
      }
    });


    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: { obj: { title: string; statePrint: boolean; }; }) => {
      if (this.print != undefined) {
        if (res.obj.title == 'F & I Product Penetration') {
          if (res.obj.statePrint == true) {
            // this.GetPrintData();
          }
        }
      }
    });

    this.Pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: { obj: { title: string; statePDF: boolean; }; }) => {
      if (this.Pdf != undefined) {
        if (res.obj.title == 'F & I Product Penetration') {
          if (res.obj.statePDF == true) {
            // this.generatePDF();
          }
        }
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: { obj: { title: string; stateEmailPdf: boolean; Email: any; notes: any; from: any; }; }) => {
      if (this.email != undefined) {
        if (res.obj.title == 'F & I Product Penetration') {
          if (res.obj.stateEmailPdf == true) {
            // this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
          }
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
    this.getEmployees();
  }
  ngOnDestroy() {
    // this.reportOpenSub.unsubscribe()
    // this.reportGetting.unsubscribe()
    // this.Pdf.unsubscribe()
    // this.print.unsubscribe()
    // this.email.unsubscribe()
    // this.excel.unsubscribe()

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


  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    } else if (value < 0) {
      return false;
    }
    return true;
  }
  expandorcollapse(ind: any, e: any, ref: any, Item: any, parentData: any) {
    let id = (e.target as Element).id;
    if (id == 'D_' + ind) {
      if (this.path2 == '') {
        // //console.log(this.path2);
        // this.openDetails(Item, parentData, '1');
      } else {
        if (ref == '-') {
          Item.Dealer = '+';
        }
        if (ref == '+') {
          Item.Dealer = '-';
        }
      }
    }
    if (id == 'DN_' + ind) {
      // //console.log(Item)
      // if (this.path3 == '') {
      //   this.openDetails(Item, parentData, '2');
      // } else {
      if (ref == '-') {
        Item.data2sign = '+';
        Item.SubData = Item.SubData.map((obj: any, i: any) => ({ ...obj, data3sign: '+' }))
        // //console.log(Item);

      }
      if (ref == '+') {
        Item.data2sign = '-';
        Item.Dealer = '-';
      }
    }
    if (id == 'DNS_' + ind) {
      // //console.log(Item, ref, ind)
      // if (this.path3 == '') {
      //   this.openDetails(Item, parentData, '2');
      // } else {
      if (ref == '-') {
        Item.data3sign = '+';
      }
      if (ref == '+') {
        Item.data3sign = '-';
        Item.Dealer = '-';
        Item.data2sign = '-'
      }
    }
    // }
  }

  setDates(type: any) {
    this.DateType == 'C' ? this.displaytime = this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy') + ' - ' + this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy') :
      this.displaytime = '(' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ')';
    // this.maxDate = new Date();
    // this.minDate = new Date();
    // this.minDate.setFullYear(this.maxDate.getFullYear() - 3);
    // this.maxDate.setDate(this.maxDate.getDate());
    this.Dates.FromDate = this.FromDate;
    this.Dates.ToDate = this.ToDate;
    this.Dates.MinDate = this.minDate;
    this.Dates.MaxDate = this.maxDate;
    this.Dates.DateType = this.DateType;
    this.Dates.DisplayTime = this.displaytime;
    console.log(this.FromDate, this.ToDate, this.DateType, this.displaytime, '..............');
  }
  //------------Reports---------//
  selectedDataGrouping: any = [];
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  displaytime: any = '';
  Performance: any = 'Load';
  dataGrouping: any = [
    { "ARG_ID": 40, "ARG_LABEL": "Store", "ARG_SEQ": 0, "Active": 'Y', "columnname": 'Store_Name' },
    { "ARG_ID": 26, "ARG_LABEL": "F&I Manager", "ARG_SEQ": 1, "Active": 'Y', "columnname": 'FIMANAGER' },
  ];
  GroupingDetails: any = [];
  datagrp: any = [];

  financeManager: any = [];
  fiManagersname: any = '';
  salesPersons: any = [];
  salesManagers: any = [];
  selectedstorevalues: any = [];
  // stores: any = [];
  AllStores: boolean = true;
  overallSelectedpeople: any = 0
  Groupingcols = [
    { id: 40, columnname: 'Store_Name' },
    { id: 26, columnname: 'FIMANAGER' },
    // { id: 27, columnname: 'DealType' },

  ];
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
  activePopover: number = -1;
  togglePopover(popoverIndex: number) {
    if (this.activePopover === popoverIndex) {
      this.activePopover = -1;
    } else {
      this.activePopover = popoverIndex;
    }
  }
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card , .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }
  updatedDates(data: any) {
    console.log(data);
    this.FromDate = data.FromDate;
    this.ToDate = data.ToDate;
    this.DateType = data.DateType;
    this.displaytime = data.DisplayTime
  }

  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
  }
  SelectedData(val: any) {
    // //console.log(val);

    const index = this.selectedDataGrouping.findIndex((i: any) => i == val);
    if (index >= 0) {
      this.selectedDataGrouping.splice(index, 1);
    } else {
      if (this.selectedDataGrouping.length >= 2) {

        this.toast.show('Select up to 2 Filters only to Group your data.', 'danger', 'Error');

      } else {
        this.selectedDataGrouping.push(val)
      }
    }

  }
  getDataGroupings() {
    this.selectedDataGrouping.push(this.dataGrouping[0]);
    // this.selectedDataGrouping.push(this.dataGrouping[1]);
  }

  spinnerLoaderPeople: boolean = false;
  getEmployees(val?: any, ids?: any, count?: any, bar?: any) {
    const obj = {
      AS_ID: this.storeIds.toString(),
      type: 'F',
    };
    this.spinnerLoaderPeople = true;
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetEmployeesDev', obj).subscribe(
      (res: any) => {
        if (res && res.status == 200) {
          this.spinnerLoaderPeople = false;
          // if (val == 'F') {
          this.financeManager = res.response.filter((e: any) => e.FiName != 'Unknown');
          this.financeManagerId = this.financeManager.map(function (a: any) { return a.FiId; });
        } else {
          this.spinnerLoaderPeople = false;
          this.toast.show('Invalid Details.', 'danger', 'Error');
        }
      },
      (error: any) => { /* ignore console errors */ }
    );
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
  isApplyClicked = false;
  viewreport() {
    this.activePopover = -1;
    this.isApplyClicked = true;
    if (this.storeIds.length == 0) {
      this.toast.show('Please select atleast any one Store', 'warning', 'Warning');
    }
    else if (this.dataGrouping && this.dataGrouping.length == 0) {

      this.toast.show('Please select atleast any one grouping', 'warning', 'Warning');
    }
    else {
      this.setHeaderData()
      this.getFandIPPData()

    }


  }

  //--------Details-------------//

  NoDatacategories: boolean = false
  category: boolean = false
  spinnerLoader: boolean = false
  fandippcategories: any = []
  getCategoryDealLevel() {
    this.category = false;
    this.spinnerLoader = true
    const obj = {
      "PT_ID": this.fandippdetails.LAYER2.id,
      "AP_ID": '',
      "StartDate": this.fandippdetails.STARTDATE,
      "EndDate": this.fandippdetails.ENDDATE,
      "Store": this.fandippdetails.LAYER1.store_id == undefined ? this.fandippdetails.LAYER5 : this.fandippdetails.LAYER1.store_id,
      "FI_ID": this.fandippdetails.LAYER4 != '' ? (this.fandippdetails.LAYER4.FIMANAGER_ID == '' ? 0 : this.fandippdetails.LAYER4.FIMANAGER_ID) : '',
      "Dealtype": this.fandippdetails.Grouping2 == 'DealType' && this.fandippdetails.LAYER4 != '' ? this.fandippdetails.LAYER4.DealType : '',
      "ProductType": this.selectedProductType && this.selectedProductType.length > 0 ? this.selectedProductType.map(function (a: any) { return a.TID; }).toString() : 0,
      "ProductCategory": this.selectedProductCategory && this.selectedProductCategory.length > 0 ? this.selectedProductCategory.map(function (a: any) { return a.CID; }).toString() : 0,
    }
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetFandIProductPenetrationbyCategoriesDealLevelNew', obj).subscribe(
      (res: any) => {
        if (res.status == 200) {
          this.fandippcategories = res.response
          this.spinnerLoader = false
          this.category = true
          if (this.fandippcategories.length > 0) {
            this.NoDatacategories = false;
          } else {
            this.NoDatacategories = true;
          }
        }
        else {
          this.toast.show('Invalid Details.', 'danger', 'Error');

        }
      },
      (error: any) => {
        //console.log(error);
      })
  }



  // exportToExcel(): void {

  //   const workbook = this.shared.getWorkbook();
  //   const worksheet = workbook.addWorksheet('F & I Product Penetration');

  //   /* =======================
  //      DATE HEADER
  //   ======================= */
  //   const dateText =
  //     `Date : ${this.shared.datePipe.transform(this.FromDate,'MMMM dd')}` +
  //     ` - ${this.shared.datePipe.transform(this.ToDate,'MMMM dd, yyyy')}`;

  //   /* =======================
  //      HEADER ROW 1 (GROUPS)
  //   ======================= */
  //   const headerRow1: any[] = [dateText, 'Product Totals'];
  //   this.Headings.forEach((h: any) => headerRow1.push(h));

  //   worksheet.addRow(headerRow1);

  //   worksheet.mergeCells('A1:A2');
  //   worksheet.mergeCells('B1:I1');

  //   let colIndex = 10;
  //   this.Headings.forEach(() => {
  //     worksheet.mergeCells(1, colIndex, 1, colIndex + 4);
  //     colIndex += 5;
  //   });

  //   worksheet.getRow(1).font = { bold: true };
  //   worksheet.getRow(1).alignment = { horizontal: 'center', vertical: 'middle' };

  //   /* =======================
  //      HEADER ROW 2 (COLUMNS)
  //   ======================= */
  //   const headerRow2 = [
  //     '',
  //     'Gross','PVR','Count','Retail Units','Pen %','Per Trans','PPS',
  //     ...this.Headings.flatMap(() => [
  //       'Count','Pen %','Gross','PVR','PPS'
  //     ])
  //   ];

  //   worksheet.addRow(headerRow2);
  //   worksheet.getRow(2).font = { bold: true };
  //   worksheet.getRow(2).alignment = { horizontal: 'center' };

  //   /* =======================
  //      DATA ROWS
  //   ======================= */
  //   for (const FIPP of this.FandIPPData) {

  //     /* ===== MAIN ROW ===== */
  //     const mainRow: any[] = [
  //       FIPP.data1 || '-',
  //       FIPP.toatlBG || '-',
  //       FIPP.TotalPVR || '-',
  //       FIPP.TotalProduct || '-',
  //       FIPP.Retailcount || '-',
  //       FIPP.TotalPerSale ? FIPP.TotalPerSale + '%' : '-',
  //       FIPP.totproductpertrans || '-',
  //       FIPP.ProfitPerSale || '-',
  //     ];

  //     FIPP.productdata?.forEach((sub: { Products: any[]; }) => {
  //       sub.Products?.forEach(pd => {
  //         mainRow.push(
  //           pd.PrdocuctCatcount || '-',
  //           pd.ProductcatPersale ? pd.ProductcatPersale + '%' : '-',
  //           pd.ProductcatBG || '-',
  //           pd.ProductcatPVR || '-',
  //           pd.ProfitPerSale || '-'
  //         );
  //       });
  //     });

  //     const row = worksheet.addRow(mainRow);
  //     row.font = { bold: true };

  //     /* ===== SUB ROWS ===== */
  //     FIPP.data2?.forEach((dt2: { data1: any; toatlBG: any; TotalPVR: any; TotalProduct: any; Retailcount: any; TotalPerSale: string; totproductpertrans: any; ProfitPerSale: any; productdata: any[]; }) => {

  //       const subRow: any[] = [
  //         '   ' + (dt2.data1 || '-'),
  //         dt2.toatlBG || '-',
  //         dt2.TotalPVR || '-',
  //         dt2.TotalProduct || '-',
  //         dt2.Retailcount || '-',
  //         dt2.TotalPerSale ? dt2.TotalPerSale + '%' : '-',
  //         dt2.totproductpertrans || '-',
  //         dt2.ProfitPerSale || '-',
  //       ];

  //       dt2.productdata?.forEach((sub: { Products: any[]; }) => {
  //         sub.Products?.forEach((pd: { PrdocuctCatcount: any; ProductcatPersale: string; ProductcatBG: any; ProductcatPVR: any; ProfitPerSale: any; }) => {
  //           subRow.push(
  //             pd.PrdocuctCatcount || '-',
  //             pd.ProductcatPersale ? pd.ProductcatPersale + '%' : '-',
  //             pd.ProductcatBG || '-',
  //             pd.ProductcatPVR || '-',
  //             pd.ProfitPerSale || '-'
  //           );
  //         });
  //       });

  //       worksheet.addRow(subRow);
  //     });
  //   }

  //   /* =======================
  //      FORMATTING
  //   ======================= */
  //   worksheet.eachRow((row: { eachCell: (arg0: (cell: any, colNumber: any) => void) => void; }, rowNumber: any) => {
  //     row.eachCell((cell: { value: any; alignment: { horizontal: string; }; numFmt: string; }, colNumber: number) => {

  //       if (typeof cell.value === 'number') {
  //         cell.alignment = { horizontal: 'right' };

  //         // currency columns
  //         if (colNumber !== 1) {
  //           cell.numFmt = '"$"#,##0';
  //         }
  //       }
  //     });
  //   });

  //   worksheet.columns.forEach(col => {
  //     col.width = 18;
  //   });

  //   worksheet.views = [{ state: 'frozen', ySplit: 2 }];

  //   /* =======================
  //      DOWNLOAD
  //   ======================= */
  //   workbook.xlsx.writeBuffer().then((buffer: BlobPart) => {
  //     const blob = new Blob([buffer], {
  //       type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  //     });
  //     this.shared.exportToExcel(workbook, 'F & I Product Penetration');
  //   });
  // }
  // exportToExcel(): void {

  //   const workbook = this.shared.getWorkbook();
  //   const worksheet = workbook.addWorksheet('F & I Product Penetration');

  //   /* =======================
  //      DATE TEXT
  //   ======================= */
  //   const dateText =
  //     `Date : ${this.shared.datePipe.transform(this.FromDate,'MMMM dd')}` +
  //     ` - ${this.shared.datePipe.transform(this.ToDate,'MMMM dd, yyyy')}`;

  //   /* =======================
  //      HEADER ROW 1 (GROUPS)
  //   ======================= */
  //   worksheet.addRow([]);

  //   worksheet.getCell('A1').value = dateText;

  //   worksheet.getCell('B1').value = 'Product Totals';
  //   worksheet.mergeCells(1, 2, 1, 9); // 8 cols

  //   let startCol = 10; // J
  //   this.Headings.forEach((head: any) => {
  //     worksheet.getCell(1, startCol).value = head;
  //     worksheet.mergeCells(1, startCol, 1, startCol + 4); // colspan=5
  //     startCol += 5;
  //   });

  //   worksheet.getRow(1).font = { bold: true };
  //   worksheet.getRow(1).alignment = {
  //     horizontal: 'center',
  //     vertical: 'middle'
  //   };

  //   /* =======================
  //      HEADER ROW 2 (COLUMNS)
  //   ======================= */
  //   const headerRow2 = [
  //     '',
  //     'Gross','PVR','Count','Retail Units','Pen %','Per Trans','PPS',
  //     ...this.Headings.flatMap(() => [
  //       'Count','Pen %','Gross','PVR','PPS'
  //     ])
  //   ];

  //   worksheet.addRow(headerRow2);
  //   worksheet.getRow(2).font = { bold: true };
  //   worksheet.getRow(2).alignment = { horizontal: 'center' };

  //   worksheet.views = [{ state: 'frozen', ySplit: 2 }];

  //   /* =======================
  //      DATA ROWS
  //   ======================= */
  //   for (const FIPP of this.FandIPPData) {

  //     /* ===== MAIN ROW ===== */
  //     const mainRow: any[] = [
  //       FIPP.data1 || '-',
  //       FIPP.toatlBG || '-',
  //       FIPP.TotalPVR || '-',
  //       FIPP.TotalProduct || '-',
  //       FIPP.Retailcount || '-',
  //       FIPP.TotalPerSale ? FIPP.TotalPerSale + '%' : '-',
  //       FIPP.totproductpertrans || '-',
  //       FIPP.ProfitPerSale || '-',
  //     ];

  //     FIPP.productdata?.forEach((sub: { Products: any[]; }) => {
  //       sub.Products?.forEach((pd: { PrdocuctCatcount: any; ProductcatPersale: string; ProductcatBG: any; ProductcatPVR: any; ProfitPerSale: any; }) => {
  //         mainRow.push(
  //           pd.PrdocuctCatcount || '-',
  //           pd.ProductcatPersale ? pd.ProductcatPersale + '%' : '-',
  //           pd.ProductcatBG || '-',
  //           pd.ProductcatPVR || '-',
  //           pd.ProfitPerSale || '-'
  //         );
  //       });
  //     });

  //     const row = worksheet.addRow(mainRow);
  //     row.font = { bold: true };

  //     /* ===== SUB ROWS ===== */
  //     FIPP.data2?.forEach((dt2: { data1: any; toatlBG: any; TotalPVR: any; TotalProduct: any; Retailcount: any; TotalPerSale: string; totproductpertrans: any; ProfitPerSale: any; productdata: any[]; }) => {

  //       const subRow: any[] = [
  //         '   ' + (dt2.data1 || '-'),
  //         dt2.toatlBG || '-',
  //         dt2.TotalPVR || '-',
  //         dt2.TotalProduct || '-',
  //         dt2.Retailcount || '-',
  //         dt2.TotalPerSale ? dt2.TotalPerSale + '%' : '-',
  //         dt2.totproductpertrans || '-',
  //         dt2.ProfitPerSale || '-',
  //       ];

  //       dt2.productdata?.forEach((sub: { Products: any[]; }) => {
  //         sub.Products?.forEach((pd: { PrdocuctCatcount: any; ProductcatPersale: string; ProductcatBG: any; ProductcatPVR: any; ProfitPerSale: any; }) => {
  //           subRow.push(
  //             pd.PrdocuctCatcount || '-',
  //             pd.ProductcatPersale ? pd.ProductcatPersale + '%' : '-',
  //             pd.ProductcatBG || '-',
  //             pd.ProductcatPVR || '-',
  //             pd.ProfitPerSale || '-'
  //           );
  //         });
  //       });

  //       worksheet.addRow(subRow);
  //     });
  //   }

  //   /* =======================
  //      FORMATTING
  //   ======================= */
  //   worksheet.eachRow((row: { eachCell: (arg0: (cell: any, colNumber: any) => void) => void; }) => {
  //     row.eachCell((cell: { value: any; alignment: { horizontal: string; }; numFmt: string; }, colNumber: number) => {

  //       if (typeof cell.value === 'number') {
  //         cell.alignment = { horizontal: 'right' };
  //         if (colNumber !== 1) {
  //           cell.numFmt = '"$"#,##0';
  //         }
  //       }
  //     });
  //   });

  //   worksheet.columns.forEach(col => {
  //         col.width = 18;
  //       });
  //   /* =======================
  //      DOWNLOAD
  //   ======================= */
  //   workbook.xlsx.writeBuffer().then((buffer: BlobPart) => {
  //     const blob = new Blob([buffer], {
  //       type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  //     });
  //     this.shared.exportToExcel(workbook, 'F & I Product Penetration');
  //   });
  // }
  exportToExcel(): void {

    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('F & I Product Penetration');

    let storeValue = 'All Stores';
    if (this.storeIds && this.storeIds.length > 0) {
      storeValue = this.stores.filter((s: any) => this.storeIds.includes(s.ID)).map((s: any) => s.storename).join(', ');
    }

    const capitalize = (str: string) =>
      str ? str.toString().replace(/\b\w/g, char => char.toUpperCase()) : '';

    const groupingText =
      this.selectedDataGrouping.length === 1
        ? this.selectedDataGrouping[0].ARG_LABEL
        : this.selectedDataGrouping.length > 1
          ? this.selectedDataGrouping.map((g: { ARG_LABEL: any; }) => g.ARG_LABEL).join(', ')
          : 'All';

    const timeFrameText =
      `${this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy')} - ` +
      `${this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy')}`;


    const peopleText = this.financeManagerId?.length
      ? this.financeManager
        .filter((fm: { FiId: any; }) => this.financeManagerId.includes(fm.FiId))
        .map((fm: { FiName: any; }) => fm.FiName)
        .join(', ')
      : 'All F&I Managers';


    worksheet.addRow(['F & I PRODUCT PENETRATION']);
    worksheet.mergeCells('A1:C1');
    worksheet.getRow(1).font = { bold: true, size: 14 };
    worksheet.getRow(1).alignment = { horizontal: 'center', vertical: 'middle' };


    const filters = [
      { label: 'Store', value: storeValue },
      { label: 'Groupings', value: groupingText },
      { label: 'Time Frame', value: timeFrameText },
      { label: 'People', value: capitalize(peopleText) },
      { label: 'Product Type', value: this.selectedProductType && this.selectedProductType.length > 0 ? this.selectedProductType.map(function (a: any) { return a.TNAME; }).toString() : 'All', },
      { label: 'Product Category', value: this.selectedProductCategory && this.selectedProductCategory.length > 0 ? this.selectedProductCategory.map(function (a: any) { return a.CNAME; }).toString() : 'All', }


    ];

    let rowIndex = 2;
    filters.forEach(f => {
      worksheet.addRow([`${f.label} :`, f.value]);
      worksheet.mergeCells(`B${rowIndex}:Z${rowIndex}`);
      worksheet.getCell(`A${rowIndex}`).font = { bold: true };
      rowIndex++;
    });

    rowIndex++;
    worksheet.addRow([]);


    const dateText =
      `Date : ${this.shared.datePipe.transform(this.FromDate, 'MMMM dd')} - ` +
      `${this.shared.datePipe.transform(this.ToDate, 'dd, yyyy')}`;

    worksheet.addRow([]);
    const headerStartRow = rowIndex + 1;

    worksheet.getCell(`A${headerStartRow}`).value = dateText;

    worksheet.getCell(`B${headerStartRow}`).value = 'Product Totals';
    worksheet.mergeCells(headerStartRow, 2, headerStartRow, 9);

    let colIndex = 13;
    this.Headings.forEach((head: any) => {
      worksheet.getCell(headerStartRow, colIndex).value = head;
      worksheet.mergeCells(headerStartRow, colIndex, headerStartRow, colIndex + 4);
      colIndex += 5;
    });

    worksheet.getRow(headerStartRow).font = { bold: true };
    worksheet.getRow(headerStartRow).alignment = {
      horizontal: 'center',
      vertical: 'middle'
    };


    const headerRow2 = [
      '',
      'Gross', 'PVR', 'Products', 'Units', 'Per Trans', 'PPS', 'Finance', 'Income', 'Pen %', 'PFD', 'PPU',
      ...this.Headings.flatMap(() => [
        'Products', 'Pen %', 'Gross', 'PVR', 'PPS'
      ])
    ];

    worksheet.addRow(headerRow2);
    worksheet.getRow(headerStartRow + 1).font = { bold: true };
    worksheet.getRow(headerStartRow + 1).alignment = { horizontal: 'center' };

    worksheet.views = [{ state: 'frozen', ySplit: headerStartRow + 1 }];
    let subData: any = [13]
    let count = 13
    for (const FIPP of this.FandIPPData) {

      const mainRow: any[] = [
        FIPP.data1 || '-',
        FIPP.toatlBG || '-',
        FIPP.TotalPVR || '-',
        FIPP.TotalProduct || '-',
        FIPP.Retailcount || '-',
        // FIPP.TotalPerSale ? this.cp.transform(FIPP.TotalPerSale, 'USD', '', '1.0-0') + '%' : '-',
        FIPP.totproductpertrans ? this.cp.transform(FIPP.totproductpertrans, 'USD', '', '1.2-2') : '-',
        FIPP.ProfitPerSale || '-',
        FIPP.fin_DealCount || '-',
        FIPP.fin_income || '-',
        FIPP.Finance_pen ? this.cp.transform(FIPP.Finance_pen, 'USD', '', '1.2-2') + '%' : '-',
        FIPP.perFD || '-',
        FIPP.perRU || '-',

      ];

      FIPP.productdata?.forEach((sub: { Products: any[]; }) => {
        sub.Products?.forEach((pd: { PrdocuctCatcount: any; ProductcatPersale: string; ProductcatBG: any; ProductcatPVR: any; ProfitPerSale: any; }) => {
          subData.push(count + 5)
          mainRow.push(
            pd.PrdocuctCatcount ? this.cp.transform(pd.PrdocuctCatcount, 'USD', '', '1.0-0') : '-',
            pd.ProductcatPersale ? this.cp.transform(pd.ProductcatPersale, 'USD', '', '1.0-0') + '%' : '-',
            pd.ProductcatBG || '-',
            pd.ProductcatPVR || '-',
            pd.ProfitPerSale || '-'
          );
        });
      });

      worksheet.addRow(mainRow).font = { bold: true };

      FIPP.data2?.forEach((dt2: {
        data1: any; toatlBG: any; TotalPVR: any; TotalProduct: any; Retailcount: any; totproductpertrans: any;
        ProfitPerSale: any; fin_DealCount: any; fin_income: any; Finance_pen: any; perFD: any; perRU: any; productdata: any[];
      }) => {

        const subRow: any[] = [
          '   ' + (dt2.data1 || '-'),
          dt2.toatlBG || '-',
          dt2.TotalPVR || '-',
          dt2.TotalProduct || '-',
          dt2.Retailcount || '-',
          // dt2.TotalPerSale ? this.cp.transform(dt2.TotalPerSale, 'USD', '', '1.0-0') + '%' : '-',
          dt2.totproductpertrans ? this.cp.transform(dt2.totproductpertrans, 'USD', '', '1.2-2') : '-',
          dt2.ProfitPerSale || '-',
          dt2.fin_DealCount || '-',
          dt2.fin_income || '-',
          dt2.Finance_pen ? this.cp.transform(dt2.Finance_pen, 'USD', '', '1.2-2') + '%' : '-',
          dt2.perFD || '-',
          dt2.perRU || '-',
        ];

        dt2.productdata?.forEach((sub: { Products: any[]; }) => {
          sub.Products?.forEach((pd: { PrdocuctCatcount: any; ProductcatPersale: string; ProductcatBG: any; ProductcatPVR: any; ProfitPerSale: any; }) => {
            subRow.push(
              pd.PrdocuctCatcount ? this.cp.transform(pd.PrdocuctCatcount, 'USD', '', '1.0-0') : '-',
              pd.ProductcatPersale ? this.cp.transform(pd.ProductcatPersale, 'USD', '', '1.0-0') + '%' : '-',
              pd.ProductcatBG || '-',
              pd.ProductcatPVR || '-',
              pd.ProfitPerSale || '-'
            );
          });
        });

        worksheet.addRow(subRow);
      });
    }


    worksheet.eachRow((row: { eachCell: (arg0: (cell: any, colNumber: any) => void) => void; }) => {
      row.eachCell((cell: { value: any; alignment: { horizontal: string; }; numFmt: string; }, colNumber: number) => {
        if (typeof cell.value === 'number') {
          cell.alignment = { horizontal: 'right' };
          if (colNumber !== 1 && colNumber !== 4 && colNumber !== 5 && colNumber !== 6 && colNumber !== 8 && colNumber != 10 && !subData.includes(colNumber)) {
            cell.numFmt = '"$"#,##0';
          }
        }
      });
    });

    worksheet.columns.forEach(col => {
      col.width = 18;
    });

    workbook.xlsx.writeBuffer().then((buffer: any) => {
      this.shared.exportToExcel(workbook, 'F & I Product Penetration');
    });
  }

  exportAsXLSX() {
    let localarray = this.fandippcategories.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );
    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('F&I Product Penetration Details');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 12, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A13', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('')

    const DateToday = this.shared.datePipe.transform(new Date(), 'MM.dd.yyyy h:mm:ss a');

    const titleRow = worksheet.getCell("A2"); titleRow.value = 'F&I Product Penetration Details';
    titleRow.font = { name: 'Arial', family: 4, size: 15, bold: true };
    titleRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }



    // const DateBlock = worksheet.getCell("L2"); DateBlock.value = DateToday;
    // DateBlock.font = { name: 'Arial', family: 4, size: 10 };
    // DateBlock.alignment = { vertical: 'middle', horizontal: 'center' }
    // worksheet.addRow([''])

    const Store_Name = worksheet.addRow([this.fandippdetails.Grouping1 + ' :']);
    Store_Name.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true, };
    Store_Name.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }
    const StoreName = worksheet.getCell("B4"); StoreName.value = this.fandippdetails.LAYER1.data1;
    StoreName.font = { name: 'Arial', family: 4, size: 9 };
    StoreName.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }

    const grouping2 = worksheet.addRow(['Category :']);
    grouping2.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true, };
    grouping2.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }
    const grouping_2 = worksheet.getCell("B5"); grouping_2.value = this.fandippdetails.LAYER2.ProductCategory;
    grouping_2.font = { name: 'Arial', family: 4, size: 9 };
    grouping_2.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }

    if (this.fandippdetails.LAYER4 != "") {
      const grouping3 = worksheet.addRow([this.fandippdetails.Grouping2 + ' :']);
      grouping3.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true, };
      grouping3.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }
      const grouping_3 = worksheet.getCell("B6"); grouping_3.value = this.fandippdetails.LAYER4.data1;
      grouping_3.font = { name: 'Arial', family: 4, size: 9 };
      grouping_3.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }
    }
    const DATE_EXTENSION = this.shared.datePipe.transform(
      new Date(),
      'MMddyyyy'
    );
    // worksheet.addRow([DateToday]).font = {name: 'Arial',family: 4,size: 9};
    // const ReportFilter = worksheet.addRow(['Report Filters :']);
    // ReportFilter.font = {name: 'Arial',family: 4,size: 10,bold: true,};

    const StartDealDate = worksheet.addRow(['Start Date :']);
    StartDealDate.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true, };
    StartDealDate.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }
    // StartDealDate.border = {top: { style: 'thin' },left: { style: 'thin' },bottom: { style: 'thin' },right: { style: 'thin' }}
    const startdealdate = worksheet.getCell("B7");
    startdealdate.value = this.fandippdetails.STARTDATE;
    startdealdate.font = { name: 'Arial', family: 4, size: 9 };
    startdealdate.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }
    // startdealdate.border = {top: { style: 'thin' },left: { style: 'thin' },bottom: { style: 'thin' },right: { style: 'thin' }}

    const EndDealDate = worksheet.addRow(['End Date :']);
    EndDealDate.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true, };
    EndDealDate.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }
    // EndDealDate.border = {top: { style: 'thin' },left: { style: 'thin' },bottom: { style: 'thin' },right: { style: 'thin' }}
    const enddealdate = worksheet.getCell("B8");
    enddealdate.value = this.fandippdetails.ENDDATE;
    enddealdate.font = { name: 'Arial', family: 4, size: 9 };
    enddealdate.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' }


    worksheet.addRow('')
    let Headings = [
      'Sno', 'Stock #', 'Deal #', 'Product Name', 'Customer name', 'F&I Manager', 'Date', 'Sale', 'Cost', 'Total Gross'

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
      d.Date = this.shared.datePipe.transform(d.Date, 'MM.dd.yyyy');

      var obj = [
        count,
        (d.Stock == '' ? '-' : (d.Stock == null ? '-' : (d.Stock))),
        (d.dealno == '' ? '-' : (d.dealno == null ? '-' : (d.dealno))),
        (d.AP_NAME == '' ? '-' : (d.AP_NAME == null ? '-' : (d.AP_NAME))),
        (d.Customer == '' ? '-' : (d.Customer == null ? '-' : (d.Customer))),
        (d.FIManager == '' ? '-' : (d.FIManager == null ? '-' : (d.FIManager))),
        (d.Date == '' ? '-' : (d.Date == null ? '-' : (d.Date))),
        (d.frontgross == '' ? '-' : (d.frontgross == null ? '-' : (parseInt(d.frontgross)))),
        (d.backgross == '' ? '-' : (d.backgross == null ? '-' : (parseInt(d.backgross)))),
        (d.Cost == '' ? '-' : (d.Cost == null ? '-' : (d.Cost))),


      ];


      const Data1 = worksheet.addRow(obj);

      // //console.log(Data1);

      Data1.font = { name: 'Arial', family: 4, size: 8 }
      Data1.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 }
      // Data1.getCell(1).alignment = {indent: 1,vertical: 'top', horizontal: 'left'}
      Data1.eachCell((cell, number) => {
        cell.border = { right: { style: 'thin' } }
        if (number == 6) {
          //  //console.log(obj[number] , number , '................');      
          if (obj[number] < 0) {
            //console.log(obj[number] , number , '................');
            Data1.getCell(number + 1).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } };
          }
          Data1.getCell(number + 1).numFmt = '#,##0'
          Data1.getCell(number + 1).alignment = { indent: 1, vertical: 'middle', horizontal: 'right' }
        }
        if (number > 6 && number < 11 && obj[number] != undefined) {
          //  //console.log(obj[number] , number , '................');
          // cell.alignment = {indent:1,vertical:'middle',horizontal:'right'}           
          if (obj[number] < 0) {
            //console.log(obj[number] , number , '................');
            Data1.getCell(number + 1).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } };
          }
          Data1.getCell(number + 1).numFmt = '$#,##0'
        }


        //  else if(number > 6 && number < 12){
        //   cell.alignment = {indent:1,vertical:'middle',horizontal:'right'}
        //  }else if(number > 15 && number < 25){
        //   cell.alignment = {indent:1,vertical:'middle',horizontal:'right'}
        //  }else if(number == 6){
        //   cell.alignment = {indent:1,vertical:'middle',horizontal:'center'}
        //  }
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

    worksheet.addRow([]);
    workbook.xlsx.writeBuffer().then((buffer: any) => {
      this.shared.exportToExcel(workbook, 'F & I Product Penetration Details');
    });
  }

}
