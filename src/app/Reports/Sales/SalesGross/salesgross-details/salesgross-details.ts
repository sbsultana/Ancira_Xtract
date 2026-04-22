import { Component, ElementRef, Input, OnInit, Renderer2, ViewChild, } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { FilterPipe, FilterPipeModule } from 'ngx-filter-pipe';


@Component({
  selector: 'app-salesgross-details',
  standalone: true,
  imports: [SharedModule, FilterPipeModule],
  providers: [FilterPipe],
  templateUrl: './salesgross-details.html',
  styleUrl: './salesgross-details.scss'
})
export class SalesgrossDetails {
  @Input() Salesdetails: any = [];
  SalesPersonDetails: any = [];
  NoData!: boolean;
  spinnerLoader: boolean = true;
  SGDSearchName: any = '';
  constructor(
    public shared: Sharedservice,
    private renderer: Renderer2,

  ) {
    this.renderer.listen('window', 'click', (e: Event) => {
      const TagName = e.target as HTMLButtonElement;
      if (TagName.className === 'd-block modal fade show modal-static') {
        // this.closeBtn.nativeElement.click();
        this.close();
      }
    });
  }
  pageNumber: number = 0;
  pageSize: number = 100;
  hasMoreData: boolean = true;
  spinnerLoadersec: boolean = false;

  details: any = [];
  ngOnInit() {
    //console.log(this.Salesdetails)
    this.GetDetails();
  }
  GetDetails() {
    // this.spinner.show()
    const obj = {
      startdealdate: this.Salesdetails[0].StartDate,
      enddealdate: this.Salesdetails[0].EndDate,
      dealtype: this.Salesdetails[0].dealtype,
      saletype: this.Salesdetails[0].saletype,
      dealstatus: this.Salesdetails[0].dealstatus,
      var1: this.Salesdetails[0].var1,
      var2: this.Salesdetails[0].var2,
      var3: this.Salesdetails[0].var3,
      var1Value: this.Salesdetails[0].var1Value,
      var2Value: this.Salesdetails[0].var2Value,
      var3Value: this.Salesdetails[0].var3Value,
      PageNumber: this.pageNumber,
      PageSize: '500',
      SalesPerson: this.Salesdetails[0].SalesPerson,
      SalesManager: this.Salesdetails[0].SalesManager,
      FinanceManager: this.Salesdetails[0].FinanceManager
    };
    this.shared.api.postmethod(
      this.shared.common.routeEndpoint + 'GetSalesGrossSummaryDetails',
      obj
    ).subscribe((res) => {

      if (res.status === 200) {

        this.details = res.response || [];

        let spDetails = [
          ...this.SalesPersonDetails,
          ...this.details
        ];
        spDetails.forEach((x: any) => {
          if (x.products && typeof x.products === 'string') {
            try {
              x.products = JSON.parse(x.products);
            } catch {
              x.products = [];
            }
          }
        });
        if (this.details.length < this.pageSize) {
          this.hasMoreData = false;
        }
        this.SalesPersonDetails = spDetails;
        this.spinnerLoader = false;
        this.spinnerLoadersec = false;
        this.NoData = this.SalesPersonDetails.length === 0;
      }

    });
  }

  viewDeal(dealData: any) {
    //   const modalRef = this.ngbmodel.open(DealRecapComponent, { size: 'md', windowClass: 'connectedmodal' });
    //   modalRef.componentInstance.data = {  dealno: dealData.dealid, storeid: dealData.dealerid,stock:dealData.Stock,vin:dealData.vin,source:dealData.source,custno: dealData?.customernumber }; // Pass data to the modal component    
    //   modalRef.result.then((result) => {
    //     console.log(result); // Handle modal close result
    //   }, (reason) => {
    //     console.log(`Dismissed: ${reason}`); // Handle dismiss reason
    //   });
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
  close() {
    this.shared.ngbmodal.dismissAll();
  }
  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }
  currentElement!: string;

  @ViewChild('scrollOne') scrollOne!: ElementRef;
  count = 0;
updateVerticalScroll(event: any): void {
  const element = event.target;
  const atBottom =
    Math.ceil(element.scrollTop + element.clientHeight) >= element.scrollHeight;
  if (!atBottom) return;
  if (this.spinnerLoader) return;
  if (!this.hasMoreData) return;  

  this.spinnerLoader = true;
  this.pageNumber++;
  this.GetDetails();
}

  trTag: any = '';
  secondtrtag: any = ''

  hoverclass(e: any, i: any) {
    if (this.trTag) {
      this.trTag.classList.remove('hover');
    }
    if (this.secondtrtag) {
      this.secondtrtag.classList.remove('hover');
    }
    const id = (e.target as Element).id;
    this.trTag = document.getElementById('SD_' + i);
    this.secondtrtag = document.getElementById('SP_' + i);
    if ((id === 'SD_' + i || id === 'SP_' + i)) {

      if (this.trTag) {
        this.trTag.classList.add('hover');
      }

      if (this.secondtrtag) {
        this.secondtrtag.classList.add('hover');
      }
    }
  }



  exportAsXLSX() {
    this.shared.spinner.show()
    const obj = {
      startdealdate: this.Salesdetails[0].StartDate,
      enddealdate: this.Salesdetails[0].EndDate,
      dealtype: this.Salesdetails[0].dealtype,
      saletype: this.Salesdetails[0].saletype,
      dealstatus: this.Salesdetails[0].dealstatus,
      var1: this.Salesdetails[0].var1,
      var2: this.Salesdetails[0].var2,
      var3: this.Salesdetails[0].var3,
      var1Value: this.Salesdetails[0].var1Value,
      var2Value: this.Salesdetails[0].var2Value,
      var3Value: this.Salesdetails[0].var3Value,
      PageNumber: 0,
      PageSize: '10000',
      SalesPerson: this.Salesdetails[0].SalesPerson,
      SalesManager: this.Salesdetails[0].SalesManager,
      FinanceManager: this.Salesdetails[0].FinanceManager
      // "store_id": this.Salesdetails[0].storeId,
      // "spid":  this.Salesdetails[0].SPID == undefined || this.Salesdetails[0].SPID == null || this.Salesdetails[0].SPID == '' ? 0 : this.Salesdetails[0].SPID,
      // "startdealdate": this.Salesdetails[0].StartDate,
      // "enddealdate": this.Salesdetails[0].EndDate,
      // "dealtype": this.Salesdetails[0].dealType,
      // "type": ""
    };
    this.shared.api
      .postmethod(this.shared.common.routeEndpoint + 'GetSalesGrossSummaryDetails', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.shared.spinner.hide();
          if (res.response != undefined) {
            if (res.response.length > 0) {
              let localarray = res.response;
              localarray.some(function (x: any) {
                if (x.products != undefined) {
                  x.products = JSON.parse(x.products);
                }
              });

              const workbook = this.shared.getWorkbook();
              const worksheet = workbook.addWorksheet('Sales Gross Details');
              const capitalize = (str: string) =>
                str ? str.toString().replace(/\b\w/g, char => char.toUpperCase()) : '';
              let filters: any = [
                { name: 'Store :', values: this.Salesdetails[0].var1Value },
                { name: 'Time Frame :', values: this.Salesdetails[0].StartDate + ' to ' + this.Salesdetails[0].EndDate },
                { name: 'Deal Type : ', values: this.Salesdetails[0].dealtype == '' ? '-' : this.Salesdetails[0].dealtype == null ? '-' : this.Salesdetails[0].dealtype.toString().replaceAll(',', ', ') },
                { name: 'Sale Type :', values: this.Salesdetails[0].saletype == '' ? '-' : this.Salesdetails[0].saletype == null ? '-' : this.Salesdetails[0].saletype.toString().replaceAll(',', ', ').replace('Rental', 'Rental/Loaner') },
                { name: 'Deal Status :', values: this.Salesdetails[0].dealstatus == '' ? '-' : this.Salesdetails[0].dealstatus == null ? '-' : this.Salesdetails[0].dealstatus.toString().replaceAll(',', ', ').replace('Capped', 'Booked').replace('Finalized', 'Finalized') },
              ]
              const ReportFilter = worksheet.addRow(['Sales Gross Details']);
              ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };
              let startIndex = 2
              filters.forEach((val: any) => {
                startIndex++
                worksheet.addRow('');
                worksheet.getCell(`A${startIndex}`);
                worksheet.mergeCells(`B${startIndex}:C${startIndex}`);
                worksheet.getCell(`A${startIndex}`).value = val.name;
                worksheet.getCell(`B${startIndex}`).value = val.values
              })

              let secondHeader = [
               'Sl.no',
                'Deal #',
                'Stock #',
                'Date',
                'Customer',
                'Deal Type',
                'VIN',
                'Trade',
                'Sale Price',
                'Front Gross',
                'Back Gross',
                'Total Gross',
                'Salesperson 1',
                'Salesperson 2',
                'F&I Mgr.',
                'Desk Mgr.',
                'Fin. Reserve ',
                'Product Count',
              ];
              var dynamicheads = this.SalesPersonDetails[0].products.map(function (a: any) {
                return a.PCName;
              })
              secondHeader = [...secondHeader, ...dynamicheads]
              localarray = localarray.map(
                (v: any) => ({
                  ...v,
                  ...dynamicheads
                })
              )
              worksheet.addRow(secondHeader);
              worksheet.getRow(8).height = 25;
              worksheet.getRow(8).font = { bold: true, color: { argb: 'FFFFFFFF' } };
              worksheet.getRow(8).alignment = { vertical: 'middle', horizontal: 'center' };
              worksheet.getRow(8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };

              // let bindingHeaders = [
              //   'dealid', 'contractdate', 'customername', 'ad_Year', 'ad_make',
              //   'ad_model', 'ad_dealtype', 'vin', 'TradeAcv', 'frontgross_decimal', 'backgross_decimal',
              //   'totalgross_decimal', 'salesperson1', 'salesperson2', 'fimanager', 'salesmanager',
              //   'FinReserve', 'productcount'
              // ];

              // productsary.forEach((e: any) => {
              //   bindingHeaders = [
              //     ...bindingHeaders, e
              //   ];
              // })

              const currencyFields = ['TradeAcv', 'frontgross', 'backgross', 'totalgross', 'FinReserve'];
              var count = 0
              for (const d of localarray) {
                count++
                d.contractdate = this.shared.datePipe.transform(d.contractdate, 'MM/dd/yyyy');
                var productsary: any = []
                d.products.forEach((p: any, i: any) => {
                  productsary.push((p.Amount == '' ? '-' : (p.Amount == null ? '-' : (parseFloat(p.Amount)))))
                  // d[count]=p.amount
                })
                var obj = [count,
             (d.dealid == '' ? '-' : (d.dealid == null ? '-' : (d.dealid))),
                   (d.Stock == '' ? '-' : (d.Stock == null ? '-' : (d.Stock))),
                  (d.contractdate == '' ? '-' : (d.contractdate == null ? '-' : (d.contractdate))),
                  (d.customername == '' ? '-' : (d.customername == null ? '-' : (d.customername))),
                  (d.ad_dealtype == '' ? '-' : (d.ad_dealtype == null ? '-' : (d.ad_dealtype))),
                  (d.vin == '' ? '-' : (d.vin == null ? '-' : (d.vin))),
                  (d.TradeAcv == '' ? '-' : (d.TradeAcv == null ? '-' : (d.TradeAcv))),
                  (d.SalePrice == '' ? '-' : (d.SalePrice == null ? '-' : (d.SalePrice))),
                  (d.frontgross == '' ? '-' : (d.frontgross == null ? '-' : (parseInt(d.frontgross)))),
                  (d.backgross == '' ? '-' : (d.backgross == null ? '-' : (parseInt(d.backgross)))),
                  (d.totalgross == '' ? '-' : (d.totalgross == null ? '-' : (parseInt(d.totalgross)))),
                  (d.salesperson1 == '' ? '-' : (d.salesperson1 == null ? '-' : (d.salesperson1))),
                  (d.salesperson2 == '' ? '-' : (d.salesperson2 == null ? '-' : (d.salesperson2))),
                  (d.fimanager == '' ? '-' : (d.fimanager == null ? '-' : (d.fimanager))),
                  (d.salesmanager == '' ? '-' : (d.salesmanager == null ? '-' : (d.salesmanager))),
                  (d.FinReserve == '' ? '-' : (d.FinReserve == null ? '-' : (d.FinReserve))),
                  (d.productcount == '' ? '-' : (d.productcount == null ? '-' : (d.productcount))),
                ];
                productsary.forEach((e: any) => {
                  //console.log(e,'.............');

                  obj = [
                    ...obj,
                    e == '' ? '-' : e == null ? '-' : e,
                  ];
                })

                // (productsary[0] == ''?'-':(productsary[0] == null ? '-' :(productsary[0]))),
                // (productsary[1] == ''?'-':(productsary[1] == null ? '-' :(productsary[1]))),
                // (productsary[2] == ''?'-':(productsary[2] == null ? '-' :(productsary[2]))),
                // (productsary[3] == ''?'-':(productsary[3] == null ? '-' :(productsary[3]))),
                // (productsary[4] == ''?'-':(productsary[4] == null ? '-' :(productsary[4]))),
                // (productsary[5] == ''?'-':(productsary[5] == null ? '-' :(productsary[5]))),
                // (productsary[6] == ''?'-':(productsary[6] == null ? '-' :(productsary[6])))
                const Data1 = worksheet.addRow(obj);

                // //console.log(Data1);

                Data1.font = { name: 'Arial', family: 4, size: 8 }
                Data1.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 }
                // Data1.getCell(1).alignment = {indent: 1,vertical: 'top', horizontal: 'left'}
                Data1.eachCell((cell, number) => {
                  cell.border = { right: { style: 'thin' } }
                  cell.numFmt = '$#,##0.00'
                  if (number == 18) {
                    cell.numFmt = '#,##0'
                  } else if (number == 1) {
                    cell.numFmt = '###0'
                  } else if (number > 7 && number < 13) {
                    cell.alignment = { indent: 2, vertical: 'middle', horizontal: 'right' }
                  } else if (number > 16 && number < 26) {
                    cell.alignment = { indent: 2, vertical: 'middle', horizontal: 'right' }
                  } else if (number == 6) {
                    cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'center' }
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
              worksheet.columns.forEach((column: any) => {
                let maxLength = 15;
                column.width = maxLength + 2;
              });
              workbook.xlsx.writeBuffer().then((buffer: any) => {
                const blob = new Blob([buffer], {
                  type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                });
                this.shared.exportToExcel(workbook, 'Sales Gross Details')
              });
            }
          }
        }
        else {
          this.shared.spinner.hide()
          // this.toast.error(res.error)
        }
      });
  }




  fontSize: number = .75;
  minFontSize: number = .75;
  maxFontSize: number = .9;

  changeFontSize(action: 'increase' | 'decrease') {
    if (action === 'increase' && this.fontSize < this.maxFontSize) {
      this.fontSize += .1;
    } else if (action === 'decrease' && this.fontSize > this.minFontSize) {
      this.fontSize -= .1;
    }
  }
  get zoomInPossible(): boolean {
    return this.fontSize < this.maxFontSize;
  }
  get zoomOutPossible(): boolean {
    return this.fontSize > this.minFontSize;
  }
  get isMaxZoom(): boolean {
    return this.fontSize === this.maxFontSize;
  }
  get isMinZoom(): boolean {
    return this.fontSize === this.minFontSize;
  }

  sortedColumn: string | null = null;
  isAscending: boolean = true;

  // sortColumn(columnName: string) {
  //   if (this.sortedColumn === columnName) {
  //     this.isAscending = !this.isAscending;
  //   } else {
  //     this.sortedColumn = columnName;
  //     this.isAscending = true;
  //   }

  //   this.SalesPersonDetails.sort((a: any, b: any) => {
  //     const valueA = a[columnName];
  //     let stringValueA = String(valueA);
  //     const valueB = b[columnName];
  //     let stringValueB = String(valueB);
  //     return this.isAscending ? stringValueA.localeCompare(stringValueB) : stringValueB.localeCompare(stringValueA);
  //   });
  // }

  isDesc: boolean = false;
  column: string = 'CategoryName';
  sortColumn(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;

    this.SalesPersonDetails.sort((a: any, b: any) => {
      const valA = a[property] ?? ''; // replace null/undefined with empty string
      const valB = b[property] ?? '';
      if (valA < valB) {
        return -1 * direction;
      } else if (valA > valB) {
        return 1 * direction;
      } else {
        return 0;
      }
    });

  }

  getSortIcon(columnName: string): string {
    if (this.sortedColumn === columnName) {
      return this.isAscending ? 'asc' : 'desc';
    }
    return '';
  }

  DealAcctDetails: any;
  DealDeatils(val: any) {
    this.DealAcctDetails = [];
    //console.log(val);
    this.DealAcctDetails.push(val)
  }
}
