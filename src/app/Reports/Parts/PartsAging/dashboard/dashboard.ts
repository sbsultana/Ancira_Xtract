import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { common } from '../../../../common';
import { Subscription } from 'rxjs';
import { environment } from '../../../../../environments/environment';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import { CurrencyPipe } from '@angular/common';
import { Workbook } from 'exceljs';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';

@Component({
  selector: 'app-dashboard',
  imports: [SharedModule, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {

  PartsData: any = [];
  IndividualPartsGross: any = [];
  TotalPartsGross: any = [];
  PartsSource: any = []
  selectedpartssource: any = [];
  TotalReport: any = 'B';
  NoData: boolean = false;
  DateType: any = '30';
  responcestatus: any = '';
  partsDetails: any = 'Y'

  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  storeIds: any = 0;
  otherStoresArray: any = [];
  otherStoreIds: any = [];

  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'Y',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname, otherStoresArray: this.otherStoresArray, otherStoreIds: this.otherStoreIds
  };

  // index: any;

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .reportstores-card, .timeframe');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  constructor(
    public shared: Sharedservice, public setdates: Setdates, private comm: common, private cp: CurrencyPipe, private toast: ToastService,
  ) {
    this.shared.setTitle(this.comm.titleName + '-Parts Aging');
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
      this.storeIds = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',')
      this.otherStoreIds = JSON.parse(localStorage.getItem('otherstoreids')!);

    }
    if (this.shared.common.groupsandstores.length > 0) {
      this.groupsArray = this.shared.common.groupsandstores.filter((val: any) => val.sg_id != this.shared.common.reconID);
      this.otherStoresArray = this.shared.common.OtherStoresData && this.shared.common.OtherStoresData.length > 0 ? this.shared.common.OtherStoresData[0].Stores : []

      this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
      this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_Name : this.groupName = ''
      this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
      this.getStoresandGroupsValues()
    }
    this.setHeaderData()
    this.getPartssource()
    this.GetData();

  }

  ngOnInit(): void {
  }

  setHeaderData() {
    const data = {
      title: 'Parts Aging',
      stores: this.storeIds,
      partssource: this.PartsSource,
      ToporBottom: this.TotalReport,
      groups: this.groupId,
      otherstoreids: this.otherStoreIds

    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }
  getPartssource(count?: any) {
    const obj = {
      StoreID: this.storeIds.toString()
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsSource', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          this.PartsSource = res.response.filter((x: any) => x.AP_Source != '')
          this.selectedpartssource = this.PartsSource.map(function (a: any) {
            return a.AP_Source;
          });
        } else {
          this.toast.show('Invalid Details', 'danger', 'Error');
        }
      },
      (error) => {
        // //console.log(error);
      }
    );
  }

  allsource() {
    if (this.selectedpartssource.length == this.PartsSource.length) {
      this.selectedpartssource = []
    } else {
      this.selectedpartssource = this.PartsSource.map(function (a: any) {
        return a.AP_Source;
      });
    }
  }
  getPartsData() {
    this.NoData = false
    this.responcestatus = '';
    this.IndividualPartsGross = [];
    this.TotalPartsGross = [];
    this.shared.spinner.show();
    if (this.storeIds != '' || this.otherStoreIds != '') {
      this.GetData();
    } else {
      this.NoData = true;
      this.shared.spinner.hide()
    }
    // this.GetTotalData();
  }

  GetData() {
    this.IndividualPartsGross = [];
    this.shared.spinner.show();
    const obj = {
      StoreID: [...this.storeIds, ...this.otherStoreIds],
      UserID: 0,
      Source: this.selectedpartssource.toString(),
      DateType: this.DateType
    };
    this.shared.api.postmethod(this.comm.routeEndpoint + 'GetPartsAgingReportV1', obj).subscribe(
      (res) => {
        if (res.status == 200) {
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.shared.spinner.hide();
              this.IndividualPartsGross = [];
              this.IndividualPartsGross = res.response;
              //console.log(this.IndividualPartsGross)
              this.responcestatus = this.responcestatus + 'I';
              let idi_len = this.IndividualPartsGross.length;
              this.IndividualPartsGross.some(function (x: any) {
                if (x.AgeGroupData != undefined && x.AgeGroupData != '') {
                  x.AgeGroupData = JSON.parse(x.AgeGroupData);
                  x.Dealer = '-';
                }
                // if (idi_len == 1) {
                //   x.Dealer = '-';
                // } else {
                //   x.Dealer = '+';
                // }
              });
              console.log(this.IndividualPartsGross, 'chkkk')

              // this.combineIndividualandTotal();
              // if (this.TotalReport == 'T') {
              //   let last = this.IndividualPartsGross.pop()
              //   this.IndividualPartsGross.unshift(last)
              // }

              const idx = this.IndividualPartsGross.findIndex(
                (x: any) => x?.StoreName === 'REPORT TOTAL'
              );

              if (idx > -1) {
                const [reportTotalRow] = this.IndividualPartsGross.splice(idx, 1); // remove it

                if (this.TotalReport === 'T') {
                  this.IndividualPartsGross.unshift(reportTotalRow); // bring to top
                } else {
                  this.IndividualPartsGross.push(reportTotalRow);    // bring to bottom
                }
              }
              //               else {
              //   const first = this.IndividualPartsGross.shift();
              //   if (first !== undefined) this.IndividualPartsGross.push(first);
              // }
            }
            else {
              // this.toast.error('Empty Response','');
              this.shared.spinner.hide();
              this.NoData = true;

            }
          } else {
            // this.toast.error('Empty Response');
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




  calculateQtyPer(qty: number, total: number): number {
    if (!qty || !total || total === 0) {
      return 0;
    }
    return (qty * 100) / total;
  }

  calculatestock(saleinfo: any) {
    if (!saleinfo || !saleinfo.AgeGroupData || saleinfo.Stock_Cost == 0) return 0;
    const ageGroup12Plus = saleinfo.AgeGroupData.find((a: any) => a.Never_sold === '12+');
    if (ageGroup12Plus && ageGroup12Plus.Stock_Cost != null) {
      return (ageGroup12Plus.Stock_Cost * 100) / saleinfo.Stock_Cost;
    }
    return 0;
  }
  calculatenonstock(saleinfo: any) {
    if (!saleinfo || !saleinfo.AgeGroupData || saleinfo.NonStock_Cost == 0) return 0;
    const ageGroup12Plus = saleinfo.AgeGroupData.find((a: any) => a.Never_sold === '12+');
    if (ageGroup12Plus && ageGroup12Plus.NonStock_Cost != null) {
      return (ageGroup12Plus.NonStock_Cost * 100) / saleinfo.NonStock_Cost;
    }
    return 0;
  }
  calculatetotal(saleinfo: any) {
    if (!saleinfo || !saleinfo.AgeGroupData || saleinfo.Total_Cost == 0) return 0;
    const ageGroup12Plus = saleinfo.AgeGroupData.find((a: any) => a.Never_sold === '12+');
    if (ageGroup12Plus && ageGroup12Plus.Total_Cost != null) {
      return (ageGroup12Plus.Total_Cost * 100) / saleinfo.Total_Cost;
    }
    return 0;
  }

  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    else if (value < 0) {
      return false;
    }
    return true
  }
  subdataindex: any = 0;
  expandorcollapse(ind: any, e: any, ref: any, Item: any, parentData: any) {
    let id = (e.target as Element).id;
    if (id == 'D_' + ind) {

      if (ref == '-') {
        Item.Dealer = '+';
      }
      if (ref == '+') {
        Item.Dealer = '-';
      }

    }
    if (id == 'DN_' + ind) {
      if (ref == '-') {
        Item.data2sign = '+';
      }
      if (ref == '+') {
        Item.data2sign = '-';
        Item.Dealer = '-';
      }
    }
  }
  selectedData: any = { store_id: '', daterange: '' }
  daterange: any;

  openDetails(Item: any, ParentItem: any, ref: any) {
    if (ParentItem.StoreName != 'REPORT TOTAL') {
      if (ref == '2') {
        this.selectedData.daterange = Item.Never_sold
        this.selectedData.store_id = ParentItem.Store
        this.daterange = Item.Never_sold?.replace(' months', '');
        this.GetDetails()
      }
    }

  }
  details: any = []
  PartsPersonDetails: any = []
  spinnerLoader: boolean = false;
  spinnerLoadersec: boolean = false
  GetDetails() {
    this.partsDetails = 'N'
    this.PartsPersonDetails = []
    this.details = []
    this.shared.spinner.show()
    const obj = {
      store_id: this.selectedData.store_id,
      daterange: this.selectedData.daterange,
      // PageNumber: this.pageNumber,
      // PageSize: '100',

    };
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetFbPartsAgingJSONDetails', obj)
      .subscribe((res) => {
        if (res.status == 200) {
          this.details = res.response;
          this.PartsPersonDetails = [
            ...this.PartsPersonDetails,
            ...this.details,
          ];
          console.log(this.PartsPersonDetails);
          this.shared.spinner.hide()
          this.spinnerLoader = false;
          this.spinnerLoadersec = false;
          // this.PartsPersonDetails=res.response
          console.log(this.PartsPersonDetails);

          this.PartsPersonDetails.some(function (x: any) {
            if (x.PartsJsonDetails != undefined && x.PartsJsonDetails != '') {
              x.PartsJsonDetails = JSON.parse(x.PartsJsonDetails);
              x.Dealer = '-';
            }
          });


          // this.shared.spinner.hide()
          this.spinnerLoader = false;
          if (this.PartsPersonDetails.length > 0) {
            this.NoData = false;
          } else {
            this.NoData = true;
          }
        }
      });
  }
  backtoWR() {
    this.partsDetails = 'Y'
  }
  isDesc: boolean = false;
  column: string = 'CategoryName';

  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // //console.log(property)
    this.PartsData.sort(function (a: any, b: any) {
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
  ngAfterViewInit(): void {

    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == 'Parts Aging') {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.otherStoresArray = this.shared.common.OtherStoresData && this.shared.common.OtherStoresData.length > 0 ? this.shared.common.OtherStoresData[0].Stores : []

          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.storeIds.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.storeIds.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.storeIds)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Inventory Browser') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Inventory Browser') return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Inventory Browser') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Inventory Browser') return;
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
  Selectedpartssource(e: any) {
    const index = this.selectedpartssource.findIndex((i: any) => i == e.AP_Source);
    if (index >= 0) {
      this.selectedpartssource.splice(index, 1);
    } else {
      this.selectedpartssource.push(e.AP_Source);
    }
  }

  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
    this.otherStoreIds = data.otherStoreIds;

    this.getPartssource()
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
      'type': 'M', 'others': 'Y', otherStoresArray: this.otherStoresArray,
      otherStoreIds: this.otherStoreIds
    };
  }
  multipleorsingle(block: any, e: any) {
    if (block == 'TB') {
      this.TotalReport = e;
    }
  }
  activePopover: number = -1;

  togglePopover(popoverIndex: number) {
    this.activePopover = this.activePopover === popoverIndex ? -1 : popoverIndex;
  }
  viewreport() {
    this.activePopover = -1

    if (this.storeIds.length == 0 && this.otherStoreIds.length == 0) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    } else {
      this.setHeaderData();
      this.GetData()
    }

  }
  generatePDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save('Parts Aging.pdf'); // ✅ only here
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
    this.shared.spinner.show();
    try {
      const pdfBlob = this.generatePDFBlob();

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
      const pdfFile = this.blobToFile(pdfBlob, 'Inventory Browser.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Inventory Browser');
      formData.append('file', pdfFile);
      formData.append('notes', notes);
      formData.append('from', from);

      this.shared.api.postmethod(this.comm.routeEndpoint + 'mail', formData)
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
    if (val == null || val === '' || Number(val) === 0) return '-';
    return Number(val).toFixed(2) + '%'; // or toFixed(2) if needed
  }

  private createPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });

    doc.setFontSize(14);
    doc.text('Parts Aging', 14, 12);

    const startY = 18;

    /* ================= HEADER ================= */
    const head = [
      [
        { content: 'AS OF', colSpan: 1, styles: { halign: 'center' } },
        { content: 'Stock', colSpan: 4, styles: { halign: 'center' } },
        { content: 'Non-Stock', colSpan: 4, styles: { halign: 'center' } },
        { content: 'Total', colSpan: 4, styles: { halign: 'center' } }
      ],
      [
        '',
        'Qty', 'Qty%', 'Cost', 'Idle',
        'Qty', 'Qty%', 'Cost', 'Idle',
        'Qty', 'Qty%', 'Cost', 'Idle'
      ]
    ];

    const body: any[] = [];

    /* ================= HELPER ================= */
    const isRowEmpty = (row: any[]): boolean => {
      return row.every(val =>
        val === '' ||
        val === '-' ||
        val === null ||
        val === undefined ||
        val === '0' ||
        val === '$ 0'
      );
    };

    /* ================= BODY ================= */
    this.IndividualPartsGross.forEach((d: any) => {

      /* ===== LEVEL 1 (STORE ROW) ===== */
      const level1Row = [
        d.StoreName || '-',

        this.formatNumber(d.Stock_Qty),
        '',
        this.formatCurrency(d.Stock_Cost),
        this.formatCurrency(d.Stock_Idle),

        this.formatNumber(d.NonStock_Qty),
        '',
        this.formatCurrency(d.NonStock_Cost),
        this.formatCurrency(d.NonStock_Idle),

        this.formatNumber(d.Total_Qty),
        '',
        this.formatCurrency(d.Total_Cost),
        this.formatCurrency(d.Total_Idle)
      ];

      if (!isRowEmpty(level1Row)) {
        body.push({
          row: level1Row,
          indent: 0,
          isTotalRow: d.data1 === 'Reports Total'
        });
      }

      /* ===== LEVEL 2 ===== */
      d.AgeGroupData?.forEach((d1: any) => {

        const level2Row = [
          d1.Never_sold || '-',

          this.formatNumber(d1.Stock_Qty),
          this.formatPercent(this.calculateQtyPer(d1.Stock_Qty, d.Stock_Qty)),
          this.formatCurrency(d1.Stock_Cost),
          this.formatCurrency(d1.Stock_Idle),

          this.formatNumber(d1.NonStock_Qty),
          this.formatPercent(this.calculateQtyPer(d1.NonStock_Qty, d.NonStock_Qty)),
          this.formatCurrency(d1.NonStock_Cost),
          this.formatCurrency(d1.NonStock_Idle),

          this.formatNumber(d1.Total_Qty),
          this.formatPercent(this.calculateQtyPer(d1.Total_Qty, d.Total_Qty)),
          this.formatCurrency(d1.Total_Cost),
          this.formatCurrency(d1.Total_Idle)
        ];

        if (!isRowEmpty(level2Row)) {
          body.push({
            row: level2Row,
            indent: 1
          });
        }

        /* ===== IDLE % ROW AFTER 12+ ===== */
        if (d1.Never_sold === '12+') {

          const idleRow = [
            '',
            '',
            '',
            '',
            this.formatPercent(this.calculateStockIdle(d)),

            '',
            '',
            '',
            '',
            this.formatPercent(this.calculateNonStockIdle(d)),

            '',
            '',
            '',
            '',
            this.formatPercent(this.calculateTotalIdle(d))
          ];

          if (!isRowEmpty(idleRow)) {
            body.push({
              row: idleRow,
              indent: 2,
              isIdleRow: true
            });
          }
        }

      });

    });

    const bodyMeta = body;

    /* ================= TABLE ================= */
    autoTable(doc, {
      startY,
      head: head as any,

      body: body.map(b => ({
        ...b.row,
        raw: b
      })),

      theme: 'grid',

      styles: {
        fontSize: 9,
        cellPadding: 2,
        halign: 'right',
        lineWidth: 0.1,
        lineColor: [200, 200, 200],
        textColor: [20, 20, 20],
        valign: 'middle'
      },

      headStyles: {
        fillColor: [5, 84, 239],
        textColor: [255, 255, 255],
        fontStyle: 'bold'
      },

      didParseCell: (data: any) => {

        /* ===== HEADER ===== */
        if (data.section === 'head') {
          if (data.row.index === 1) {
            data.cell.styles.fillColor = [69, 132, 255];
            data.cell.styles.fontStyle = 'bold';
          }
        }

        if (data.section === 'body') {

          const b = bodyMeta[data.row.index] || {};

          /* ALIGN */
          data.cell.styles.halign =
            data.column.index === 0 ? 'left' : 'right';

          /* REPORT TOTAL */
          if (b.isTotalRow) {
            data.cell.styles.fillColor = [141, 180, 255];
            data.cell.styles.fontStyle = 'bold';
            return;
          }

          /* STORE ROW BG (#D9E7FF) */
          if (b.indent === 0) {
            data.cell.styles.fillColor = [217, 231, 255];
            data.cell.styles.fontStyle = 'bold';
          }

          /* ZEBRA (only child rows) */
          if (b.indent !== 0) {
            if (data.row.index % 2 === 0) {
              data.cell.styles.fillColor = [249, 251, 255];
            } else {
              data.cell.styles.fillColor = [255, 255, 255];
            }
          }

          /* INDENT */
          if (b.indent === 1 && data.column.index === 0) {
            data.cell.styles.cellPadding = { left: 6, top: 2, bottom: 2, right: 2 };
          }

          if (b.indent === 2 && data.column.index === 0) {
            data.cell.styles.cellPadding = { left: 10, top: 2, bottom: 2, right: 2 };
          }
        }
      }
    });

    return doc;
  }
  private calculateStockIdle(saleinfo: any): number {
    if (!saleinfo || !saleinfo.AgeGroupData || !saleinfo.Stock_Cost) return 0;

    const row12 = saleinfo.AgeGroupData.find((a: any) => a.Never_sold === '12+');
    return row12?.Stock_Cost ? (row12.Stock_Cost * 100) / saleinfo.Stock_Cost : 0;
  }

  private calculateNonStockIdle(saleinfo: any): number {
    if (!saleinfo || !saleinfo.AgeGroupData || !saleinfo.NonStock_Cost) return 0;

    const row12 = saleinfo.AgeGroupData.find((a: any) => a.Never_sold === '12+');
    return row12?.NonStock_Cost ? (row12.NonStock_Cost * 100) / saleinfo.NonStock_Cost : 0;
  }

  private calculateTotalIdle(saleinfo: any): number {
    if (!saleinfo || !saleinfo.AgeGroupData || !saleinfo.Total_Cost) return 0;

    const row12 = saleinfo.AgeGroupData.find((a: any) => a.Never_sold === '12+');
    return row12?.Total_Cost ? (row12.Total_Cost * 100) / saleinfo.Total_Cost : 0;
  }
  ExcelStoreNames: any = [];
  exportToExcel() {
    let storeNames: any = [];
    let store = this.storeIds
  
    storeNames = this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.filter((item: any) =>
      store.some((cat: any) => cat === item.ID.toString())
    );
    // //console.log(store,res.response.length);

    if (store.length == this.comm.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.length) {
      this.ExcelStoreNames = 'All Stores'
    } else {
      this.ExcelStoreNames = storeNames.map(function (a: any) {
        return a.storename;
      });
    }
    const PartsData = this.IndividualPartsGross.map((_arrayElement: any) =>
      Object.assign({}, _arrayElement)
    );

    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Parts Aging');
    worksheet.views = [
      {
        state: 'frozen',
        ySplit: 12, // Number of rows to freeze (2 means the first two rows are frozen)
        topLeftCell: 'A13', // Specify the cell to start freezing from (in this case, the third row)
        showGridLines: false,
      },
    ];
    worksheet.addRow('');
    const titleRow = worksheet.addRow(['Parts Aging']);
    titleRow.eachCell((cell: any, number: any) => {
      cell.alignment = { indent: 1, vertical: 'middle', horizontal: 'left' };
    });
    titleRow.font = { name: 'Arial', family: 4, size: 12, bold: true };
    titleRow.worksheet.mergeCells('A2', 'D2');
    worksheet.addRow('');

    const DateToday = this.shared.datePipe.transform(
      new Date(),
      'MM/dd/yyyy h:mm:ss a'
    );
    const DATE_EXTENSION = this.shared.datePipe.transform(
      new Date(),
      'MMddyyyy'
    );
    worksheet.addRow([DateToday]).font = { name: 'Arial', family: 4, size: 9 };

    const ReportFilter = worksheet.addRow(['Report Controls :']);
    ReportFilter.font = { name: 'Arial', family: 4, size: 10, bold: true };

    const Stores = worksheet.addRow(['Stores :']);
    Stores.getCell(1).font = { name: 'Arial', family: 4, size: 9, bold: true };
    const Groups = worksheet.getCell('B6');
    Groups.value = 'Groups :';
    const groups = worksheet.getCell('D6');
    groups.value =
      this.comm.groupsandstores.filter((val: any) => val.sg_id == this.groupId.toString())[0].sg_name;
    // this.groups == 1 ? this.comm.excelName : (this.groups == 2 ? 'Domestic' : (this.groups == 3 ? 'Import' : (this.groups == 4 ? 'Warehouse' : '-')));
    groups.font = { name: 'Arial', family: 4, size: 9 };
    const Brands = worksheet.getCell('B7');
    Brands.value = 'Brands :';
    const brands = worksheet.getCell('D7');
    brands.value = '-';
    brands.font = { name: 'Arial', family: 4, size: 9 };
    const Stores1 = worksheet.getCell('B8');
    Stores1.value = 'Stores :';
    worksheet.mergeCells('D8', 'O10');
    const stores1 = worksheet.getCell('D8');
    stores1.value =
      this.ExcelStoreNames == 0
        ? 'All Stores'
        : this.ExcelStoreNames == null
          ? '-'
          : this.ExcelStoreNames.toString().replaceAll(',', ', ');
    stores1.font = { name: 'Arial', family: 4, size: 9 };
    stores1.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };

    worksheet.addRow('');

    let dateYear = worksheet.getCell('A11');
    dateYear.value = 'As of ';
    dateYear.alignment = { vertical: 'middle', horizontal: 'center' };
    dateYear.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    dateYear.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    dateYear.border = { right: { style: 'thin' } };

    worksheet.mergeCells('B11', 'D11');
    let totalparts = worksheet.getCell('B11');
    totalparts.value = 'Stock';
    totalparts.alignment = { vertical: 'middle', horizontal: 'center' };
    totalparts.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    totalparts.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    totalparts.border = { right: { style: 'thin' } };

    worksheet.mergeCells('E11', 'G11');
    let mechanical = worksheet.getCell('E11');
    mechanical.value = 'Non-Stock';
    mechanical.alignment = { vertical: 'middle', horizontal: 'center' };
    mechanical.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    mechanical.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    mechanical.border = { right: { style: 'thin' } };

    worksheet.mergeCells('H11', 'J11');
    let retail = worksheet.getCell('H11');
    retail.value = 'Total';
    retail.alignment = { vertical: 'middle', horizontal: 'center' };
    retail.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: true,
      color: { argb: 'FFFFFF' },
    };
    retail.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '2a91f0' },
      bgColor: { argb: 'FF0000FF' },
    };
    retail.border = { right: { style: 'thin' } };



    let Headings = [
      '',

      'Qty',
      'Cost',
      'Idle',

      'Qty',
      'Cost',
      'Idle',

      'Qty',
      'Cost',
      'Idle',

    ];
    const headerRow = worksheet.addRow(Headings);
    headerRow.font = {
      name: 'Arial',
      family: 4,
      size: 9,
      bold: false,
      color: { argb: 'FFFFFF' },
    };
    headerRow.alignment = { indent: 1, vertical: 'middle', horizontal: 'center' };
    headerRow.eachCell((cell: any, number: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '788494' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = { right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    for (const d of PartsData) {
      const Data1 = worksheet.addRow([
        d.StoreName == '' ? '-' : d.StoreName == null ? '-' : d.StoreName,

        d.Stock_Qty == '' ? '-' : d.Stock_Qty == null ? '-' : d.Stock_Qty,
        d.Stock_Cost == ''
          ? '-'
          : d.Stock_Cost == null
            ? '-'
            : d.Stock_Cost,
        d.Stock_Idle == ''
          ? '-'
          : d.Stock_Idle == null
            ? '-'
            : d.Stock_Idle,

        d.NonStock_Qty == '' ? '-' : d.NonStock_Qty == null ? '-' : d.NonStock_Qty,
        d.NonStock_Cost == '' ? '-' : d.NonStock_Cost == null ? '-' : d.NonStock_Cost,
        d.NonStock_Idle == '' ? '-' : d.NonStock_Idle == null ? '-' : d.NonStock_Idle,

        d.Total_Qty == '' ? '-' : d.Total_Qty == null ? '-' : d.Total_Qty,
        d.Total_Cost == ''
          ? '-'
          : d.Total_Cost == null
            ? '-'
            : d.Total_Cost,
        d.Total_Idle == ''
          ? '-'
          : d.Total_Idle == null
            ? '-'
            : d.Total_Idle,
      ]);
      // Data1.outlineLevel = 1; // Grouping level 1
      Data1.font = { name: 'Arial', family: 4, size: 9 };
      Data1.getCell(1).alignment = {
        indent: 1,
        vertical: 'middle',
        horizontal: 'left',
      };
      Data1.eachCell((cell: any, number: any) => {
        cell.border = { right: { style: 'thin' } };
        if (number == 2 || number == 5 || number == 8) {
          cell.numFmt = '#,##0';
          cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
        } if (number > 2 && number < 5) {
          cell.numFmt = '$#,##0';
          cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
        } if (number > 5 && number < 8) {
          cell.numFmt = '$#,##0';
          cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
        } if (number > 8 && number < 11) {
          cell.numFmt = '$#,##0';
          cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
        }
        if (number != 1) {
          cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
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
      if (d.Total_PartsSale < 0) {
        Data1.getCell(2).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGross < 0) {
        Data1.getCell(4).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGross_Pace < 0) {
        Data1.getCell(5).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGrossTarget < 0) {
        Data1.getCell(6).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGross_Diff < 0) {
        Data1.getCell(7).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.ServiceGross < 0) {
        Data1.getCell(8).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.ServiceGross_Pace < 0) {
        Data1.getCell(9).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.ServiceGross_Target < 0) {
        Data1.getCell(10).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Total_PartsGross_Diff < 0) {
        Data1.getCell(11).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.PartsGross < 0) {
        Data1.getCell(12).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.PartsGross_Pace < 0) {
        Data1.getCell(13).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.PartsGross_Target < 0) {
        Data1.getCell(14).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.PartsGross_Diff < 0) {
        Data1.getCell(15).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Parts_RO < 0) {
        Data1.getCell(16).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Lost_PerDay < 0) {
        Data1.getCell(17).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      } if (d.Retention < 0) {
        Data1.getCell(18).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
      }
      if (d.AgeGroupData != undefined) {
        for (const d1 of d.AgeGroupData) {
          const Data2 = worksheet.addRow([
            d1.Never_sold == '' ? '-' : d1.Never_sold == null ? '-' : d1.Never_sold,

            d1.Stock_Qty == '' ? '-' : d1.Stock_Qty == null ? '-' : d1.Stock_Qty,
            d1.Stock_Cost == ''
              ? '-'
              : d1.Stock_Cost == null
                ? '-'
                : d1.Stock_Cost,
            d1.Stock_Idle == ''
              ? '-'
              : d1.Stock_Idle == null
                ? '-'
                : d1.Stock_Idle,

            d1.NonStock_Qty == '' ? '-' : d1.NonStock_Qty == null ? '-' : d1.NonStock_Qty,
            d1.NonStock_Cost == '' ? '-' : d1.NonStock_Cost == null ? '-' : d1.NonStock_Cost,
            d1.NonStock_Idle == '' ? '-' : d1.NonStock_Idle == null ? '-' : d1.NonStock_Idle,

            d1.Total_Qty == '' ? '-' : d1.Total_Qty == null ? '-' : d1.Total_Qty,
            d1.Total_Cost == ''
              ? '-'
              : d1.Total_Cost == null
                ? '-'
                : d1.Total_Cost,
            d1.Total_Idle == ''
              ? '-'
              : d1.Total_Idle == null
                ? '-'
                : d1.Total_Idle,

          ]);
          Data2.outlineLevel = 1; // Grouping level 2
          Data2.font = { name: 'Arial', family: 4, size: 8 };
          Data2.getCell(1).alignment = {
            indent: 2,
            vertical: 'middle',
            horizontal: 'left',
          };
          Data2.eachCell((cell: any, number: any) => {
            cell.border = { right: { style: 'thin' } };
            if (number == 2 || number == 5 || number == 8) {
              cell.numFmt = '#,##0';
              cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
            } if (number > 2 && number < 5) {
              cell.numFmt = '$#,##0';
              cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
            } if (number > 5 && number < 8) {
              cell.numFmt = '$#,##0';
              cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
            } if (number > 8 && number < 11) {
              cell.numFmt = '$#,##0';
              cell.alignment = { vertical: 'middle', horizontal: 'center', indent: 1 };
            }
            if (number != 1) {
              cell.alignment = {
                vertical: 'middle',
                horizontal: 'center',
                indent: 1,
              };
            }
          });
          if (Data2.number % 2) {
            Data2.eachCell((cell, number) => {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'e5e5e5' },
                bgColor: { argb: 'FF0000FF' },
              };
            });
          }
          if (d1.Total_PartsSale < 0) {
            Data2.getCell(2).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGross < 0) {
            Data2.getCell(4).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGross_Pace < 0) {
            Data2.getCell(5).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGrossTarget < 0) {
            Data2.getCell(6).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGross_Diff < 0) {
            Data2.getCell(7).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.ServiceGross < 0) {
            Data2.getCell(8).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.ServiceGross_Pace < 0) {
            Data2.getCell(9).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.ServiceGross_Target < 0) {
            Data2.getCell(10).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Total_PartsGross_Diff < 0) {
            Data2.getCell(11).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.PartsGross < 0) {
            Data2.getCell(12).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.PartsGross_Pace < 0) {
            Data2.getCell(13).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.PartsGross_Target < 0) {
            Data2.getCell(14).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.PartsGross_Diff < 0) {
            Data2.getCell(15).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Parts_RO < 0) {
            Data2.getCell(16).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Lost_PerDay < 0) {
            Data2.getCell(17).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          } if (d1.Retention < 0) {
            Data2.getCell(18).font = { name: 'Arial', family: 4, size: 9, color: { argb: 'FFFF0000' } }; // Font color red
          }
        }
      }
      if (d.data1 === 'Reports Total') {
        Data1.eachCell((cell) => {
          cell.font = { name: 'Arial', family: 4, size: 9, bold: true };
          // cell.border = {
          //   top: { style: 'thin' },
          //   bottom: { style: 'thin' },
          // };
        });
      }
    }

    worksheet.eachRow((row, rowIndex) => {
      row.eachCell((cell, colIndex) => {
        if (rowIndex > 1 && rowIndex < 19) { // Skip the header row
          // Apply conditional alignment based on your conditions
          if (colIndex === 1) {
            // Apply right alignment to the second column
            cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
          }
        }
      });
    });
    worksheet.getColumn(1).width = 30;
    worksheet.getColumn(2).width = 15;
    worksheet.getColumn(3).width = 15;
    worksheet.getColumn(4).width = 15;
    worksheet.getColumn(5).width = 15;
    worksheet.getColumn(6).width = 15;
    worksheet.getColumn(7).width = 15;
    worksheet.getColumn(8).width = 15;
    worksheet.getColumn(9).width = 15;
    worksheet.getColumn(10).width = 15;
    worksheet.getColumn(11).width = 15;
    worksheet.getColumn(12).width = 15;
    worksheet.getColumn(13).width = 15;
    worksheet.getColumn(14).width = 15;
    worksheet.getColumn(15).width = 15;
    worksheet.getColumn(16).width = 15;
    worksheet.getColumn(17).width = 15;
    worksheet.getColumn(18).width = 15;
    worksheet.addRow([]);

    this.shared.exportToExcel(workbook, 'Parts Aging_' + DATE_EXTENSION);

  }


}


