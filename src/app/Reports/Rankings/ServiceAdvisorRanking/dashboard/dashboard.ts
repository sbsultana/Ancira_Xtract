import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Subscription } from 'rxjs';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { NgbModal, NgbModule } from '@ng-bootstrap/ng-bootstrap';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, DateRangePicker, Stores, NgbModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss',
})
export class Dashboard {

  ServiceAdvisorData: any = [];
  NoData: boolean = false;

  Paytype: any = ['C', 'W', 'I', 'S', 'M', 'E'];
  LaborTypeVal: any = ''
  LaborState: any = 'S';
  labourType: any = 'N';
  labortypes: any = []

  columnName: any = 'Rank';
  columnState: any = 'asc';
  storeorgrp: any = 'G';
  zeroro: any = 'E';

  FromDate: any = '';
  ToDate: any = '';
  DupFromDate: any = '';
  DupToDate: any = ''
  minDate!: Date;
  maxDate!: Date;
  DateType: any = 'MTD';
  DupDateType: any = 'MTD';
  displaytime: any = '';


  storeIds!: any;
  stores: any = []
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;


  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'Y',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname,
   

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

  constructor(public shared: Sharedservice, public setdates: Setdates, private comm: common, private toast: ToastService, private ngbmodal: NgbModal) {
    this.initializeDates(this.DateType)
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
      this.getStoresandGroupsValues()
    }
    this.shared.setTitle(this.shared.common.titleName + '-Service Advisor Rankings');
    this.setHeaderData();
    this.getlabourTypesData('FL', 'N');

  }

  ngOnInit(): void { this.shared.spinner.show() }


  setHeaderData() {
    const data = {
      title: 'Service Advisor Rankings',
      stores: this.storeIds.toString(),
      labortype: this.LaborTypeVal.toString(),
      laborstate: this.LaborState,
      datetype: 'MTD',
      fromdate: this.FromDate,
      todate: this.ToDate,
      groups: this.groupId,
      storeorgroup: this.storeorgrp,
      zeroro: this.zeroro,
      Paytype: this.Paytype.toString(),
    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }
  tabClick(col_Name: any) {
    if (this.columnName !== col_Name) {
      this.columnName = col_Name;
      this.columnState = 'asc';
    }
    else {
      this.columnState = this.columnState === 'asc' ? 'desc' : 'asc';
    }
    this.GetData(this.columnName, this.columnState);
  }
  StoresData(data: any) {
    this.storeIds = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;
    this.getlabourTypesData('FR', this.labourType)

  }

  getlabourTypesData(block: any, type: any) {
    if (this.storeIds != '') {
      const obj = {
        StoreID: this.storeIds.toString(),
        type: type == 'N' ? 'A' : type
      };
      this.shared.api.postmethod(this.comm.routeEndpoint + 'GetLaborTypesTechEfficiency', obj).subscribe((res) => {
        this.spinnerLoaderlabor = false;
        this.labortypes = res.response;
        this.LaborTypeVal = res.response.map(function (a: any) {
          return a.ASD_labortype;
        });
        if (block == 'FL') {
          this.shared.spinner.show();
          this.GetData('Rank', 'asc');
        }
      })
    } else {
      // this.NoData = true
    }
  }
  alllabortypes(type: any) {
    if (type == 'Y') {
      this.LaborTypeVal = this.labortypes.map(function (a: any) {
        return a.ASD_labortype;
      });
    }
    else {
      this.LaborTypeVal = [];
    }
  }

  individualLabortypes(e: any) {
    const index = this.LaborTypeVal.findIndex(
      (i: any) => i == e.ASD_labortype
    );
    if (index >= 0) {
      this.LaborTypeVal.splice(index, 1);
    } else {
      this.LaborTypeVal.push(e.ASD_labortype);
    }
  }



  GetData(sortdata?: any, sortstate?: any) {
    this.DupFromDate = this.FromDate;
    this.DupToDate = this.ToDate
    this.DupDateType = this.DateType
    this.ServiceAdvisorData = [];
    this.shared.spinner.show();
    const obj = {
      StartDate: this.FromDate,
      EndDate: this.ToDate,
      StoreID: [...this.storeIds],
      Exp: sortdata,
      OrderType: sortstate,
      RankBy: this.storeorgrp,
      UserID: 0,
      LaborTypes: this.LaborTypeVal.toString(),
      ZeroHours: this.zeroro,
      Paytype: this.Paytype.toString(),
    };
    this.shared.api
      .postmethod(this.comm.routeEndpoint + 'GetServiceAdvisorRankingsV2', obj)
      .subscribe(
        (res: { message: any; status: number; response: string | any[] | undefined; }) => {
          if (res.status == 200) {
            if (res.response != undefined) {
              if (res.response.length > 0) {
                this.ServiceAdvisorData = res.response;
                this.shared.spinner.hide();
                this.NoData = false;
              } else {
                // this.toast.error('Empty Response', '');
                this.shared.spinner.hide();
                this.NoData = true;
              }
            } else {
              this.shared.spinner.hide();
              this.NoData = true;
            }
          } else {
            // this.toast.error(res.status, '');
            this.shared.spinner.hide();
            this.NoData = true;
          }
        },
        (error: any) => {
          // this.toast.error('502 Bad Gate Way Error', '');
          this.shared.spinner.hide();
          this.NoData = true;
        }
      );
  }

  openDetails(data: any) {
    const DetailsServicePerson = this.shared.ngbmodal.open(
      // ServiceDetailsV3Component,
      {
        // size:'xl',
        backdrop: 'static',
      }
    );
    DetailsServicePerson.componentInstance.Servicedetails = [
      {
        StartDate: this.FromDate,
        EndDate: this.ToDate,
        var1: 'Store_Name',
        var2: 'ServiceAdvisor_Name',
        var3: '',
        var1Value: data.StoreName,
        var2Value: data.ServiceAdvisor,
        var3Value: '',
        PaytypeC: 'C',
        PaytypeW: 'W',
        PaytypeI: 'I',
        DepartmentS: 'S',
        DepartmentP: 'P',            // DepartmentP: '',
        DepartmentQ: 'Q',
        DepartmentB: 'B',
        PolicyAccount: 'N',
        userName: data.ServiceAdvisor,
        Grosstype: '',
        layer: 2,
        zeroHours: this.zeroro == 'I' ? 'Y' : ''
      },
    ];
  }


  public inTheGreen(value: number): boolean {
    if (value >= 0) {
      return true;
    }
    return false;
  }

  excel!: Subscription;
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  ngAfterViewInit(): void {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.comm.pageName == 'Service Advisor Rankings') {
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
      if (!obj || obj.title !== 'Service Advisor Rankings') return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res) => {
      if (res && res.obj && res.obj.title == 'Service Advisor Rankings' && res.obj.stateEmailPdf == true) {
        this.sendEmailData(res.obj.Email, res.obj.notes, res.obj.from);
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res) => {
      if (res && res.obj && res.obj.title == 'Service Advisor Rankings' && res.obj.statePrint == true) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Service Advisor Rankings') return;
      if (obj.statePDF) {
        this.generatePDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== 'Service Advisor Rankings') return;
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
  // currentElement: string;

  // @ViewChild('scrollOne') scrollOne: ElementRef;
  // @ViewChild('scrollTwo') scrollTwo: ElementRef;

  // updateVerticalScroll(event): void {
  //   if (this.currentElement === 'scrollTwo') {
  //     this.scrollOne.nativeElement.scrollTop = event.target.scrollTop;
  //   } else if (this.currentElement === 'scrollOne') {
  //     this.scrollTwo.nativeElement.scrollTop = event.target.scrollTop;
  //   }
  // }

  // updateCurrentElement(element: 'scrollOne' | 'scrollTwo') {
  //   this.currentElement = element;
  // }

  // openDetails(Item) {
  //   this.CompleteComponentState = false;
  //   const DetailsSalesPeron = this.ngbmodal.open(SalespersonsDealsComponent, {
  //     // size:'xl',
  //     backdrop: 'static',
  //   });
  //   DetailsSalesPeron.componentInstance.Dealdetails = Item;
  //   DetailsSalesPeron.result.then(
  //     (data) => {},
  //     (reason) => {
  //       // on dismiss
  //       this.CompleteComponentState = true;
  //     }
  //   );
  // }
  ExcelStoreNames: any = []
  exportToExcel(): void {

    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Service Advisor Rankings');

    /* ================= COLORS ================= */

    const COLORS = {
      darkBlue: 'FF1E50C8',
      lightBlue: 'FFD9E7FF',
      border: 'FFCCCCCC'
    };

    /* ================= TITLE ================= */

    const title = worksheet.addRow(['Service Advisor Rankings']);
    title.font = { size: 14, bold: true };
    worksheet.mergeCells('A1:R1');

    worksheet.addRow([]);

    /* ================= FILTERS ================= */

    const formattedFromDate = this.shared.datePipe.transform(this.FromDate, 'dd-MMM-yyyy');
    const formattedToDate = this.shared.datePipe.transform(this.ToDate, 'dd-MMM-yyyy');

    const payTypeMap: any = {
      'C': 'Customer Pay',
      'W': 'Warranty',
      'I': 'Internal',
      'S': 'Sublet Gross',
      'M': 'Misc',
    };
    let formattedPayType = 'All';
    if (Array.isArray(this.Paytype) && this.Paytype.length > 0) {
      formattedPayType = this.Paytype.map(pt => payTypeMap[pt] || pt).join(', ');
    } else if (typeof this.Paytype === 'string' && this.Paytype !== '') {
      formattedPayType = payTypeMap[this.Paytype] || this.Paytype;
    }

    let formattedZeroHours = 'All';
    if (this.zeroro) {
      formattedZeroHours = this.zeroro === 'E' ? 'Exclude' :
        this.zeroro === 'I' ? 'Include' : this.zeroro;
    }

    let storeNames: any[] = [];
    const store = this.storeIds;

    storeNames = this.shared.common.groupsandstores
      .filter((v: any) => v.sg_id == this.groupId)[0]
      .Stores.filter((item: any) => store.includes(item.ID));

    if (store.length == this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores.length) {
      this.ExcelStoreNames = 'All Stores';
    } else {
      this.ExcelStoreNames = storeNames.map((a: any) => a.storename);
    }

    const filters = [
      {
        name: 'Stores :', values: this.ExcelStoreNames == null || this.ExcelStoreNames.length === 0
          ? 'All Stores' : this.ExcelStoreNames.toString().replaceAll(',', ', ')
      },
      { name: 'Time Frame:', values: `${formattedFromDate} to ${formattedToDate}` },
      { name: 'Labor Types:', values: this.LaborTypeVal.toString() },
      { name: 'Zero Hours:', values: formattedZeroHours },
      { name: 'Rank By:', values: this.storeorgrp == 'S' ? 'Store' : 'Group' },
      { name: 'Pay Type:', values: formattedPayType },
    ];

    filters.forEach(filter => {
      const row = worksheet.addRow([filter.name, filter.values]);
      row.getCell(1).font = { bold: true };

      worksheet.mergeCells(`B${row.number}:R${row.number}`);
    });

    worksheet.addRow([]);

    /* ================= HEADER ================= */
    const periodTitle = this.DupDateType == "MTD" ? "MTD" : this.DupDateType == "YTD" ? "YTD" : this.DupDateType == "QTD" ? "QTD" : this.DupDateType == "LM" ? "LAST MONTH" : this.DupDateType == "LY" ? "LAST YEAR" : `${this.DupFromDate} - ${this.DupToDate}`;

    const headerRow1 = worksheet.addRow([
      periodTitle, '', '',
      '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''
    ]);

    const headerRow2 = worksheet.addRow([
      'Rank', 'Service Advisor', 'Store Name',
      'RO', 'Hours', 'Hrs/RO', 'ELR', 'RO/Day',
      'Sales', 'Labor Sale', 'Labor Gross',
      'Parts Sale', 'Parts Gross',
      'Total Gross', 'Misc Gross', 'Sublet Gross',
      'GP %', 'Discount'
    ]);

    worksheet.mergeCells(`A${headerRow1.number}:C${headerRow1.number}`);
    worksheet.mergeCells(`D${headerRow1.number}:R${headerRow1.number}`);

    /* ================= HEADER STYLING ================= */

    headerRow1.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: COLORS.darkBlue }
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    headerRow2.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: COLORS.lightBlue }
      };
      cell.font = { bold: true, color: { argb: 'FF000000' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    [headerRow1, headerRow2].forEach(row => {
      row.eachCell(cell => {
        cell.border = {
          top: { style: 'thin', color: { argb: COLORS.border } },
          left: { style: 'thin', color: { argb: COLORS.border } },
          bottom: { style: 'thin', color: { argb: COLORS.border } },
          right: { style: 'thin', color: { argb: COLORS.border } }
        };
      });
    });

    /* ================= DATA ================= */

    this.ServiceAdvisorData.forEach((info: any) => {

      const rowData = [
        info.Rank,
        info.ServiceAdvisor,
        info.StoreName,
        info.RO ?? '-',
        info.ActualHours ?? '-',
        info.HoursPerRO ?? '-',
        info.ELR ?? '-',
        info.ROPerDay ?? '-',
        info.Sales ?? 0,
        info.LaborSale ?? 0,
        info.LaborGross ?? 0,
        info.PartsSale ?? 0,
        info.PartsGross ?? 0,
        info.TotalGross ?? 0,
        info.MiscGross ?? 0,
        info.SubletGross ?? 0,
        info.GP ?? '-',
        info.Discount ?? 0
      ];

      const row = worksheet.addRow(rowData);

      row.eachCell((cell, colNumber) => {

        if ([9, 10, 11, 12, 13, 14, 15, 16, 18].includes(colNumber)) {
          cell.alignment = { horizontal: 'right', vertical: 'middle' };

          if (typeof cell.value === 'number') {
            cell.numFmt = '"$"#,##0.00';
          }
        }
        else if ([2, 3].includes(colNumber)) {
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        }
        else {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }

        cell.border = {
          top: { style: 'thin', color: { argb: COLORS.border } },
          left: { style: 'thin', color: { argb: COLORS.border } },
          bottom: { style: 'thin', color: { argb: COLORS.border } },
          right: { style: 'thin', color: { argb: COLORS.border } }
        };

      });
    });

    /* ================= COLUMN WIDTH ================= */

    worksheet.columns = [
      { width: 15 },
      { width: 25 },
      { width: 25 },
      { width: 10 },
      { width: 10 },
      { width: 10 },
      { width: 10 },
      { width: 10 },
      { width: 14 },
      { width: 14 },
      { width: 14 },
      { width: 14 },
      { width: 14 },
      { width: 16 },
      { width: 14 },
      { width: 14 },
      { width: 10 },
      { width: 14 }
    ];

    /* ================= EXPORT ================= */

    workbook.xlsx.writeBuffer().then(() => {
      this.shared.exportToExcel(workbook, 'Service Advisor Rankings');
    });
  }
  private createPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });

    doc.setFontSize(14);
    doc.text('Service Advisor Rankings', 14, 12);

    const periodTitle = this.DupDateType == "MTD" ? "MTD" : this.DupDateType == "YTD" ? "YTD" : this.DupDateType == "QTD" ? "QTD" : this.DupDateType == "LM" ? "LAST MONTH" : this.DupDateType == "LY" ? "LAST YEAR" : `${this.DupFromDate} - ${this.DupToDate}`;

    const headers = [
      [
        { content: periodTitle, colSpan: 3 },
        { content: '', colSpan: 15 }
      ],
      [
        'Rank', 'Service Advisor', 'Store Name',
        'RO', 'Hours', 'Hrs/RO', 'ELR', 'RO/Day',
        'Sales', 'Labor Sale', 'Labor Gross',
        'Parts Sale', 'Parts Gross',
        'Total Gross', 'Misc Gross', 'Sublet Gross',
        'GP %', 'Discount'
      ]];

    // ✅ Rows Mapping
    const rows = this.ServiceAdvisorData.map((x: any) => [
      x.Rank,
      x.ServiceAdvisor,
      x.StoreName,
      x.RO ?? '-',
      x.ActualHours ?? '-',
      x.HoursPerRO ?? '-',
      x.ELR ?? '-',
      x.ROPerDay ?? '-',
      x.Sales ?? 0,
      x.LaborSale ?? 0,
      x.LaborGross ?? 0,
      x.PartsSale ?? 0,
      x.PartsGross ?? 0,
      x.TotalGross ?? 0,
      x.MiscGross ?? 0,
      x.SubletGross ?? 0,
      x.GP ?? '-',
      x.Discount ?? 0
    ]);

    autoTable(doc, {
      head: headers,
      body: rows,
      startY: 20,

      theme: 'plain',

      styles: {
        fontSize: 7,
        cellPadding: 1.8,
        valign: 'middle',
        halign: 'center',
        lineWidth: 0.3, lineColor: [200, 200, 200]
      },
      headStyles: {
        fontStyle: 'bold',
        halign: 'center',
        lineWidth: 0.3, lineColor: [200, 200, 200]
      },

      tableWidth: 'auto',
      margin: { left: 8, right: 8 },

      columnStyles: {
        0: { cellWidth: 'auto' },

        1: { cellWidth: 'auto', halign: 'left' },
        2: { cellWidth: 'auto', halign: 'left' },

        3: { cellWidth: 'auto' },
        4: { cellWidth: 'auto' },
        5: { cellWidth: 'auto' },
        6: { cellWidth: 'auto' },
        7: { cellWidth: 'auto' },
        8: { cellWidth: 'auto', halign: 'right' },
        9: { cellWidth: 'auto', halign: 'right' },
        10: { cellWidth: 'auto', halign: 'right' },
        11: { cellWidth: 'auto', halign: 'right' },
        12: { cellWidth: 'auto', halign: 'right' },
        13: { cellWidth: 'auto', halign: 'right' },
        14: { cellWidth: 'auto', halign: 'right' },
        15: { cellWidth: 'auto', halign: 'right' },

        16: { cellWidth: 10 }, // GP %
        17: { cellWidth: 'auto', halign: 'right' } // Discount
      },

      didParseCell: (data: any) => {

        // 🔵 TOP HEADER
        if (data.section === 'head' && data.row.index === 0) {

          if (!data.cell.raw) return;

          data.cell.styles.fillColor = [30, 80, 200];
          data.cell.styles.textColor = 255;
          data.cell.styles.fontStyle = 'bold';

          // ✅ ADD RIGHT DIVIDER AFTER PERIOD TITLE (colSpan = 3 → index 2)
          if (data.column.index === 2) {
            data.cell.styles.lineWidth = { right: 0.6 };
            data.cell.styles.lineColor = [255, 255, 255];
          }
        }

        // ⚪ SECOND HEADER
        if (data.section === 'head' && data.row.index === 1) {
          data.cell.styles.fillColor = '#d9e7ff';
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.textColor = '#000';
        }

        // 💰 FORMAT NUMBERS
        if (data.section === 'body') {
          const col = data.column.index;
          const value = data.cell.raw;

          const currencyCols = [8, 9, 10, 11, 12, 13, 14, 15, 17];

          if (currencyCols.includes(col) && !isNaN(value)) {
            const formatted = `$${Math.abs(value).toLocaleString()}`;
            data.cell.text = [value < 0 ? `-${formatted}` : formatted];
          }
        }
      }
    });
    return doc;
  }

  generatePDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createPDF();
      doc.save('Service Advisor Rankings.pdf'); // ✅ only here
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
      const pdfFile = this.blobToFile(pdfBlob, 'Service Advisor Rankings.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', 'Service Advisor Rankings');
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
      'type': 'M', 'others': 'Y', 
    };

  }


  //------Reports-----------//
  activePopover: number = -1;
  initializeDates(type: any) {
    let dates: any = this.setdates.setDates(type)
    this.FromDate = dates[0];
    this.ToDate = dates[1];
    localStorage.setItem('time', type);
    this.setDates(type)
  }

  updatedDates(data: any) {
    console.log(data);
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

  spinnerLoaderlabor: boolean = false;


  zeroros(block: any, e: any) {
    this.zeroro = [];
    this.zeroro.push(e);
  }
  storeorgroups(block: any, e: any) {
    this.storeorgrp = [];
    this.storeorgrp.push(e);
  }
  multipleorsingle(block: any, e: any) {
    if (block == 'PT') {
      const index = this.Paytype.findIndex((i: any) => i == e);
      if (index >= 0) {
        this.Paytype.splice(index, 1);
      } else {
        this.Paytype.push(e);
      }
    }
  }
  viewreport() {
    this.activePopover = -1
    if ((!this.storeIds || this.storeIds.length === 0) ) {
      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
    }
    else if (!this.LaborTypeVal || this.LaborTypeVal.length === 0) {
      this.toast.show('Please select any labor type', 'warning', 'Warning');
      return;
    }
    else if (!this.Paytype || (Array.isArray(this.Paytype) && this.Paytype.length === 0)) {
      this.toast.show('Please select any Pay Type', 'warning', 'Warning');
      return;
    }
    else {
      this.setHeaderData()
      this.GetData(this.columnName, this.columnState);

    }

  }

}
