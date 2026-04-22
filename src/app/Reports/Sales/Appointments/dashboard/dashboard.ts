import { Component, OnInit, ViewChild, ElementRef, HostListener } from '@angular/core';
import { Sharedservice } from '../../../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { Setdates } from '../../../../Core/Providers/SetDates/setdates';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { common } from '../../../../common';
// import { Stores } from '../../../../CommonFilters/stores/stores';
import { DateRangePicker } from '../../../../CommonFilters/date-range-picker/date-range-picker';
import { Subscription } from 'rxjs';
import { Stores } from '../../../../CommonFilters/stores/stores';
import { FilterPipe } from '../../../../Core/Providers/filterpipe/filter.pipe';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [SharedModule, BsDatepickerModule, FilterPipe, DateRangePicker, Stores],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {
  Appointmentdata: any = [];
  NoData: any = false;
  appointment: any = [];
  selectedStore = '0';
  selectedstrid = 0;
  // stores: any = [];
  PageCount = 1;
  LastCount!: boolean;
  Pagination: boolean = false;
  storeIds: any = '0';

  // groups: any = 1
  filteredCustomers: any;
  Appointmentsearch: any;
  searcdata: any;
  FromDate: any = '';
  ToDate: any = '';

  minDate!: Date;
  maxDate!: Date;
  todaytitle: any;
  reportOpenSub!: Subscription;
  reportGetting!: Subscription;

  // FromDate: any;
  // ToDate: any;
  DateType: string = 'C';
  displaytime: any = new Date();
  Dates: any = {
    'FromDate': this.FromDate, 'ToDate': this.ToDate, "MaxDate": this.maxDate, 'MinDate': this.minDate, 'DateType': this.DateType, 'DisplayTime': this.displaytime,
    Types: []
  }
  StoreValues: any = '0';
  popup: any = [{ type: 'Popup' }];
  groups: any = 1;
  gridvisibility: any;
  bsRangeValue!: Date[];
  groupsArray: any = [];
  storename: any = ''
  storecount: any = null;
  storedisplayname: any = '';
  groupName: any = '';
  groupId: any = 0;
  stores: any = [];



  // activePopover: number = -1;
  storesFilterData: any = {
    'groupsArray': this.groupsArray, 'groupId': this.groupId, 'storesArray': this.stores, 'storeids': '1', 'type': 'M', 'others': 'Y',
    'groupName': this.groupName, 'storename': this.storename, storecount: null, 'storedisplayname': this.storedisplayname,
 
  };
  // @HostListener('document:click', ['$event'])
  // onDocumentClick(event: MouseEvent) {
  //   const clickedInside = (event.target as HTMLElement).closest('.dropdown-toggle, .dropdown-menu , .timeframe, .reportstores-card');
  //   if (!clickedInside) {
  //     this.activePopover = -1;
  //   }
  // }
  constructor(public shared: Sharedservice, public setdates: Setdates, private toast: ToastService,

  ) {
    if (localStorage.getItem('userInfo') != null && localStorage.getItem('userInfo') != undefined) {
      if (localStorage.getItem('flag') == 'V') {
        this.StoreValues = [];
        console.log(JSON.parse(localStorage.getItem('userInfo')!), JSON.parse(localStorage.getItem('userInfo')!).user_Info, 'Widget Stores............');
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.groupid
        JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store.indexOf(',') > 0 ?
          this.StoreValues = JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store.split(',') :
          this.StoreValues.push(JSON.parse(localStorage.getItem('userInfo')!).WidgetData.store)
        localStorage.setItem('flag', 'M')
      }
      else {
        this.groupId = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Preferences
        this.StoreValues = JSON.parse(localStorage.getItem('userInfo')!).user_Info.Storeids.split(',');

      }
      if (this.StoreValues.toString().indexOf(',') > 0) {
        this.gridvisibility = 'DL';
      } else {
        this.gridvisibility = 'SL';
      }
    }

    localStorage.setItem('time', 'C');

    this.titleSetting()
    this.shared.setTitle(this.shared.common.titleName + '- Sales Appointments');

    this.todaytitle = this.shared.datePipe.transform(new Date(), 'MM.dd.yyyy');
    let today = new Date();
    const tomorrow = new Date();
    tomorrow.setDate(today.getDate());
    this.FromDate = this.shared.datePipe.transform(today, 'MM.dd.yyyy');
    this.ToDate = this.shared.datePipe.transform(tomorrow, 'MM.dd.yyyy');
    this.setDates(this.DateType)
    this.setHeaderData()

    this.GetAppointment();
    // }
  }
  dynamicTitle = 'Sales Appointments'
  titleSetting() {
    console.log(this.todaytitle, this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy'), this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy'));
    this.dynamicTitle = 'Sales Appointments'
    this.setHeaderData()
  }

  setHeaderData() {
    const data = {
      title: this.dynamicTitle,
      stores: this.storeIds,
      groups: this.groupId,
      fromdate: this.FromDate,
      todate: this.ToDate,
    

    };
    this.shared.api.SetHeaderData({
      obj: data,
    });
  }
  ngOnInit(): void {
    this.getStores();
  }


  searchQuery: string = '';
  isDesc: boolean = false;
  column: string = '';

  sort(property: any) {
    this.isDesc = !this.isDesc; //change the direction
    this.column = property;
    let direction = this.isDesc ? 1 : -1;
    // // console.log(property)
    this.Appointmentdata.sort(function (a: any, b: any) {
      if (a[property] < b[property]) {
        return -1 * direction;
      } else if (a[property] > b[property]) {
        return 1 * direction;
      } else {
        return 0;
      }
    });
  }
  getStoresandGroupsValues() {
    this.storesFilterData.groupsArray = this.groupsArray;
    this.storesFilterData.groupId = this.groupId;
    this.storesFilterData.storesArray = this.stores;
    this.storesFilterData.storeids = this.StoreValues;
    this.storesFilterData.groupName = this.groupName;
    this.storesFilterData.storename = this.storename;
    this.storesFilterData.storecount = this.storecount;
    this.storesFilterData.storedisplayname = this.storedisplayname;
  
    this.storesFilterData = {
      groupsArray: this.groupsArray,
      groupId: this.groupId,
      storesArray: this.stores,
      storeids: this.StoreValues,
      groupName: this.groupName,
      storename: this.storename,
      storecount: this.storecount,
      storedisplayname: this.storedisplayname,
      'type': 'M', 'others': 'Y',
    };

    // this.setHeaderData();
    // this.GetData();

  }
  StoresData(data: any) {
    this.StoreValues = data.storeids;
    this.groupId = data.groupId;
    this.storename = data.storename;
    this.groupName = data.groupName;
    this.storecount = data.storecount;
    this.storedisplayname = data.storedisplayname;

  }
  getStores() {
    this.selectedstorevalues = [];
  }

  GetAppointment() {
    this.Appointmentdata = [];
    this.NoData = false;
    this.shared.spinner.show();
    let obj = {};
    if (localStorage.getItem('DetailsObject') != null) {
      const InvObj = JSON.parse(localStorage.getItem('DetailsObject')!);
      obj = {
        DealerId: InvObj.dataobj.Data1,
        StartDate: InvObj.dataobj.Data2,
        EndDate: InvObj.dataobj.Data2,
      };
      this.FromDate = InvObj.dataobj.Data2
      this.ToDate = InvObj.dataobj.Data2
      this.storeIds = InvObj.dataobj.Data1
      console.log(this.todaytitle);
      this.dynamicTitle = 'Sales Appointments'
      this.setHeaderData()
    } else {
      obj = {
        DealerId: [...this.StoreValues],
        StartDate: this.shared.datePipe.transform(this.FromDate, 'MM/dd/yyyy'),
        EndDate: this.shared.datePipe.transform(this.ToDate, 'MM/dd/yyyy'),
      };
      console.log(obj);

    }
    this.shared.api.postmethod(this.shared.common.routeEndpoint + 'GetVSAppointmentsData', obj).subscribe(
      (res: { message: any; status: number; response: string | any[] | undefined; }) => {
        if (res.status == 200) {
          this.Appointmentdata = []
          this.Appointmentsearch = []
          this.appointment = []
          if (res.response != undefined) {
            if (res.response.length > 0) {
              this.shared.spinner.hide();
              this.Appointmentdata = [...res.response];
              this.Appointmentsearch = res.response;
              console.log(this.Appointmentdata, 'Appointment');
              this.appointment = res.response;
              if (this.appointment.length == 0) {
                this.Pagination = false;
                this.LastCount = false;
                this.NoData = true;
              } else {
                this.Pagination = true;
                this.LastCount = true;
              }
            }
            else {
              this.shared.spinner.hide();
              this.NoData = true;


            }
          }
          else {
            // this.toast.error('Empty Response');
            this.NoData = true;

            this.shared.spinner.hide();
          }

          localStorage.removeItem('DetailsObject');
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
  pdf!: Subscription;
  print!: Subscription;
  email!: Subscription;
  excel!: Subscription;
  ngAfterViewInit() {
    this.shared.api.getStores().subscribe((res: any) => {
      if (this.shared.common.pageName == this.dynamicTitle) {
        if (res.obj.storesData != undefined) {
          this.groupsArray = res.obj.storesData;
          this.stores = this.shared.common.groupsandstores.filter((v: any) => v.sg_id == this.groupId)[0].Stores;
          this.StoreValues.length == this.stores.length ? this.groupName = this.stores[0].sg_name : this.groupName = ''
          this.StoreValues.length == 1 ? this.storename = this.stores.filter((e: any) => e.ID == this.StoreValues)[0].storename : this.storename = ''
          this.getStoresandGroupsValues()
        }
      }
    })
    this.excel = this.shared.api.getExportToExcelAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== this.dynamicTitle) return;
      if (obj.state) {
        this.exportToExcel();
      }
    });
    this.print = this.shared.api.getExportToPrintAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== this.dynamicTitle) return;
      if (obj.statePrint) {
        this.printPDF();
      }
    });
    this.pdf = this.shared.api.getExportToPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== this.dynamicTitle) return;
      if (obj.statePDF) {
        this.downloadPDF();
      }
    });
    this.email = this.shared.api.getExportToEmailPDFAllReports().subscribe((res: any) => {
      const obj = res?.obj;
      if (!obj || obj.title !== this.dynamicTitle) return;
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

  Scrollpercent: any = 0;
  scrollCurrentposition: any = 0
  @ViewChild('scrollcent') scrollcent!: ElementRef;

  updateVerticalScroll(event: any): void {

    this.scrollCurrentposition = event.target.scrollTop
    const scrollDemo = document.querySelector('#scrollcent') as HTMLElement;
    this.Scrollpercent = Math.round(
      (event.target.scrollTop /
        (event.target.scrollHeight - scrollDemo.clientHeight)) *
      100
    );
  }

  //------Reports-----------//
  activePopover: number = -1;
  Teams = 'S';
  selectedstorevalues: any = [];
  labortypes: any = [];
  selectedlabortypevalues: any = [];
  Performance: any = 'Load';
  changesvalues: any = [];
  AllLabortypes: boolean = true;
  toporbottom: any = ['T'];

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
    this.bsRangeValue = [new Date(this.FromDate), new Date(this.ToDate)];

  }
  open() {
    this.Dates = { ...this.Dates }
    this.setDates(this.DateType)
  }
  setDates(type: any) {
    this.DateType == 'C' ? this.displaytime = ' (  ' + this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy') + ' - ' + this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy') + ' ) ' :
      this.displaytime = 'Time Frame (' + this.Dates.Types.filter((val: any) => val.code == type)[0].name + ')';
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

  viewreport() {
    if (!this.StoreValues || this.StoreValues.length === 0) {

      this.toast.show('Please Select Atleast One Store', 'warning', 'Warning');
      return;
    }

    this.activePopover = -1
    this.titleSetting()
    this.setHeaderData()
    this.GetAppointment()

  }
  getExcelStoreLabel(): string {
    // 0 selected
    if (!this.StoreValues || this.StoreValues.length === 0) {
      return 'Selected (0)';
    }

    // All selected → show all store names
    if (this.StoreValues.length === this.stores.length) {
      return this.stores
        .map((s: any) => s.storename)
        .join(', ');
    }

    // Single selected → store name
    if (this.StoreValues.length === 1) {
      const store = this.stores.find(
        (s: any) => s.ID === this.StoreValues[0]
      );
      return store ? store.storename : '-';
    }

    // Multiple (not all)
    return `Selected (${this.StoreValues.length})`;
  }


  ExcelStoreNames: any = []
  exportToExcel(): void {

    const workbook = this.shared.getWorkbook();
    const worksheet = workbook.addWorksheet('Sales Appointments');

    /* ================= STORE VALUE ================= */

    let storeValue = '';

    if (!this.StoreValues || this.StoreValues.length === 0) {
      storeValue = this.stores.map((s: any) => s.storename).join(', ');
    }
    else if (this.StoreValues.length === this.stores.length) {
      storeValue = this.stores.map((s: any) => s.storename).join(', ');
    }
    else {
      storeValue = this.stores
        .filter((s: any) => this.StoreValues.includes(s.ID))
        .map((s: any) => s.storename)
        .join(', ');
    }

    /* ================= TITLE ================= */

    worksheet.addRow([]);
    const titleRow = worksheet.addRow([this.dynamicTitle]);
    worksheet.mergeCells(`A2:S2`);

    titleRow.font = { size: 14, bold: true };

    /* ================= DATE ================= */

    const from = this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy');
    const to = this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy');

    worksheet.addRow([]);

    /* ================= FILTERS ================= */

    const filters = [
      { name: 'Stores:', values: storeValue },
      { name: 'Time Frame:', values: `${from} - ${to}` }
    ];

    filters.forEach((f) => {
      const row = worksheet.addRow([f.name, f.values]);

      row.getCell(1).font = { bold: true };
      row.getCell(2).font = { color: { argb: 'FF000000' } };

      worksheet.mergeCells(`B${row.number}:S${row.number}`);
    });

    worksheet.addRow([]);

    /* ================= HEADER ================= */

    const headers = [
      'Dealer', 'Time', 'Date', 'Stock Number', 'Year', 'Make', 'Model', 'Customer',
      'Assigned User', 'Appt ID', 'Appt Status', 'Lead Type', 'Status',
      'Sales Representative', 'BD Agent', 'Manager', 'Created', 'Source', 'Contacted'
    ];

    const headerRow = worksheet.addRow(headers);

    headerRow.eachCell((cell: any) => {
      cell.font = {
        bold: true,
        color: { argb: 'FFFFFFFF' },
        name: 'Arial',
        size: 10
      };

      // 🔷 SAME AS PDF HEADER COLOR
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0554EF' } // RGB(5,84,239)
      };

      cell.alignment = {
        horizontal: 'center',
        vertical: 'middle'
      };

      cell.border = {
        top: { style: 'thin', color: { argb: 'FFC8C8C8' } },
        left: { style: 'thin', color: { argb: 'FFC8C8C8' } },
        bottom: { style: 'thin', color: { argb: 'FFC8C8C8' } },
        right: { style: 'thin', color: { argb: 'FFC8C8C8' } }
      };
    });


    /* ================= DATA ================= */

    const bindingHeaders = [
      'Dealer', 'Appt Start Time', 'Appt Start Date', 'Stock Number', 'Year', 'Make',
      'Model', 'Customer', 'Assigned User', 'Appt ID', 'Appt Status', 'Lead Type',
      'Status', 'Sales Representative', 'BD Agent', 'Manager', 'Created', 'Source', 'Contacted'
    ];

    this.Appointmentdata.forEach((info: any) => {

      const rowData = bindingHeaders.map(key => {

        let val = info[key];

        if (key === 'Appt Start Date' || key === 'Created') {
          if (val) {
            const d = new Date(val);
            const mm = String(d.getMonth() + 1).padStart(2, '0');
            const dd = String(d.getDate()).padStart(2, '0');
            const yyyy = d.getFullYear();
            val = `${mm}.${dd}.${yyyy}`; // match PDF format
          } else {
            val = '-';
          }
        }

        return (val === 0 || val == null || val === '') ? '-' : val;
      });

      const row = worksheet.addRow(rowData);

      row.eachCell((cell: any, colNumber: number) => {

        // // 🔷 ALIGNMENT LIKE PDF
        // if (colNumber === 2 || colNumber === 3) {
        //   cell.alignment = { horizontal: 'center' }; // Time & Date
        // }
        // else if (colNumber === 4) {
        //   cell.alignment = { horizontal: 'left' }; // Stock
        // }
        // else {
        //   cell.alignment = { horizontal: 'center' };
        // }

        // 🔷 GRID LIKE PDF
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          right: { style: 'thin', color: { argb: 'FFE0E0E0' } }
        };

        cell.font = { name: 'Arial', size: 9 };
      });
    });

    /* ================= AUTO WIDTH ================= */

    worksheet.columns.forEach((column: any, index: number) => {

      if (index === 0) {
        column.width = 50; // First column
      } else {
        column.width = 18; // 
      }
    
    });
    /* ================= EXPORT ================= */

    workbook.xlsx.writeBuffer().then((buffer: any) => {
      this.shared.exportToExcel(workbook, 'Sales Appointments');
    });
  }

  private createSalesAppointmentsPDF(): jsPDF {

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a2' });

    // ===== TITLE =====
    doc.setFontSize(14);
    const title = `${this.dynamicTitle}`;
    doc.text(title, 14, 12);

    // ===== DATE =====
    const from = this.shared.datePipe.transform(this.FromDate, 'MM.dd.yyyy');
    const to = this.shared.datePipe.transform(this.ToDate, 'MM.dd.yyyy');

    doc.setFontSize(10);
    doc.text(`${from} - ${to}`, 14, 18);

    let startY = 22;

    // ===== SAFETY =====
    if (!this.Appointmentdata || this.Appointmentdata.length === 0) {
      doc.text('No Data Available', 14, startY);
      return doc;
    }



    // =========================
    // 🔷 MAIN TABLE
    // =========================

    const showDealer =
      this.storeIds?.toString().indexOf(',') > 0 || this.storeIds == 0;

    const head = [[
      ...(showDealer ? ['Dealer'] : []),
      'Time',
      'Date',
      'Stock #',
      'Year',
      'Make',
      'Model',
      'Customer',
      'Assigned User',
      'Appt ID',
      'Appt Status',
      'Lead Type',
      'Status',
      'Sales Rep',
      'BD Agent',
      'Manager',
      'Created',
      'Source',
      'Contacted'
    ]];

    const rows = (this.Appointmentdata || []).map((d: any) => [
      ...(showDealer ? [d["Dealer"] || '-'] : []),
      d["Appt Start Time"] || '-',
      d["Appt Start Date"]
        ? this.shared.datePipe.transform(d["Appt Start Date"], 'MM.dd.yyyy')
        : '-',
      d["Stock Number"] || '-',
      d["Year"] || '-',
      d["Make"] || '-',
      d["Model"] || '-',
      d["Customer"] || '-',
      d["Assigned User"] || '-',
      d["Appt ID"] || '-',
      d["Appt Status"] || '-',
      d["Lead Type"] || '-',
      d["Status"] || '-',
      d["Sales Representative"] || '-',
      d["BD Agent"] || '-',
      d["Manager"] || '-',
      d["Created"]
        ? this.shared.datePipe.transform(d["Created"], 'MM.dd.yyyy')
        : '-',
      d["Source"] || '-',
      d["Contacted"] || '-'
    ]);

    autoTable(doc, {
      startY,
      head,
      body: rows,
      theme: 'grid',

      styles: {
        fontSize: 7,
        cellPadding: 1.5,
        lineWidth: 0.3,
        lineColor: [200, 200, 200]
      },

      headStyles: {
        fillColor: [5, 84, 239],
        textColor: 255,
        halign: 'center',
        lineWidth: 0.3,
        lineColor: [200, 200, 200]
      },

      didParseCell: (data: any) => {
        if (data.section === 'body') {

          const offset = showDealer ? 1 : 0;

          // Align like UI
          if (data.column.index === (0 + offset)) data.cell.styles.halign = 'center'; // time
          if (data.column.index === (1 + offset)) data.cell.styles.halign = 'center'; // date
          if (data.column.index === (3 + offset)) data.cell.styles.halign = 'left';   // stock
        }
      }
    });

    return doc;
  }

  downloadPDF() {
    this.shared.spinner.show();
    try {
      const doc = this.createSalesAppointmentsPDF();
      doc.save(this.dynamicTitle + '.pdf');
    } catch (e) {
      console.error(e);
      this.toast.show('Error generating PDF', 'danger', 'Error');
    } finally {
      this.shared.spinner.hide();
    }
  }
  printPDF() {
    try {
      const doc = this.createSalesAppointmentsPDF();
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
      const doc = this.createSalesAppointmentsPDF()
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
      const pdfFile = this.blobToFile(pdfBlob, this.dynamicTitle + '.pdf');
      const formData = new FormData();
      formData.append('to_email', Email);
      formData.append('subject', this.dynamicTitle);
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
}




