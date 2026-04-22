import { DatePipe } from '@angular/common';
import { Component, EventEmitter, Input, Output, SimpleChanges, ViewChild } from '@angular/core';
import { BsDatepickerConfig, BsDatepickerModule, BsDaterangepickerDirective } from 'ngx-bootstrap/datepicker';
import { Setdates } from '../../Core/Providers/SetDates/setdates';
import { SharedModule } from '../../Core/Providers/Shared/shared.module';

@Component({
  selector: 'app-date-range-picker',
  imports: [SharedModule, BsDatepickerModule],
  templateUrl: './date-range-picker.html',
  styleUrl: './date-range-picker.scss'
})
export class DateRangePicker {
  @Input() Dates: any = {}
  @Output() updatedDates = new EventEmitter();

  FromDate: any;
  ToDate: any;
  Data: any = [];
  minDate: any;
  maxDate: any
  bsConfig: Partial<BsDatepickerConfig> = {
    dateInputFormat: 'MMMM/YYYY',
    minMode: 'month'
  };
  custom: boolean = false;
  bsValue = new Date();
  bsRangeValue!: Date[];
  DateType: any = 'MTD'
  @ViewChild('datepicker') datepicker!: BsDaterangepickerDirective;

  groupName: any = ''
  groupsArray: any = [];
  storename: any = '';
  groupId: any = [];
  type: any = '';
  others: any = '';
  storedisplayname: any = ''
  storecount: any = null
  customeVisibility: any = 'Y'

  constructor(private datepipe: DatePipe, private datesSrvc: Setdates) { }

  ngOnChanges(changes: SimpleChanges) {
    console.log(changes, '...................');
    if (changes['Dates'] && changes['Dates'].currentValue) {
      this.Data = changes['Dates'].currentValue.Types;
      this.FromDate = changes['Dates'].currentValue.FromDate;
      this.ToDate = changes['Dates'].currentValue.ToDate;
      this.minDate = changes['Dates'].currentValue.MinDate;
      this.maxDate = changes['Dates'].currentValue.MaxDate;
      this.DateType = changes['Dates'].currentValue.DateType;
      this.bsRangeValue = [new Date(this.FromDate), new Date(this.ToDate)];
      this.customeVisibility =  changes['Dates'].currentValue?.custom ? changes['Dates'].currentValue?.custom : 'Y';
      console.log(this.Data, '........');

      if (this.Data == undefined || this.Data.length == 0) {
        setTimeout(() => {
          this.datepicker.hide();

          this.openbardate()
        }, 500);
      }
    }
  }
  SetDates(type: any, block?: any) {
    //  alert('Hi')
    this.DateType = type;
    // localStorage.setItem('time', this.DateType);
    if (block == 'B') {
      this.datepicker.show();
    } else {
      this.custom = false;
      let dates: any = this.datesSrvc.setDates(type)
      this.FromDate = dates[0]
      this.ToDate = dates[1]
      this.bsRangeValue = [new Date(this.FromDate), new Date(this.ToDate)];
      this.Dates.FromDate = this.FromDate;
      this.Dates.ToDate = this.ToDate;
      this.Dates.DateType = this.DateType;
      this.displaytime()
    }
  }
  openbardate() {
    // alert('HI')
    this.datepicker.show();
  }
  displaytime() {
    console.log(this.DateType);

    if (this.DateType == 'C') {
      this.Dates.DisplayTime = ' (  ' + this.datepipe.transform(this.FromDate, 'MM.dd.yyyy') + ' - ' + this.datepipe.transform(this.ToDate, 'MM.dd.yyyy') + ' ) '
      console.log(this.FromDate, this.ToDate, ' Display Time Function ')

    } else {
      this.Dates.DisplayTime = ' (  ' + this.Dates.Types.filter((val: any) => val.code == this.DateType)[0].name + '  )';
    }
    this.Dates.FromDate = this.FromDate;
    this.Dates.ToDate = this.ToDate;
    this.Dates.DateType = this.DateType;
    console.log(this.Dates);

    this.updatedDates.emit(this.Dates)
  }

  dateRangeCreated($event: any) {

    if ($event !== null) {
      let startDate = $event[0].toJSON();
      let endDate = $event[1].toJSON();
      this.FromDate = this.datepipe.transform(startDate, 'MM-dd-yyyy');
      this.ToDate = this.datepipe.transform(endDate, 'MM-dd-yyyy');
      if (this.DateType == 'C') {
        this.custom = true;
      }
      this.displaytime()
    }
  }
}

