import { Injectable } from '@angular/core';

import { common } from '../../../common';
import { Sharedservice } from '../Shared/sharedservice';

@Injectable({
  providedIn: 'root'
})
export class Setdates {

  constructor(public shared: Sharedservice, private comm: common) { }

  setDates(type: any, selecteddate?: any) {
    let today = new Date();
    let enddate = new Date()

    if ((this.comm.pageName == 'Fixed Gross GL' || this.comm.pageName == 'Parts Gross GL' || this.comm.pageName == 'Fixed Summary GL') && today.getDate() < 5) {
      enddate = new Date(today.setDate(today.getDate()));
    } else {
      if (selecteddate == undefined) {
        today = new Date()
        enddate = new Date(today.setDate(today.getDate() - 1));

      } else {
        today = new Date(selecteddate)
        enddate = new Date(today.setDate(today.getDate()));

      }
    }
    console.log(selecteddate, enddate);




    if (type == 'MTD') {
      if (selecteddate != undefined) {
        console.log('Selected Date', selecteddate);
        if (selecteddate.getMonth() == new Date().getMonth() && selecteddate.getFullYear() == new Date().getFullYear()) {
          return [('0' + (enddate.getMonth() + 1)).slice(-2) + '-01' + '-' + enddate.getFullYear(),
          ('0' + (enddate.getMonth() + 1)).slice(-2) + '-' + ('0' + new Date().getDate()).slice(-2) + '-' + enddate.getFullYear()]
        } else {
          var lastDayOfMonth = new Date(selecteddate.getFullYear() - 1, selecteddate.getMonth() + 1, 0);
          console.log(lastDayOfMonth, 'Last Day Of Month');
          return [('0' + (enddate.getMonth() + 1)).slice(-2) + '-01' + '-' + enddate.getFullYear(), ('0' + (enddate.getMonth() + 1)).slice(-2) + '-' + ('0' + lastDayOfMonth.getDate()).slice(-2) + '-' + (enddate.getFullYear())]
        }
      } else {
        console.log('MTD', enddate);
        return [('0' + (enddate.getMonth() + 1)).slice(-2) + '-01' + '-' + enddate.getFullYear(),
        ('0' + (enddate.getMonth() + 1)).slice(-2) + '-' + ('0' + enddate.getDate()).slice(-2) + '-' + enddate.getFullYear()]
      }

    }

    if (type == 'QTD') {
      if (enddate.getMonth() == 0) {
        return ['10-01-' + (enddate.getFullYear() - 1), '12-31-' + (enddate.getFullYear() - 1)]
      } else {
        let d = new Date(enddate)
        d.setMonth(d.getMonth() - 3)
        let localstringdate = d.toISOString();
        return [this.shared.datePipe.transform(localstringdate, 'MM-dd-yyyy'), ('0' + (enddate.getMonth() + 1)).slice(-2) + '-' + ('0' + enddate.getDate()).slice(-2) + '-' + enddate.getFullYear()]
      }
    }
    if (type == 'YTD') {
      return [('0' + 1).slice(-2) + '-01' + '-' + enddate.getFullYear(), ('0' + (enddate.getMonth() + 1)).slice(-2) + '-' + ('0' + enddate.getDate()).slice(-2) + '-' + enddate.getFullYear()]
    }
    if (type == 'PYTD') {
      return ['01-01-' + (enddate.getFullYear() - 1), ('0' + (enddate.getMonth() + 1)).slice(-2) + '-' + ('0' + enddate.getDate()).slice(-2) + '-' + (enddate.getFullYear() - 1)]
    }
    if (type == 'PM') {
      var lastDayOfMonth = new Date(enddate.getFullYear() - 1, enddate.getMonth() + 1, 0);
      return [('0' + (enddate.getMonth() + 1)).slice(-2) + '-01' + '-' + (enddate.getFullYear() - 1), ('0' + (enddate.getMonth() + 1)).slice(-2) + '-' + ('0' + lastDayOfMonth.getDate()).slice(-2) + '-' + (enddate.getFullYear() - 1)]
    }
    if (type == 'LM') {
      if (enddate.getMonth() == 0) {
        return ['12-01-' + (enddate.getFullYear() - 1), '12-31-' + (enddate.getFullYear() - 1)]
      } else {
        var lastDayOfMonth = new Date(enddate.getFullYear(), enddate.getMonth(), 0);
        return [('0' + enddate.getMonth()).slice(-2) + '-01' + '-' + enddate.getFullYear(), ('0' + enddate.getMonth()).slice(-2) + '-' + ('0' + lastDayOfMonth.getDate()).slice(-2) + '-' + enddate.getFullYear()]
      }
    }

    if (type == 'LMGL') {
      let today = new Date()
      let enddate = new Date(today.setDate(today.getDate()));
      if (today.getMonth() == 0) {
        return ['12-01-' + (enddate.getFullYear() - 1), '12-31-' + (enddate.getFullYear() - 1)]
      } else {
        var lastDayOfMonth = new Date(today.getFullYear(), today.getMonth(), 0);
        return [('0' + today.getMonth()).slice(-2) + '-01' + '-' + today.getFullYear(), ('0' + today.getMonth()).slice(-2) + '-' + ('0' + lastDayOfMonth.getDate()).slice(-2) + '-' + today.getFullYear()]
      }
    }
    if (type == 'LY') {
      return ['01-01-' + (enddate.getFullYear() - 1), '12-31-' + (enddate.getFullYear() - 1)]
    }

    if (type == 'TD') {
      return [this.shared.datePipe.transform(new Date(), 'MM-dd-yyyy'), this.shared.datePipe.transform(new Date(), 'MM-dd-yyyy')]
    }
    if (type == 'YD') {
      return [this.shared.datePipe.transform(enddate, 'MM-dd-yyyy'), this.shared.datePipe.transform(enddate, 'MM-dd-yyyy')];
    }
    if (type == 'Overall') {
      return ['', ''];
    }
    if (type == 'All') {
      return ['', ''];
    }
    if (type == '3') {
      let startDate = new Date()
      let FromDate = new Date(startDate.setDate(startDate.getDate() - 3))
      return [this.shared.datePipe.transform(FromDate, 'MM-dd-yyyy'), ('0' + (enddate.getMonth() + 1)).slice(-2) + '-' + ('0' + enddate.getDate()).slice(-2) + '-' + enddate.getFullYear()]
    }
    if (type == '10') {
      let startDate = new Date()
      let FromDate = new Date(startDate.setDate(startDate.getDate() - 10))
      return [this.shared.datePipe.transform(FromDate, 'MM-dd-yyyy'), ('0' + (enddate.getMonth() + 1)).slice(-2) + '-' + ('0' + enddate.getDate()).slice(-2) + '-' + enddate.getFullYear()]
    }
    if (type == '30') {
      let dt = new Date(today.setDate(today.getDate()));
      dt.setDate(dt.getDate() - 30);
      return [this.shared.datePipe.transform(dt, 'MM-dd-yyyy'), ('0' + (enddate.getMonth() + 1)).slice(-2) + '-' + ('0' + enddate.getDate()).slice(-2) + '-' + enddate.getFullYear()];

    }
    if (type == '90') {
      let dt = new Date(today.setDate(today.getDate()));
      dt.setDate(dt.getDate() - 90);
      return [this.shared.datePipe.transform(dt, 'MM-dd-yyyy'), ('0' + (enddate.getMonth() + 1)).slice(-2) + '-' + ('0' + enddate.getDate()).slice(-2) + '-' + enddate.getFullYear()];

    }
    if (type == '60') {
      let dt = new Date(today.setDate(today.getDate()));
      dt.setDate(dt.getDate() - 60);
      return [this.shared.datePipe.transform(dt, 'MM-dd-yyyy'), ('0' + (enddate.getMonth() + 1)).slice(-2) + '-' + ('0' + enddate.getDate()).slice(-2) + '-' + enddate.getFullYear()];

    }
    if (type == '12') {
      const date = new Date(enddate.getFullYear(), enddate.getMonth() + 1, 0);
      let d = new Date(date)
      d.setMonth(d.getMonth() - 12)
      let localstringdate = d.toISOString();
      console.log(date, d, localstringdate, '....................');

      return [this.shared.datePipe.transform(d, 'MM') + '-01-' + this.shared.datePipe.transform(d, 'yyyy'), this.shared.datePipe.transform(date, 'MM-dd-yyyy')];
    }
    if (type == '18') {
      const date = new Date(enddate.getFullYear(), enddate.getMonth() + 1, 0);
      let d = new Date(date)
      d.setMonth(d.getMonth() - 18)
      let localstringdate = d.toISOString();
         return [this.shared.datePipe.transform(d, 'MM') + '-01-' + this.shared.datePipe.transform(d, 'yyyy'), this.shared.datePipe.transform(date, 'MM-dd-yyyy')];

    }
    if (type == '24') {
      const date = new Date(enddate.getFullYear(), enddate.getMonth() + 1, 0);
      let d = new Date(date)
      d.setMonth(d.getMonth() - 24)
      let localstringdate = d.toISOString();
          return [this.shared.datePipe.transform(d, 'MM') + '-01-' + this.shared.datePipe.transform(d, 'yyyy'), this.shared.datePipe.transform(date, 'MM-dd-yyyy')];

    }
    return ['', '']
  }
}
