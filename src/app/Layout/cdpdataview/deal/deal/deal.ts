import { Component, ElementRef, EventEmitter, Input, Output, ViewChild } from '@angular/core';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { Api } from '../../api';

@Component({
  selector: 'app-deal',
  standalone: false,
  templateUrl: './deal.html',
  styleUrl: './deal.scss'
})
export class Deal {
  @Input() data: any;
  amountDetails: any;
  feeDetails: any;
  grossDetails: any;
  rebateDetails: any;
  sales: any;
  summary: any;
  userDetails: any;
  isLoading: boolean = false;
  storedata: any;
  @Output() onClose1 = new EventEmitter()
  trade1: any = [];
  trade2: any = [];
  @ViewChild('recap1') recapElement: ElementRef | undefined;
  salesMembers: any = [];
  constructor(private http: Api, private ngbModal: NgbModal, private activeModal: NgbActiveModal) {

  }
  ngOnInit() {
    console.log(this.data);
    this.getSalesData()

  }
  getSalesData() {
    let obj = {
      "dealnumber": this.data.dealno ? this.data.dealno : '',
      "vin": this.data.vin ? this.data.vin : '',
      "storeid": this.data.storeid ? this.data.storeid : '',
      "stock": this.data.stock ? this.data.stock : '',
      "custno": this.data.custno ? this.data.custno : '',
    }
    this.isLoading = true;
    this.http.post('cc/basic/getdealview', obj).subscribe((res: any) => {
      this.isLoading = false;
      console.log(res);
      if (res.response[0]?.amountdetails)
        this.amountDetails = JSON.parse(res.response[0].amountdetails);
      if (res.response[0]?.feedetails)
        this.feeDetails = JSON.parse(res.response[0].feedetails);
      if (res.response[0]?.storedata)
        this.grossDetails = JSON.parse(res.response[0].grossdetails);
      if (res.response[0]?.rebatedetails)
        this.rebateDetails = JSON.parse(res.response[0].rebatedetails);
      if (res.response[0]?.sales)
        this.sales = JSON.parse(res.response[0].sales);
      if (res.response[0]?.summary)
        this.summary = JSON.parse(res.response[0].summary);
      if (res.response[0]?.userdetails)
        this.userDetails = JSON.parse(res.response[0].userdetails);
      if (res.response[0]?.storedata)
        this.storedata = JSON.parse(res.response[0]?.storedata);
      if (res.response[0]?.trade1)
        this.trade1 = JSON.parse(res.response[0]?.trade1);
      if (res.response[0]?.trade2)
        this.trade2 = JSON.parse(res.response[0]?.trade2);
      if (res.response[0]?.salesmembers)
        this.salesMembers = JSON.parse(res.response[0]?.salesmembers);
      console.log("amountDetails", this.amountDetails);
      console.log("feeDetails", this.feeDetails);
      console.log("grossDetails", this.grossDetails);
      console.log("reDateDetails", this.rebateDetails);
      console.log("sales", this.sales);
      console.log("summary", this.summary);
      console.log("userDetails", this.userDetails);
      console.log("store", this.storedata);
      console.log("SalesMembers", this.salesMembers);

    })
  }
  checkSymbols(value: any): boolean {
    if (value !== '') {
      let dataval: any = value?.toString();
      return (dataval?.includes('-') && dataval?.includes('$'));
    }
    else {
      return false;
    }
  }
  formatPhoneNumber(phoneNumber: string): string {
    if (!phoneNumber) return '';
    const cleaned = ('' + phoneNumber).replace(/\D/g, '');
    const match = cleaned.match(/^(\d{3})(\d{3})(\d{4})$/);
    if (match) {
      return `(${match[1]}) ${match[2]}-${match[3]}`;
    }
    return phoneNumber;
  }
  print() {
    const totalRow = document.querySelector('.recap1');
    if (totalRow) {
      console.log(totalRow);
      const printWindow: any = window.open('', '', 'height=500, width=500');
      // printWindow.document.write('<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8" /><style> @media print {@page {margin: 0;}}body {margin: 0;}header, .title, .date, footer {display: none !important;}.some-other-element { display: none !important;}.content {width: 100%;padding: 20px;}</style></head><body><h2>Home page</h2></body></html>')
      // printWindow.document.write('<html><head><style type="text/css" media="print">@print {@page :footer {display: none}@page :header {display: none} }</style></head><body>');
      printWindow.document.write(`<html><head><title>Deal</title> <style>
      @font-face {font-family: 'Roboto';
      src: url('assets/fonts/Roboto-Regular.ttf') ;
      font-weight: 700;
    }
     @font-face {
        font-family: 'FaktPro-Normal';
        src:url('${this.http.fontUrl}FaktPro-Normal.otf') format("opentype");
        font-weight: 400;
      }
      @font-face {
        font-family: 'FaktPro-Bold';
        src:url('${this.http.fontUrl}FaktPro-Bold.otf') format("opentype");
        font-weight: 500;
      }
      @font-face {
        font-family: 'FaktPro-Light';
        src:url('${this.http.fontUrl}FaktPro-Light.otf') format("opentype");
        font-weight: 700;
      }
        @media print {
          @font-face {
            font-family: 'FaktPro-Medium';
            src: url("${this.http.fontUrl}FaktPro-Medium.otf") format("opentype");
            font-weight: 300;
          }
          @font-face {
            font-family: 'FaktPro-Light';
            src:url('${this.http.fontUrl}FaktPro-Light.otf') format("opentype");
            font-weight: 400;
          }
          @font-face {
            font-family: 'FaktPro-Bold';
            src:url('${this.http.fontUrl}FaktPro-Bold.otf') format("opentype");
            font-weight: 500;
          }@font-face {
            font-family: 'FaktPro-Normal';
            src:url('${this.http.fontUrl}FaktPro-Normal.otf') format("opentype");
            font-weight: 700;}body {font-family: 'FaktPro-Normal';}}
body {font-family: 'FaktPro-Normal';}</style></head > <body>`);
      // printWindow.document.write('<style>@print {@page :footer { display: none }@page :header {display: none}}</style>');
      printWindow.document.write(totalRow.innerHTML); // Copy the content you want to print
      printWindow.document.write('</body></html>');
      printWindow.focus();
      setTimeout(() => {
        printWindow.print();

      }, 2000);
    }
  }
  closeModal() {
    this.activeModal.close('Modal Closed');
    this.onClose1.emit();
  }
}
