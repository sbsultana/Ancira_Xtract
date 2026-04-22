import { Component, Input } from '@angular/core';
import { NgbActiveModal } from '@ng-bootstrap/ng-bootstrap';
import { DecimalPipe } from '@angular/common';
import { Api } from '../../api';

@Component({
  selector: 'app-repair',
  standalone: false,
  templateUrl: './repair.html',
  styleUrl: './repair.scss',
  providers: [DecimalPipe]
})
export class Repair {
  @Input() data: any;
  rocompleteDetails: any;
  rosummary: any;
  customerpay: any;
  internalpay: any;
  roDetails: any = [];
  rotimeDetails: any;
  storeDetails: any;
  userDetails: any;
  vehicleDetails: any;
  options: any = [];
  total: any = [];
  isLoading!: boolean;
  showData!: boolean;
  constructor(private activeModal: NgbActiveModal, private http: Api, private decimalPipe: DecimalPipe) { }
  ngOnInit(): void {
    this.getServiceData()
  }
  formatNumaricNumber(number: number): string | null {
    return this.decimalPipe.transform(number, '1.2-2');
  }

  closeModal() {
    this.activeModal.close('Modal Closed');
  }
  getServiceData() {
    let obj = {
      "ronumber": this.data.ro ? this.data.ro : "",
      "vin": this.data.vin ? this.data.vin : "",
      "storeid": this.data.storeid ? this.data.storeid : "",
      "vehicleid": this.data.vehicleid ? this.data.vehicleid : "",
      "custno": this.data.custno ? this.data.custno : '',

    }
    this.isLoading = true;
    this.http.post('cc/basic/getroview', obj).subscribe((res: any) => {
      if (JSON.parse(res.response[0].userdetails) == null) {
        this.showData = false;
      }
      else {
        this.showData = true;
      }
      if (res.response[0]?.rocompletedetails)
        this.rocompleteDetails = JSON.parse(res.response[0]?.rocompletedetails);
      if (res.response[0]?.rosummary)
        this.rosummary = JSON.parse(res.response[0]?.rosummary);
      if (res.response[0]?.rodetails)
        this.roDetails = JSON.parse(res.response[0]?.rodetails);
      if (res.response[0]?.rotimedetails)
        this.rotimeDetails = JSON.parse(res.response[0]?.rotimedetails);
      if (res.response[0]?.storedetails)
        this.storeDetails = JSON.parse(res.response[0]?.storedetails);      
      if (res.response[0]?.userdetails)
        this.userDetails = JSON.parse(res.response[0]?.userdetails);
      if (res.response[0]?.vehicledetails)
        this.vehicleDetails = JSON.parse(res.response[0]?.vehicledetails);
      if (res.response[0]?.customerpay)
        this.customerpay = JSON.parse(res.response[0]?.customerpay);
      if (res.response[0]?.internalpay)
        this.internalpay = JSON.parse(res.response[0]?.internalpay);
      this.isLoading = false;
      console.log("rocompleteDetails", this.rocompleteDetails);
      console.log("roDetails", this.roDetails);
      console.log("rotimeDetails", this.rotimeDetails);
      console.log("storeDetails", this.storeDetails);
      console.log("userDetails", this.userDetails);
      console.log("vehicleDetails", this.vehicleDetails);
      console.log("customerpay", this.customerpay);
      let options;
      options = this.roDetails.filter((item:any) => item.datatitle == "Options:");
      if (options[0]?.datavalue) 
      this.options = JSON.parse(options[0]?.datavalue)
      this.roDetails = this.roDetails.filter((item:any) => item.datatitle != "Options:");
      console.log(this.options);
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
    if (!phoneNumber && phoneNumber != '') return '';
    const cleaned = ('' + phoneNumber).replace(/\D/g, '');
    const match = cleaned.match(/^(\d{3})(\d{3})(\d{4})$/);
    if (match) {
      return `(${match[1]}) ${match[2]}-${match[3]}`;
    }
    return phoneNumber;
  }
}
