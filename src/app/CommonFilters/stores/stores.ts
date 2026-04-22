import { Component, EventEmitter, HostListener, Input, Output, SimpleChanges } from '@angular/core';
import { common } from '../../common';
import { SharedModule } from '../../Core/Providers/Shared/shared.module';


@Component({
  selector: 'app-stores',
  imports: [SharedModule],
  standalone: true,
  templateUrl: './stores.html',
  styleUrl: './stores.scss'
})
export class Stores {
  @Input() storesFilterData: any = {}
  @Output() StoresData = new EventEmitter();
  @Output() storeLable = new EventEmitter();


  storeIds: any = []
  stores: any = []
  groupName: any = ''
  groupsArray: any = [];
  storename: any = '';
  groupId: any;
  type: any = '';
  others: any = '';
  storedisplayname: any = ''
  storecount: any = null
  Defaultload: any = ''
  otherStoreIds: any = []
  otherstoresArray: any = []
  constructor(private comm: common) { }

  ngOnChanges(changes: SimpleChanges) {
    console.log(changes);

    if (changes['storesFilterData'] && changes['storesFilterData'].currentValue) {
      this.type = changes['storesFilterData'].currentValue.type;
      this.others = changes['storesFilterData'].currentValue.others;
      this.groupsArray = changes['storesFilterData'].currentValue.groupsArray;
      this.stores = changes['storesFilterData'].currentValue.storesArray;
      this.storeIds = changes['storesFilterData'].currentValue.storeids;
      this.groupId = changes['storesFilterData'].currentValue.groupId;
      this.groupName = changes['storesFilterData'].currentValue.groupName;
      this.storename = changes['storesFilterData'].currentValue.storename;
      this.storecount = changes['storesFilterData'].currentValue.storecount;
      this.storedisplayname = changes['storesFilterData'].currentValue.storedisplayname;
      this.Defaultload = changes['storesFilterData'].currentValue.DefaultLoad;
      this.otherstoresArray = changes['storesFilterData'].currentValue.otherStoresArray;
      this.otherStoreIds = changes['storesFilterData'].currentValue.otherStoreIds ? changes['storesFilterData'].currentValue.otherStoreIds : [];

      if (this.storesFilterData.storesArray && this.storesFilterData.storesArray.length > 0) {
        this.getGroupBaseStores(this.groupId.toString(), 'FD')
      }

    }
  }
  ngOnInit() {
    // this.getStoresandGroupsValues()

  }
  get dynamicGroupRowCount(): number {
    return Math.min(this.groupsArray?.length || 0, 9);
  }
  get dynamicStoreCount(): number {
    return Math.min(this.stores?.length || 0, 4);
  }
  get dynamicOtherStoreCount(): number {
    return Math.min(this.otherstoresArray?.length || 0, 4);
  }
  // get dynamicStoreCount(): number {
  //   return Math.max(1, Math.ceil(this.stores.length / 4));
  // }
  // get dynamicOtherStoreCount(): number {
  //   return Math.max(1, Math.ceil(this.otherstoresArray.length / 4));
  // }

  individualgroups(e: any) {
    this.groupId = []
    this.groupId.push(e.sg_id);
    this.groupName = this.groupsArray.filter((val: any) => val.sg_id == this.groupId)[0].sg_name;
    this.getGroupBaseStores(this.groupId.toString(), 'FR')
  }
  getGroupBaseStores(id: any, block: any) {
    this.comm.pageName == 'Used Vehicle Stocking' || this.comm.pageName == 'Fixed Income/Expense w/o Overhead'  || this.comm.pageName == 'Variable Income/Expense w/o Overhead' || this.comm.pageName == 'Enterprise Income/Expense w/o Overhead'  ?
      this.stores = this.comm.groupsandstoresAll.filter((v: any) => v.sg_id == id)[0]?.Stores :
      this.stores = this.comm.groupsandstores.filter((v: any) => v.sg_id == id)[0]?.Stores;
    console.log(this.storeIds, 'Store IDs');
    let singleid = this.storeIds
    this.storeIds = []
    if (block == 'FD') {
      this.type == 'S' ? this.storeIds.push(parseInt(singleid)) : (this.Defaultload == 'E' ? this.storeIds = [] : ( singleid.toString().indexOf(',') > 0 ?  this.storeIds = singleid.map(Number) : this.storeIds.push(parseInt(singleid))  ))
    } else {
      this.type == 'S' ? this.storeIds.push(parseInt(singleid)) : (this.Defaultload == 'E' ? this.storeIds = [] : this.storeIds = this.stores?.map(function (a: any) {
        return a.ID;
      }))
    }

    console.log(this.storeIds, 'After condition');

    if (this.storeIds && this.storeIds.length == 1) {
      this.storename = this.stores.filter((val: any) => val.ID == this.storeIds.toString())[0].storename;
    } else if (this.storeIds && this.storeIds.length == 0 && this.otherStoreIds && this.otherStoreIds.length == 1) {
      this.storename = this.otherstoresArray.filter((val: any) => val.ID == this.otherStoreIds.toString())[0].storename
    } else {
      this.storename = ''
    }
    this.groupName = this.groupsArray.filter((val: any) => val.sg_id == this.groupId)[0].sg_name;
    this.setValues()
  }
  individualStores(e: any) {
    const index = this.storeIds.findIndex((i: any) => i == e.ID);
    if (index >= 0) {
      this.storeIds.splice(index, 1);
    } else {
      this.type == 'S' ? this.storeIds = [] : ''
      this.type == 'S' ? this.storeIds.push(e.ID) : this.storeIds.push(e.ID);
    }
    if (this.storeIds.length == 1) {
      this.storename = this.stores.filter((val: any) => val.ID == this.storeIds.toString())[0].storename
    } else if (this.storeIds && this.storeIds.length == 0 && this.otherStoreIds && this.otherStoreIds.length == 1) {
      this.storename = this.otherstoresArray.filter((val: any) => val.ID == this.otherStoreIds.toString())[0].storename
    } else {
      this.storename = ''
    }
    this.setValues()

  }
  allstores(state: any) {
    if (state == 'N') {
      this.storeIds = [];
    } else if (state == 'Y') {
      this.storeIds = this.stores.map(function (a: any) {
        return a.ID;
      });
      this.groupName = this.groupsArray.filter((val: any) => val.sg_id == this.groupId)[0].sg_name;
    }
    this.setValues()
  }

  individualotherStores(e: any) {
    let index: any = -1
    if (this.otherStoreIds.length > 0) {
      index = this.otherStoreIds?.findIndex((i: any) => i == e.ID);
    }
    if (index >= 0) {
      this.otherStoreIds.splice(index, 1);
    } else {
      this.otherStoreIds.push(e.ID);
    }
    console.log(e, this.otherStoreIds, '................');

    this.setValues()
  }
  allotherstores(state: any) {
    if (state == 'N') {
      this.otherStoreIds = [];
    } else if (state == 'Y') {
      this.otherStoreIds = this.otherstoresArray.map(function (a: any) {
        return a.ID;
      });
    }
    this.setValues()
  }
  setValues() {


    if (this.others == 'Y') {
      let combineStores: any = []
      combineStores = [...this.storeIds, ...this.otherStoreIds]
      if (this.storeIds.length == this.stores.length) {
        this.storecount = null
        this.storedisplayname = this.groupName;
        this.storename = ''

      }
      else if (combineStores.length == 1) {
        this.storecount = null
        this.storeIds.length == 1 ? this.storename = this.stores.filter((val: any) => val.ID == this.storeIds.toString())[0].storename :
          this.storename = this.otherstoresArray.filter((val: any) => val.ID == this.otherStoreIds.toString())[0].storename
        this.storedisplayname = this.storename
      }
      else if (combineStores.length > 1) {
        this.storecount = combineStores.length
        this.storedisplayname = 'Selected'
      }
      else {
        this.storecount = null;
        this.storename = ''
        this.storedisplayname = 'Select'
      }
    } else {
      if (this.storeIds.length == 1) {
        this.storecount = null;
        this.storename = this.stores.filter((val: any) => val.ID == this.storeIds.toString())[0].storename
        this.storedisplayname = this.storename
      }
      else if (this.storeIds.length == this.stores.length) {
        this.storecount = null;
        this.storename = ''
        this.storedisplayname = this.groupName;
      }
      else if (this.storeIds.length > 0 && this.otherStoreIds.length > 0) {
        this.others == 'Y' ? this.storecount = this.storeIds.length + this.otherStoreIds.length : this.storecount = this.storeIds.length
        this.storedisplayname = 'Selected'
      }
      else if (this.storeIds.length > 1) {
        this.others == 'Y' ? this.storecount = this.storeIds.length + this.otherStoreIds.length : this.storecount = this.storeIds.length
        this.storedisplayname = 'Selected'
      }
      else if (this.storeIds.length == 0 && this.otherStoreIds.length == 1) {
        this.storename = this.otherstoresArray.filter((val: any) => val.ID == this.otherStoreIds.toString())[0].storename
        this.storedisplayname = this.storename
        this.storecount = null;
      }
      else {
        this.storename = ''
        this.storedisplayname = 'Select'
      }
    }


    this.storesFilterData.groupsArray = this.groupsArray;
    this.storesFilterData.storesArray = this.stores;
    this.storesFilterData.groupId = this.groupId;
    this.storesFilterData.storeids = this.storeIds;
    this.storesFilterData.storename = this.storename;
    this.storesFilterData.groupName = this.groupName;
    this.storesFilterData.storecount = this.storecount;
    this.storesFilterData.storedisplayname = this.storedisplayname;
    this.storesFilterData.otherStoreIds = this.otherStoreIds;

    this.StoresData.emit(this.storesFilterData);

  }


}