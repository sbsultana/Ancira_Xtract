import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Repair } from './repair/repair';



@NgModule({
  declarations: [
    Repair
  ],
  imports: [
    CommonModule
  ],
  exports:[Repair]
})
export class RepairModule { }
export {Repair}
