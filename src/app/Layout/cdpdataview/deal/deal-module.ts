import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Deal } from './deal/deal';



@NgModule({
  declarations: [
    Deal
  ],
  imports: [
    CommonModule
  ],
  exports: [Deal]
})
export class DealModule { }
export { Deal }
