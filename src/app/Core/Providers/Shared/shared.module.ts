import { NgModule } from '@angular/core';
import { CommonModule, DatePipe, DecimalPipe } from '@angular/common';
import { NgbModal, NgbActiveModal,NgbTooltipModule, NgbModule } from '@ng-bootstrap/ng-bootstrap';
import { NgxSpinnerModule } from 'ngx-spinner';
// import { ToastrModule } from 'ngx-toastr';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { FilterPipe, FilterPipeModule } from 'ngx-filter-pipe';
import { TimeConversionPipe } from '../pipes/timeconversion.pipe';


@NgModule({
  imports: [
    CommonModule,
    NgxSpinnerModule,
    FilterPipeModule,TimeConversionPipe,
    // ToastrModule.forRoot(),
  ],
  exports: [
    CommonModule,
    NgxSpinnerModule,
    // ToastrModule,
    NgbModule,
    ReactiveFormsModule,
    FormsModule,
    NgbTooltipModule,
    FilterPipeModule,TimeConversionPipe
  ],
  providers: [
    DatePipe,
    NgbModal,
    NgbActiveModal,
    FilterPipe,
    DecimalPipe
  ]
})
export class SharedModule {}