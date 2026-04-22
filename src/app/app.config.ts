import { ApplicationConfig, provideBrowserGlobalErrorListeners, provideZonelessChangeDetection } from '@angular/core';
import { provideRouter } from '@angular/router';
import { provideHttpClient, withFetch } from '@angular/common/http';
import { CurrencyPipe, DatePipe } from '@angular/common';
import { routes } from './app.routes';
import { provideClientHydration, withEventReplay } from '@angular/platform-browser';
// import { provideToastr, ToastNoAnimationModule } from 'ngx-toastr';
// import { ToastrModule, ToastrService, TOAST_CONFIG } from 'ngx-toastr';
import { NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { BsDatepickerModule } from 'ngx-bootstrap/datepicker';
import { Ng2SearchPipeModule } from 'ng2-search-filter';


export const appConfig: ApplicationConfig = {
  providers: [
    provideBrowserGlobalErrorListeners(),
    provideZonelessChangeDetection(),
    provideRouter(routes), provideClientHydration(withEventReplay()),
    provideHttpClient(withFetch()), DatePipe,
    BsDatepickerModule,
    Ng2SearchPipeModule,
    //   provideToastr({
    //   closeButton: true,
    //   timeOut: 2000, // 15 seconds
    //   progressBar: true,
    // }),
    //  ToastrService,
    // { provide: TOAST_CONFIG, useValue: { positionClass: 'toast-top-right', timeOut: 5000 } },
    NgbModal, CurrencyPipe,
  ]
};
