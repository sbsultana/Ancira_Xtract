import 'zone.js';
import { bootstrapApplication } from '@angular/platform-browser';
import { provideHttpClient } from '@angular/common/http';
import { App } from './app/app';
import { provideRouter, withHashLocation } from '@angular/router';
// import { ToastrModule } from 'ngx-toastr';
import { provideAnimations } from '@angular/platform-browser/animations';
import { routes } from './app/app.routes';
import { importProvidersFrom } from '@angular/core';
import { CurrencyPipe, DatePipe ,DecimalPipe } from '@angular/common';
import { environment } from './environments/environment';

import { APP_BASE_HREF } from '@angular/common';
import { TimeConversionPipe } from './app/Core/Providers/pipes/timeconversion.pipe';
import { Chart, registerables } from 'chart.js';

Chart.register(...registerables);
// const baseTag = document.querySelector('base');
// if (baseTag) {
//   const baseHref = environment.baseHref || '/';
//   baseTag.setAttribute('href', baseHref);
//   console.log(`✅ Base href set to: ${baseHref}`);
// } else {
//   console.warn('⚠️ <base> tag not found in index.html');
// }
bootstrapApplication(App, {
  providers: [

    provideRouter(routes),
    provideHttpClient(),
    provideAnimations(),    
    importProvidersFrom(
      // ToastrModule.forRoot({
      //   positionClass: 'toast-top-right',
      //   timeOut: 5000,
      //   closeButton: true,
      // })
    ),
    DatePipe, CurrencyPipe,
    // { provide: APP_BASE_HREF, useValue: environment.baseHref },
    DecimalPipe,TimeConversionPipe
  ]
});