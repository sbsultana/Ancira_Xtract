import { Pipe, PipeTransform } from '@angular/core';
import * as moment from 'moment-timezone';
@Pipe({
  name: 'timeconversion'
})
export class TimeConversionPipe implements PipeTransform {

   transform(value: any): string {
     if (!value) return '';
 
     try {
       // Step 1: Parse as UTC (the backend time)
       const utcMoment = moment.utc(value, [
         'MM/DD/YYYY hh:mm A',
         'YYYY-MM-DDTHH:mm:ssZ',
         moment.ISO_8601
       ]);
 
       // Step 2: Automatically detect timezone
       const detectedTz = moment.tz.guess(); // viewer’s timezone (auto)
       // If you want fixed CST for testing, uncomment below:
       // const detectedTz = 'America/Chicago';
 
       // Step 3: Convert UTC → local
       const localTime = utcMoment.clone().tz(detectedTz);
 
       // Step 4: Format as desired
       return localTime.format('MM/DD/YYYY hh:mm A z');
     } catch (err) {
       console.error('Time conversion error:', err);
       return value;
     }
   }

}
