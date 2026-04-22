import { Pipe, PipeTransform } from '@angular/core';
 
@Pipe({
  name: 'ValueExistence'
})
export class ValueExistencePipe implements PipeTransform {
 
  transform(value: number | null | undefined | string): string | number {
    return value === 0 || value == null || value == '' ? '-' : value
  }
}