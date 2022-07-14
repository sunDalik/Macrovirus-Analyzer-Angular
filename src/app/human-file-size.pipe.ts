import {Pipe, PipeTransform} from '@angular/core';

@Pipe({
  name: 'humanFileSize'
})
export class HumanFileSizePipe implements PipeTransform {
  // value is in bytes
  transform(value: number): string {
    const byteUnits = ['KB', 'MB', 'GB', 'TB'];
    let i = -1;
    while (value > 1000) {
      value = value / 1000;
      i++;
      if (i >= byteUnits.length - 1) break;
    }

    return Math.max(value, 0.1).toFixed(1) + " " + byteUnits[i];
  }
}
