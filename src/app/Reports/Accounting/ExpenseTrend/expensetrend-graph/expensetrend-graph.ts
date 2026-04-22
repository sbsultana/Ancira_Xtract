import {
  Component,
  Input,
  ViewChild,
  ElementRef,
  AfterViewInit,
  OnDestroy
} from '@angular/core';

import {
  Chart,
  registerables
} from 'chart.js';
Chart.register(...registerables);

import { NgbActiveModal } from '@ng-bootstrap/ng-bootstrap';
import { formatNumber } from '@angular/common';

@Component({
  selector: 'app-expensetrend-graph',
  standalone: true,
  templateUrl: './expensetrend-graph.html',
  styleUrls: ['./expensetrend-graph.scss'],
})
export class ExpensetrendGraph implements AfterViewInit, OnDestroy {

  @Input() ETgraphdetails: any;
  @ViewChild('mychart') mychart!: ElementRef<HTMLCanvasElement>;

  monthvalues: number[] = [];
  monthkeys: string[] = [];
  rV: number[] = [];

  private chart!: Chart;

  constructor(public modal: NgbActiveModal) { }

  ngAfterViewInit(): void {
    this.prepareData();

    const points = this.monthvalues.map((v, i) => [i, v]);
    this.lineFit(points);

    this.chart = new Chart(this.mychart.nativeElement, {
      type: 'line',
      data: {
        labels: this.monthkeys,
        datasets: [
          {
            label: this.ETgraphdetails.NAME,
            borderColor: '#00c2ff',        
            backgroundColor: 'rgba(0, 194, 255, 0.15)',
            fill: true,
            data: this.monthvalues,
            tension: 0.35,
            pointBackgroundColor: '#ffffff',
            pointBorderColor: '#00c2ff',
            pointRadius: 4,
            borderWidth: 2,
          },
          {
            label: 'Trend Line',
            data: this.rV,
            borderColor: '#ffab00',       
            borderWidth: 2,
            borderDash: [6, 6],            
            pointRadius: 0,
            fill: false,
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: {
            ticks: { color: '#ffffff' },
            grid: { display: false }
          },
          y: {
            ticks: {
              color: '#ffffff',
              callback: (value) => {
                return this.ETgraphdetails.ValueFormat === 'Percentage'
                  ? value + '%'
                  : formatNumber(Number(value), 'en-US', '1.0-0');
              }
            },
            grid: { color: '#ffffff33' }
          }
        },
        plugins: {
          legend: {
            labels: { color: '#ffffff' }
          },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const val = Number(ctx.raw);
                return this.ETgraphdetails.ValueFormat === 'Percentage'
                  ? `${val}%`
                  : formatNumber(val, 'en-US', '1.0-0');
              }
            }
          }
        }
      }
    });
  }

  prepareData() {
    const obj = { ...this.ETgraphdetails.ITEM };

    const removeKeys = ['SNO', 'PARENTLABLECODE', 'PYTD', 'YTD', 'AVG', 'ISHEAD_TOTAL', 'LABLECODE', 'DISPLAYHEAD_FLAG', 'DISPLAY_LABLE', 'DISPLAY_PARENTLABLE'];

    this.monthkeys = Object.keys(obj).filter(k => !removeKeys.includes(k));
    this.monthvalues = this.monthkeys.map(k => Number(obj[k]) || 0);

    // Format display labels
    this.monthkeys = this.monthkeys.map(k =>
      k.toUpperCase().replace('-', ' ')
    );
  }


  lineFit(points: any[]) {
    const N = points.length;
    if (N < 2) return;

    let sumX = 0, sumY = 0, sumXx = 0, sumXy = 0;
    points.forEach(([x, y]) => {
      sumX += x;
      sumY += y;
      sumXx += x * x;
      sumXy += x * y;
    });

    const slope = (N * sumXy - sumX * sumY) / (N * sumXx - sumX * sumX);
    const intercept = (sumY - slope * sumX) / N;

    this.rV = points.map(([x]) => slope * x + intercept);
  }

  close() {
    this.modal.close();
  }

  ngOnDestroy(): void {
    if (this.chart) this.chart.destroy();
  }
}
