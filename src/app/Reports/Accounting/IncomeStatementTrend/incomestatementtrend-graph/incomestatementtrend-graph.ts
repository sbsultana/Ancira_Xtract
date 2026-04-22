import {
  Component,
  OnInit,
  ViewChild,
  Input,
  Inject,
  ElementRef,
  AfterViewInit,
  OnDestroy,
} from '@angular/core';

import {
  Chart,
  ChartConfiguration,
  registerables
} from 'chart.js';
Chart.register(...registerables);

import { NgbActiveModal } from '@ng-bootstrap/ng-bootstrap';
import { LOCALE_ID } from '@angular/core';
import numeral from 'numeral';

@Component({
  selector: 'app-incomestatementtrend-graph',
  standalone: true,
  templateUrl: './incomestatementtrend-graph.html',
  styleUrls: ['./incomestatementtrend-graph.scss'],
})
export class IncomestatementtrendGraph
  implements OnInit, AfterViewInit, OnDestroy {
  @Input() INSTGRAPHdetails: any;

  @ViewChild('mychart', { static: true }) mychart!: ElementRef<HTMLCanvasElement>;

  monthvalues: number[] = [];
  monthkeys: string[] = [];
  itemdata: any;
  rV: number[] = [];
  myChart!: Chart;

  constructor(
    private ngbmodalActive: NgbActiveModal,
    @Inject(LOCALE_ID) public locale: string,
    private element: ElementRef
  ) { }

  ngOnInit(): void {
    this.itemdata = this.INSTGRAPHdetails.ITEM;
    this.monthvalues = Object.values(this.itemdata);
    this.monthkeys = Object.keys(this.itemdata);

    const keysToRemove = ['LABLEVAL', 'PYTD', 'LABLE', 'LYTD', 'YTD', 'CommentsStatus', 'Pace'];
    keysToRemove.forEach((k) => {
      const idx = this.monthkeys.indexOf(k);
      if (idx > -1) {
        this.monthkeys.splice(idx, 1);
        this.monthvalues.splice(idx, 1);
      }
    });

    if (this.INSTGRAPHdetails.REF === 'Month') {
      const idx = this.monthkeys.indexOf('LABLE');
      if (idx > -1) {
        this.monthkeys.splice(idx, 1);
        this.monthvalues.splice(idx, 1);
      }
    }

    this.monthvalues.reverse();
    this.monthkeys = this.monthkeys.reverse().map((e: any) =>
      e.toString().replace('-', ' ').toUpperCase()
    );
  }

  ngAfterViewInit() {
    const points = this.monthvalues.map((v, i) => [i, v]);
    this.lineFit(points);

    const ctx = this.mychart.nativeElement.getContext('2d');
    if (!ctx) return;

    const chartConfig: ChartConfiguration = {
      type: 'line',
      data: {
        labels: this.monthkeys,
        datasets: [
          {
            label: this.INSTGRAPHdetails.NAME,
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
          },
        ],

      },
      options: {
        responsive: true,
        plugins: {
          legend: {
            labels: {
              color: 'white',
              usePointStyle: true,
            },
          },
          tooltip: {
            callbacks: {
              label: (ctx: any) => {
                const value = ctx.parsed.y;
                const label = ctx.dataset.label || '';

                if (label.includes('Unit')) {
                  return `${numeral(value).format('0,0')}`;
                }

                if (label.includes('%')) {
                  return `${value}%`;
                }

                return numeral(value).format('$0,0[.]00');
              }
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: {
              color: 'white',
              callback: (value: any) => numeral(value).format('0a'),
            },
          },
          x: {
            ticks: {
              color: 'white',
              font: { size: 12 },
            },
          },
        },
      },
    };

    this.myChart = new Chart(ctx, chartConfig);
  }

  lineFit(points: number[][]) {
    const { slope, intercept } = this.slopeAndIntercept(points);
    this.rV = points.map((p) => slope * p[0] + intercept);
  }

  private slopeAndIntercept(points: number[][]): { slope: number; intercept: number } {
    const N = points.length;
    if (N < 2) return { slope: 0, intercept: points[0][1] };

    let sumX = 0, sumY = 0, sumXx = 0, sumXy = 0;
    points.forEach(([x, y]) => {
      sumX += x;
      sumY += y;
      sumXx += x * x;
      sumXy += x * y;
    });

    const slope = (N * sumXy - sumX * sumY) / (N * sumXx - sumX * sumX);
    const intercept = (sumY - slope * sumX) / N;

    return { slope, intercept };
  }

  close() {
    this.ngbmodalActive.close();
  }

  ngOnDestroy() {
    this.myChart?.destroy();
  }
}
