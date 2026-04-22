import { AfterViewInit, Component, ElementRef, Input, ViewChild } from '@angular/core';

import { Rive, StateMachineInput, Rive as RiveInstance } from '@rive-app/canvas';
@Component({
  selector: 'app-riv-icon',
  standalone: true,
  imports: [],
  templateUrl: './riv-icon.html',
  styleUrl: './riv-icon.scss'
})
export class RivIcon implements AfterViewInit {
  @Input() src!: string; // path to .riv file
  @Input() stateMachine: string = 'State Machine 1'; // default state machine
  @Input() triggerName: string = 'hover'; // 👈 trigger input name
  @Input() width: number = 24;
  @Input() height: number = 24;

  @ViewChild('canvasEl', { static: true }) canvasRef!: ElementRef;

  private riveInstance: any;

  ngAfterViewInit() {
    this.init();
  }

  play() {
    
    if(!this.riveInstance) this.init();
    this.riveInstance?.play();
  }

  pause() {
    
    setTimeout(() => {
      this.riveInstance = null;
    }, 500);
  }

  init(){
    
    this.riveInstance = new Rive({
      src: this.src,
      canvas: this.canvasRef.nativeElement,
      autoplay: false,
      stateMachines: [this.stateMachine],
      onLoad: () => {
        
        this.canvasRef.nativeElement.style.width = '26px';
        this.canvasRef.nativeElement.style.height = '26px';
        this.riveInstance.resizeDrawingSurfaceToCanvas();
        const names = this.riveInstance.stateMachineNames;
      },
    });
  }
}