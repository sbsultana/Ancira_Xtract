import { Injectable } from '@angular/core';
import { NgxSpinnerService } from 'ngx-spinner';
// import { ToastrService } from 'ngx-toastr';
import { DatePipe } from '@angular/common';
import { Api } from '../Api/api';

import { Title } from '@angular/platform-browser';
import { environment } from '../../../../environments/environment'
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import * as FileSaver from 'file-saver';
import { Workbook } from 'exceljs';
import { common } from '../../../common';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { BehaviorSubject } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class Sharedservice {
  constructor(
    public api: Api,
    public spinner: NgxSpinnerService,
    // public toastr: ToastrService,
    public datePipe: DatePipe,
    public titleService: Title,
    public common: common,
    // public toast: ToastrService,
    public ngbmodal: NgbModal,
 
    
  ) { }


   private collapsed = new BehaviorSubject<boolean>(this.loadInitialState());
  isCollapsed$ = this.collapsed.asObservable();
 
  toggle() {
    const newState = !this.collapsed.value;
    this.collapsed.next(newState);
    localStorage.setItem('sidebarCollapsed', JSON.stringify(newState));
  }
 
  setCollapsed(state: boolean) {
    this.collapsed.next(state);
    localStorage.setItem('sidebarCollapsed', JSON.stringify(state));
  }
 
  getCollapsed(): boolean {
    return this.collapsed.value;
  }
 
  private loadInitialState(): boolean {
    const saved = localStorage.getItem('sidebarCollapsed');
    return saved ? JSON.parse(saved) : false;
  }

  setTitle(title: string) {
    this.titleService.setTitle(title);
  }
  // showSuccess(message: string, title?: string) {
  //   this.toastr.success(message, title);
  // }

  // showError(message: string, title?: string) {
  //   this.toastr.error(message, title);
  // }

  // showInfo(message: string, title?: string) {
  //   this.toastr.info(message, title);
  // }

  // showWarning(message: string, title?: string) {
  //   this.toastr.warning(message, title);
  // }
  exportToPDF(elementId: string, fileName: string) {
    const element = document.getElementById(elementId);
    if (!element) return;
    html2canvas(element).then(canvas => {
      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF();
      // pdf.addImage(imgData, 'PNG', 0, 0);
      pdf.save(fileName);
    });
  }
  exportToExcel(workbook: Workbook, fileName: string) {
    workbook.xlsx.writeBuffer().then(buffer => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      FileSaver.saveAs(blob, fileName);
    });
  }

  getEnviUrl() {
    return environment.apiUrl;
  }

  getWorkbook(){
    return new Workbook()
  }
}
