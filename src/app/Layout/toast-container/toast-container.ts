import { ChangeDetectionStrategy, Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ToastService } from '../../Core/Providers/Shared/toast.service';
@Component({
  selector: 'app-toast-container',
  standalone: true, 
  imports: [CommonModule],
  templateUrl: './toast-container.html',
  styleUrl: './toast-container.scss'
})
export class ToastContainer {
  constructor(public toastSvc: ToastService) { }

  dismiss(id: number) {
    this.toastSvc.dismiss(id);
  }

  getIcon(type: string): string {
    switch (type) {
      case 'success': return '✔️';
      case 'warning': return '⚠️';
      case 'danger': return '❌';
      case 'info': return 'ℹ️';
      default: return '🔔';
    }
  }

  getDefaultIcon(type: string): string {
    switch (type) {
      case 'success': return 'images/AlertIcons/info-icon.svg';
      case 'warning': return 'images/AlertIcons/warning-icon.svg';
      case 'danger': return 'images/AlertIcons/danger-icon.svg';
      case 'info': return 'images/AlertIcons/info-icon.svg';
      default: return 'images/AlertIcons/info-icon.svg';
    }
  }
}
