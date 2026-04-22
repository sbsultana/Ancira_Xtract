// src/app/core/services/toast.service.ts
import { Injectable, signal } from '@angular/core';

// export interface ToastMessage {
//   id: number;
//   text: string;
//   variant: 'success' | 'danger' | 'info';
// }


export interface ToastMessage {
  id: number;
  title?: string;
  text: string;
  icon?: string; // Can be emoji or image URL
  iconType?: 'emoji' | 'img'; // <- NEW
  variant: 'success' | 'danger' | 'warning' | 'info';
}

export interface ToastData {
  type: 'success' | 'warning' | 'error' | 'info';
  title: string;
  message: string;
}

@Injectable({ providedIn: 'root' })
export class ToastService {
  private counter = 0;
  private readonly _queue = signal<ToastMessage[]>([]);
  readonly queue = this._queue.asReadonly();
  public toasts: ToastData[] = [];

  private defaultIcons: Record<string, string> = {
    success: 'images/AlertIcons/info-icon.svg',
    warning: 'images/AlertIcons/warning-icon.svg',
    danger: 'images/AlertIcons/danger-icon.svg',
    info: 'images/AlertIcons/info-icon.svg',
  };


  show(
    text: string,
    variant: ToastMessage['variant'] = 'info',
    title?: string,
    icon?: string,
    iconType: 'emoji' | 'img' = 'img'
  ) {
    const toast: ToastMessage = {
      id: this.counter++,
      text,
      title,
      variant,
      icon: icon || this.defaultIcons[variant], // ✅ auto select based on variant
      iconType,
    };

    this._queue.update(toasts => [...toasts, toast]);

    setTimeout(() => this.dismiss(toast.id), 8000);
  }

  dismiss(id: number) {
    console.log('Dismiss called for toast ID:', id);
    this._queue.update(toasts => toasts.filter(t => t.id !== id));
  }

  // showc(toast: ToastData) {
  //   this.toasts.push(toast);
  //   setTimeout(() => this.remove(toast), 4000); // auto dismiss in 4s
  // }

  remove(toast: ToastData) {
    this.toasts = this.toasts.filter(t => t !== toast);
  }

  success(message: string) {
    this.show(message, 'success');
  }
}
