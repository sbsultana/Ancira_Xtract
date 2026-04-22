import { Injectable } from '@angular/core';
import { BehaviorSubject } from 'rxjs';

@Injectable({ providedIn: 'root' })
export class SidebarService {
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
}
