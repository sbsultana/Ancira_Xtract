import { Component, OnInit, signal } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { Header } from './Layout/header/header';
import { CommonModule } from '@angular/common';
import { SidebarService } from '../app/Core/Providers/Shared/sidebar.service (1)';
import { ToastContainer } from './Layout/toast-container/toast-container';
import { UserActivityService } from './Core/Providers/Api/user-activity-service';
import { ExternalmenuComponent } from './Layout/externalmenu/externalmenu';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet, CommonModule, Header, ExternalmenuComponent, ToastContainer],
  templateUrl: './app.html',
  styleUrls: ['./app.scss']
})
export class App implements OnInit {
  isHeaderReady = false;
  isSidebarCollapsed = false;

  protected readonly title = signal('Ancira');

  constructor(private sidebarService: SidebarService,private activityService: UserActivityService) { }
  ngOnInit() {
    this.activityService.startRouteTracking();
    this.sidebarService.isCollapsed$.subscribe((collapsed) => {
      this.isSidebarCollapsed = collapsed;
    });
  }

  onHeaderReady(): void {
    this.isHeaderReady = true;
    console.log('Dashboard received Header ready event');
  }
}
