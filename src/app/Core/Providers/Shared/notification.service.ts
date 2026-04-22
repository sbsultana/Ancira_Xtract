import { Injectable } from '@angular/core';
import { BehaviorSubject, Subject } from 'rxjs';
import {  Api} from '../Api/api'; // your existing API service


@Injectable({
  providedIn: 'root',
})
export class NotificationService {
  private unreadCount = new BehaviorSubject<number>(0);
  unreadCount$ = this.unreadCount.asObservable();
  public nCount = false;

  constructor(private api: Api, public apiCall : Api) {}

  fetchUnreadCount() {
    
    setTimeout(async () => {
    var fcmTokenSession : any = localStorage.getItem('fcmToken');
    fcmTokenSession = JSON.parse(fcmTokenSession); 
    
    var endPoint = 'fcm/getunreadcount?user_to_aou_id='+fcmTokenSession?.user_aou_id+'&group_code='+fcmTokenSession?.group_code
    this.apiCall.getNotificationsUnreadedCount(endPoint).subscribe({
        next: (data: { status: number; response: { unread_count: number; }; }) => {
        console.log('Notifications :', data);
        if(data.status == 200) {
            var counts = data?.response?.unread_count || 0;
            this.unreadCount.next(counts);
            this.nCount = counts > 0 ? true : false;
            console.log('Get Notification Count : ', data);
        }
        },
        error: (err: any) => {
        console.error('Failed to get FCM notificaitons! :', err);
        }
    });
    }, 3000);
  }
  

  //Call this after sending a new notification to refresh
  refresh() {
    this.fetchUnreadCount();
  }

  //Check status
  checkStatus(){
    this.nCount = true;
  }

  // To manually reset the count, e.g., after viewing
  clear() {
    this.unreadCount.next(0);
  }


  // Notifications div
  private notificationTrigger = new Subject<void>();

  notificationTrigger$ = this.notificationTrigger.asObservable();

  triggerFetch() {
    //alert('I am treggared..!');
    this.notificationTrigger.next();
  }
}
