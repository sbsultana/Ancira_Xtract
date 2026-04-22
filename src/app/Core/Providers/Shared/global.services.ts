import { Injectable } from '@angular/core';


import { BehaviorSubject } from 'rxjs';

@Injectable({
  providedIn: 'root'
})

export class UserNetworkService {
  constructor() {}

  private storeSource = new BehaviorSubject<string | 'all'>('all');
  selectedStore$ = this.storeSource.asObservable();

  setStore(store: any | 'all') {
    this.storeSource.next(store);
  }

  getStore(): any | null {
    return this.storeSource.value;
  }

 async isInsidePrivateNetwork(url: string): Promise<boolean> {
  try {
    const fullUrl = url.startsWith('http') ? url : `https://${url}`;
    await fetch(fullUrl, {
      method: 'HEAD',
      mode: 'no-cors'
    });
    return true;
  } catch (e) {
    return false;
  }
}

  public networkStatus : any = true;

  private user: any | null = null;

  setUser(user: any): void {
    this.user = user;
  }

  getUser(): any | null {
    return this.user;
  }

  clearUser(): void {
    this.user = null;
  }
}
