import { MsalService } from '@azure/msal-angular';
import { Injectable } from '@angular/core';
import { ActivatedRouteSnapshot, CanActivate, RouterStateSnapshot, UrlTree } from '@angular/router';
import { Observable } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class UserauthguardGuard implements CanActivate {
  
  constructor(private msalService:MsalService)
  {

  }
  canActivate(
    
    route: ActivatedRouteSnapshot,
    state: RouterStateSnapshot): Observable<boolean | UrlTree> | Promise<boolean | UrlTree> | boolean | UrlTree {
      
      if(this.msalService.instance.getActiveAccount()!=null)
      {
        return true;
      }
      else
      {
        return false;
      }
     
  }
  
}