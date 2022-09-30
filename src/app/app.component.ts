import { Component } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { AuthenticationResult } from '@azure/msal-browser';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'OfficeJsDemo';

  loginLinkText:string="Logout";
  constructor(private msalService:MsalService)
  {
    if(!this.isUserLoggedIn())
    {
      this.loginLinkText="Login";
    }
  }
  login()
  {
    this.msalService.loginPopup().subscribe((response:AuthenticationResult)=>{
        this.msalService.instance.setActiveAccount(response.account);
        //console.log(JSON.stringify(response));
        localStorage.setItem('accessToken', response.accessToken);
        localStorage.setItem('idToken', response.idToken);
    })
  }
  logout()
  {
    this.msalService.logout();
  }
  isUserLoggedIn():boolean
  {
    if(this.msalService.instance.getActiveAccount()!=null)
    {
      return true;
    }
    return false;
  }
}
