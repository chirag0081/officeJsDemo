import { Component, Inject, OnInit } from '@angular/core';
import {
  MsalBroadcastService,
  MsalGuardConfiguration,
  MsalService,
  MSAL_GUARD_CONFIG,
} from '@azure/msal-angular';
import { AuthenticationResult, RedirectRequest } from '@azure/msal-browser';
import { ActivatedRoute, Router } from '@angular/router';

declare const Office: any;
@Component({
  selector: 'app-login',
  templateUrl: './login.component.html',
  styleUrls: ['./login.component.scss'],
})
export class LoginComponent implements OnInit {
  title = 'msal-angular-tutorial';
  constructor(
    private msalService: MsalService,
    @Inject(MSAL_GUARD_CONFIG) private msalGuardConfig: MsalGuardConfiguration,
    private msalBroadcastService: MsalBroadcastService,
    private activateRoute: ActivatedRoute,
    private router: Router
  ) {}

  ngOnInit(): void {
    if (!this.isUserLoggedIn()) {
      this.msalService.loginRedirect();
    } else {
      Office.context.ui.messageParent(true.toString());
    }
    //this.msalService.instance.handleRedirectPromise("true").then(()=> console.log('Success'));
  }

  isUserLoggedIn(): boolean {
    return this.msalService.instance.getActiveAccount() !== null;
  }
}
