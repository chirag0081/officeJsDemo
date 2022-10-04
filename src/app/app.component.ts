import { logging } from 'protractor';
import { Component, Inject, OnInit } from '@angular/core';
import {
  MsalBroadcastService,
  MsalGuardConfiguration,
  MsalService,
  MSAL_GUARD_CONFIG,
} from '@azure/msal-angular';
import {
  AuthenticationResult,
  InteractionStatus,
  PopupRequest,
  RedirectRequest,
} from '@azure/msal-browser';
import { Subject, Subscription } from 'rxjs';
import { filter, takeUntil } from 'rxjs/operators';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent implements OnInit {
  title = 'OfficeJsDemo';
  loggingSubscription: Subscription;
  loginLinkText: string = 'Logout';
  loginDisplay: boolean = false;
  private readonly onDestroy$ = new Subject<void>();

  constructor(
    private msalService: MsalService,
    @Inject(MSAL_GUARD_CONFIG) private msalGuardConfig: MsalGuardConfiguration,
    private msalBroadcastService: MsalBroadcastService
  ) {}
  ngOnInit(): void {
    this.msalBroadcastService.inProgress$
      .pipe(
        filter(
          (status: InteractionStatus) => status === InteractionStatus.None
        ),
        takeUntil(this.onDestroy$)
      )
      .subscribe(() => {
        this.setLoginDisplay();
        this.checkAndSetActiveAccount();
      });

    if (!this.isUserLoggedIn()) {
      this.loginLinkText = 'Login';
      this.loggingSubscription?.unsubscribe();
    }
  }

  setLoginDisplay() {
    this.loginDisplay = this.msalService.instance.getAllAccounts().length > 0;
  }

  checkAndSetActiveAccount() {
    /**
     * If no active account set but there are accounts signed in, sets first account to active account
    
   
    * To use active account set here, subscribe to inProgress$ first in your component
     * Note: Basic usage demonstrated. Your app may require more complicated account selection logic
     */
    let activeAccount = this.msalService.instance.getActiveAccount();

    if (
      !activeAccount &&
      this.msalService.instance.getAllAccounts().length > 0
    ) {
      let accounts = this.msalService.instance.getAllAccounts();
      this.msalService.instance.setActiveAccount(accounts[0]);
    }
  }

  loginRedirect() {
    if (this.msalGuardConfig.authRequest) {
      this.msalService.loginRedirect({
        ...this.msalGuardConfig.authRequest,
      } as RedirectRequest);
    } else {
      this.msalService.loginRedirect();
    }
  }

  loginPopup() {
    if (this.msalGuardConfig.authRequest) {
      this.msalService
        .loginPopup({ ...this.msalGuardConfig.authRequest } as PopupRequest)
        .subscribe((response: AuthenticationResult) => {
          this.msalService.instance.setActiveAccount(response.account);
          localStorage.setItem('accessToken', response.accessToken);
          localStorage.setItem('idToken', response.idToken);
        });
    } else {
      this.msalService
        .loginPopup()
        .subscribe((response: AuthenticationResult) => {
          this.msalService.instance.setActiveAccount(response.account);
          localStorage.setItem('accessToken', response.accessToken);
          localStorage.setItem('idToken', response.idToken);
        });
    }
  }

  ngOnDestroy(): void {
    this.onDestroy$.next(undefined);
    this.onDestroy$.complete();
  }

  login(): void {
    this.loggingSubscription = this.msalService
      .loginPopup()
      .subscribe((response: AuthenticationResult) => {
        this.msalService.instance.setActiveAccount(response.account);
        console.log(JSON.stringify(response));
        localStorage.setItem('accessToken', response.accessToken);
        localStorage.setItem('idToken', response.idToken);
      });
  }

  logout(popup?: boolean) {
    if (popup) {
      this.msalService.logoutPopup({
        mainWindowRedirectUri: '/',
      });
    } else {
      this.msalService.logoutRedirect();
    }

    localStorage.removeItem('accessToken');
    localStorage.removeItem('idToken');
    sessionStorage.clear();
    localStorage.clear();
    this.loggingSubscription?.unsubscribe();
  }
  isUserLoggedIn(): boolean {
    return this.msalService.instance.getActiveAccount() !== null;
  }
}
