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
  EventMessage,
  EventType,
  InteractionStatus,
  PopupRequest,
  RedirectRequest,
} from '@azure/msal-browser';
import { Subject, Subscription } from 'rxjs';
import { filter, takeUntil, tap } from 'rxjs/operators';

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
      .subscribe(()=> {
        this.setLoginDisplay();
        this.checkAndSetActiveAccount();
      });

    if (!this.isUserLoggedIn()) {
      this.loginLinkText = 'Login';
      this.loggingSubscription?.unsubscribe();
    }
  }
  ngOnDestroy(): void {
    this.onDestroy$.next(undefined);
    this.onDestroy$.complete();
  }

  setLoginDisplay() {
    this.loginDisplay = this.msalService.instance.getAllAccounts().length > 0;
  }

  checkAndSetActiveAccount() {
    let activeAccount = this.msalService.instance.getActiveAccount();

    if (!activeAccount && this.msalService.instance.getAllAccounts().length > 0) {
      let accounts = this.msalService.instance.getAllAccounts();
      this.msalService.instance.setActiveAccount(accounts[0]);
      
    }
  }

 

  login(popup?: boolean): void {
    if(popup){
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
    else{
      if (this.msalGuardConfig.authRequest) {
        this.msalService.loginRedirect({
          ...this.msalGuardConfig.authRequest,
        } as RedirectRequest);
      } else {
        this.msalService.loginRedirect();
      }
    }
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
