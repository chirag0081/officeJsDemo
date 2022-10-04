import { logging } from 'protractor';
import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { HttpClientModule, HTTP_INTERCEPTORS } from '@angular/common/http';
import { AppRoutingModule } from './app-routing.module';
import { MsalGuard, MsalInterceptor, MsalModule, MsalService, MSAL_INSTANCE } from '@azure/msal-angular';
import { InteractionType, IPublicClientApplication, PublicClientApplication } from '@azure/msal-browser';
import { AppComponent } from './app.component';
import { DocumentListComponent } from './document-list/document-list.component';
import { HomeComponent } from './home/home.component';
import { PageNotFoundComponent } from './page-not-found/page-not-found.component';
import { ProfileComponent } from './profile/profile.component';
import { LoginComponent } from './login/login.component';

const isIE = window.navigator.userAgent.indexOf("MSIE ") > -1 || window.navigator.userAgent.indexOf("Trident/") > -1;
export function MSALInstanceFactory(): IPublicClientApplication {
  return new PublicClientApplication({
    auth: { 
      clientId:"bccf936a-d9b3-4ced-91f1-871ffbedb83a",
      redirectUri: "https://localhost:4200",
      postLogoutRedirectUri: "https://localhost:4200"
    }
  });
}

@NgModule({
  declarations: [
    AppComponent,
    DocumentListComponent,
    ProfileComponent,
    HomeComponent,
    PageNotFoundComponent,
    LoginComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    HttpClientModule,
    MsalModule.forRoot(
      new PublicClientApplication({
        auth: {
          clientId: "bccf936a-d9b3-4ced-91f1-871ffbedb83a",
          redirectUri: "https://localhost:4200",
          postLogoutRedirectUri: "https://localhost:4200"
        },
        cache: {
          cacheLocation: "localStorage" ,
          storeAuthStateInCookie: isIE, 
        },
      }),
      {
        interactionType: InteractionType.Popup, // Msal Guard Configuration
        authRequest: {
          scopes: ["user.read"],
        },
      },
      null
    ) 
  ],
  providers: [
    // {
    //   provide: MSAL_INSTANCE,
    //   useFactory: MSALInstanceFactory,
    // },
    MsalService ,
  ],
  bootstrap: [AppComponent]
})
export class AppModule { }
