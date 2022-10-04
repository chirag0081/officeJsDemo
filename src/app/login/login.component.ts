import { Component, Inject, OnInit } from '@angular/core';
import { MsalService  } from '@azure/msal-angular';
import { AuthenticationResult } from '@azure/msal-browser';

@Component({
  selector: 'app-login',
  templateUrl: './login.component.html',
  styleUrls: ['./login.component.scss']
})
export class LoginComponent implements OnInit {
  title = 'msal-angular-tutorial';
  constructor(private msalService: MsalService) { }

  ngOnInit(): void {
   // this.msalService.instance.loginRedirect();

//    this.msalService.loginRedirect().subscribe((response)=>{
//     console.log(JSON.stringify(response))
     
// });
  }
  
}
