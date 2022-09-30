import { Component, OnInit } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { MsalService } from '@azure/msal-angular';
import {DomSanitizer} from '@angular/platform-browser';

const GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/me';

type ProfileType = {
  givenName?: string;
  surname?: string;
  userPrincipalName?: string;
  id?: string;
};


@Component({
  selector: 'app-profile',
  templateUrl: './profile.component.html',
  styleUrls: ['./profile.component.scss'],
})
export class ProfileComponent implements OnInit {
  profile!: ProfileType;
  imgUrl:string = "";
  constructor(private http: HttpClient, private msalService: MsalService,
    private sanitizer:DomSanitizer) {}

  ngOnInit() {
    this.getProfile();
    this.getProfilePhoto();
  }

  getProfile() {
    this.http
      .get(GRAPH_ENDPOINT, {
        headers: {
          Authorization: 'Bearer ' + localStorage.getItem('accessToken'),
        },
      })
      .subscribe((profile) => {
        this.profile = profile;
        //console.log("profile: " + JSON.stringify(this.profile));
      });
  }

   
  getProfilePhoto() {
    this.http
      .get('https://graph.microsoft.com/v1.0/me/photo/$value', {
        headers: {
          Authorization: 'Bearer ' + localStorage.getItem('accessToken'),
          'Content-Type': 'image/jpg' 
        },
        responseType:'blob'
      })
      .subscribe((response:any) => {
        // const url = window.URL || window.webkitURL;
        const blobUrl = window.URL.createObjectURL(response);
        //console.log('Image: ' + JSON.stringify(blobUrl));
        this.imgUrl = blobUrl;
        // const pictureBlob = response.blob();
        // console.log('pictureBlob ' + pictureBlob);
        // console.log('blobUrl: ' +blobUrl);
      });
      
  }

  sanitize(url:string){
    return this.sanitizer.bypassSecurityTrustUrl(url);
}

}
