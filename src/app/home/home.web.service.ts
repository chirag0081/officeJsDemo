import { Observable } from 'rxjs';
import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';

@Injectable({
  providedIn: 'root',
})
export class HomeWebService {
  constructor(private http: HttpClient) {}

  GetMembers(): Observable<any[]> {
    return this.http.get<any[]>('https://localhost:44386/api/MemberInfo');
  }

  PostPdfFile(pdfString: string): Observable<any> {
    
    const headers = { 'Content-Type': 'application/json'  };
    //const body=JSON.stringify(pdfString);
    //console.log(body)
    console.log('test');
    return this.http.post<any>(
      'https://localhost:44386/api/MemberInfo/download',
      JSON.stringify(pdfString),
      { headers: headers }
    );
  }
}
