import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import { environment } from 'src/environments/environment';

@Injectable({
  providedIn: 'root'
})
export class ExcelGerarService {

  constructor(private _http: HttpClient) { }

  getDados(): Observable<any[]> {
    return this._http.get<any[]>(environment.api_url + '/excel/dados', { responseType: 'json'});
  }

  gerarExcelCSharp_Blob(): Observable<any> {
    return this._http.get(environment.api_url + '/excel/gerar', { responseType: 'blob'});
  }

  gerarExcelCSharp_Base64(): Observable<any>  {
    return this._http.get<any>(environment.api_url + '/excel/gerar-base64', { responseType: 'json'});
  }
}
