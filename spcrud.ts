/**
 * Library with AngularJS operations for CRUD operations to SharePoint 2013/2016/Online lists over REST api
 *
 * Contains 6 core functions and other misc helper functions
 *
 * 1) Create    - add item to List
 * 2) Read      - find all items or single item from List
 * 3) Update    - update item in List
 * 4) Delete    - delete item in List
 * 5) jsonRead  - read JSON to List
 * 6) jsonWrite - write JSON to List ("upsert" = add if missing, update if exists)
 *
 * NOTE - 5 and 6 require the target SharePoint List to have two columns: "Title" (indexed) and "JSON" (mult-text).   These are
 * intendend to save JSON objects for JS internal application needs.   For example, saving user preferences to a "JSON-Settings" list
 * where one row is created per user (Title = current user Login) and JSON multi-text field holds the JSON blob.
 * Simple and flexible way to save data for many scenarios.
 *
 * @spjeff
 * spjeff@spjeff.com
 * http://spjeff.com
 *
 * version 0.2.01
 * last updated 06-02-2017
 *
 */

import { Injectable } from '@angular/core';

// RxJS dependency
import { Http, Headers, Response, RequestOptions } from '@angular/http';
import { Observable } from 'rxjs/Observable';
import 'rxjs/add/operator/catch';
import 'rxjs/add/operator/map';
import 'rxjs/add/operator/toPromise';

@Injectable()
export class Spcrud {

  // Data
  jsonHeader = 'application/json; odata=verbose';
  headers = new Headers({ 'Content-Type': this.jsonHeader, 'Accept': this.jsonHeader });
  options = new RequestOptions({ headers: this.headers });
  baseUrl: String;
  apiUrl: String;

  constructor(private http: Http) {
    this.setBaseUrl(null);
  }

  // HTTP Error handling
  private handleError(error: Response | any) {
    // generic from https://angular.io/docs/ts/latest/guide/server-communication.html
    let errMsg: string;
    if (error instanceof Response) {
      const body = error.json() || '';
      const err = body.error || JSON.stringify(body);
      errMsg = `${error.status || ''} - ${error.statusText || ''} ${err}`;
    } else {
      errMsg = error.message ? error.message : error.toString();
    }
    console.error(errMsg);
    return Observable.throw(errMsg);
  }

  // String ends with
  private endsWith(str, suffix) {
    return str.indexOf(suffix, str.length - suffix.length) !== -1;
  }

  // ----------SHAREPOINT GENERAL----------

  // Set base working URL path
  setBaseUrl(webUrl: String) {
    if (webUrl) {
      // user provided target Web URL
      this.baseUrl = webUrl;
    } else {
      // default local SharePoint context
      const ctx = window['_spPageContextInfo'];
      if (ctx) {
        this.baseUrl = ctx.webAbsoluteUrl;
      }
    }

    // Default to local web URL
    this.apiUrl = this.baseUrl + '/_api/web/lists/GetByTitle(\'{0}\')/items';

    // Request digest
    const el = document.querySelector('#__REQUESTDIGEST');
    if (el) {
      // Digest local to ASPX page
      // this.headers.delete('X-RequestDigest');
      this.headers.append('X-RequestDigest', el.nodeValue);
    }
  }

  // Data methods
  getData(): Observable<any> {
    return this.http.get('http://portal/sites/todo/_api/web/lists', this.options).map(function (res: Response) {
      return res.json() || {};
    }).catch(this.handleError);
  }

  // Refresh digest token
  refreshDigest(): Promise<any> {
    return this.http.post('/_api/contextinfo', this.options).toPromise().then(function (res: Response) {
      // this.headers.delete('X-RequestDigest');
      this.headers.append('X-RequestDigest', res.json().data.d.GetContextWebInformation.FormDigestValue);
    });
  }

  // send email ... TBD ...

  // send email
sendMail  (to: string, ffrom: string, subj: string, body: string) {
    //append metadata
    var tos: string[] = to.split(",");
    var recip: string[] = (tos instanceof Array) ? tos : [tos],
        message = {
            'properties': {
                '__metadata': {
                    'type': 'SP.Utilities.EmailProperties'
                },
                'To': {
                    'results': recip
                },
                'From': ffrom,
                'Subject': subj,
                'Body': body
            }
        },
        config = {
            method: 'POST',
            url: spcrud.baseUrl + '/_api/SP.Utilities.Utility.SendEmail',
            headers: spcrud.headers,
            data: angular.toJson(message)
        };
    return $http(config);
};



  // ----------SHAREPOINT USER PROFILES----------

  // Lookup SharePoint current web user
  getCurrentUser(): Observable<any> {
    const url = this.baseUrl + '/_api/web/currentuser?$expand=Groups';
    return this.http.get(url, this.options).map(function (res: Response) {
      return res.json() || {};
    }).catch(this.handleError);
  };


  // JSON blob read from SharePoint list - SharePoint list name
  jsonRead(listName: string): Promise<any> {

    return this.getCurrentUser().toPromise().then(function (res: Response) {
      // GET SharePoint Current User
      const d = res.json().data.d;
      this.currentUser = d;
      this.login = d.LoginName.toLowerCase();
      if (this.login.indexOf('\\')) {
        // Parse domain prefix
        this.login = this.login.split('\\')[1];
      }

      // GET SharePoint List Item
      const url = this.apiUrl.replace('{0}', listName) + '?$select=JSON,Id,Title&$filter=Title+eq+\'' + this.login + '\'';
      return this.http.get(url, this.options).map(function (res2: Response) {

        // Parse JSON response
        const d2 = res2.json().data.d;
        if (d2.results) {
          return d2.results[0];
        } else {
          return null;
        }

      }).catch(this.handleError);
    });
  };


//JSON blob upsert write to SharePoint list - SharePoint list name and JS object to stringify for save
jsonWrite  (listName: string, jsonBody: any) {
    return this.refreshDigest().then(function (res: Response) {
        return this.jsonRead(listName).then(function (item: any) {
            //HTTP 200 OK
            if (item) {
                //update if found
                item.JSON = JSON.stringify(jsonBody);
                return this.update(listName, item.Id, item);
            } else {
                //create if missing
                item = {
                    '__metadata': {
                        'type': 'SP.ListItem'
                    },
                    'Title': this.login,
                    'JSON': JSON.stringify(jsonBody);
                };
                return this.create(listName, item);
            }
        });
    });
};

  // **
}
