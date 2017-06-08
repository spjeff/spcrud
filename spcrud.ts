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
 * version 0.2.04
 * last updated 06-08-2017
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
    // Generic from https://angular.io/docs/ts/latest/guide/server-communication.html
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

  // Send email
  sendMail(to: string, ffrom: string, subj: string, body: string): Promise<any> {
    // Append metadata
    const tos: string[] = to.split(',');
    const recip: string[] = (tos instanceof Array) ? tos : [tos];
    const message = {
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
    };
    const url = this.baseUrl + '/_api/SP.Utilities.Utility.SendEmail';
    const data = JSON.stringify(message);
    return this.http.post(url, data, this.options).toPromise();
  };

  // ----------SHAREPOINT USER PROFILES----------

  // Lookup SharePoint current web user
  getCurrentUser(): Observable<any> {
    const url = this.baseUrl + '/_api/web/currentuser?$expand=Groups';
    return this.http.get(url, this.options).map(function (res: Response) {
      return res.json() || {};
    }).catch(this.handleError);
  };


  // Lookup my SharePoint profile
  getMyProfile(): Observable<any> {
    const url = this.baseUrl + '/_api/SP.UserProfiles.PeopleManager/GetMyProperties?select=*';
    return this.http.get(url, this.options);
  };

  // Lookup any SharePoint profile
  getProfile(login: string): Observable<any> {
    const url = this.baseUrl + '/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v=\'' + login + '\'&select=*';
    return this.http.get(url, this.options);
  };

  // Lookup any SharePoint UserInfo
  getUserInfo(Id: string): Observable<any> {
    const url = this.baseUrl + '/_api/web/getUserById(' + Id + ')';
    return this.http.get(url);
  };

  // Ensure SPUser exists in target web
  ensureUser(login: string): Observable<any> {
    const url = this.baseUrl + '/_api/web/ensureuser';
    return this.http.post(url, login, this.options);
  };

  // ----------SHAREPOINT LIST AND FIELDS----------
  // Create list
  createList(title: string, baseTemplate: string, description: string): Observable<any> {
    const data = {
      '__metadata': { 'type': 'SP.List' },
      'BaseTemplate': baseTemplate,
      'Description': description,
      'Title': title
    };
    const url = this.baseUrl + '/_api/web/lists';
    return this.http.post(url, data, this.options);
  };

  // Create field
  createField(listTitle: string, fieldName: string, fieldType: string): Observable<any> {
    const data = {
      '__metadata': { 'type': 'SP.Field' },
      'Type': fieldType,
      'Title': fieldName
    };
    const url = this.baseUrl + '/_api/web/lists/GetByTitle(\'' + listTitle + '\')/fields';
    return this.http.post(url, this.options, data);
  };


  // ----------SHAREPOINT FILES AND FOLDERS----------

  // Create folder
  createFolder(folderUrl: string): Observable<any> {
    const data = {
      '__metadata': {
        'type': 'SP.Folder'
      },
      'ServerRelativeUrl': folderUrl
    };
    const url = this.baseUrl + '/_api/web/folders';
    return this.http.post(url, this.options, data);
  };

  // Upload file to folder
  // https://kushanlahiru.wordpress.com/2016/05/14/file-attach-to-sharepoint-2013-list-custom-using-angular-js-via-rest-api/
  // http://stackoverflow.com/questions/17063000/ng-model-for-input-type-file
  // var binary = new Uint8Array(FileReader.readAsArrayBuffer(file[0]));
  uploadFile(folderUrl: string, fileName: string, binary: any): Observable<any> {
    const url = this.baseUrl + '/_api/web/GetFolderByServerRelativeUrl(\'' + folderUrl + '\')/files/add(overwrite=true, url=\'' + fileName + '\')';
    return this.http.post(url, this.options, binary);
  };

  // Upload attachment to item
  uploadAttach(listName: string, id: string, fileName: string, binary: any, overwrite: boolean): Observable<any> {
    let url = this.baseUrl + '/_api/web/lists/GetByTitle(\'' + listName + '\')/items(' + id;
    const options = this.options;
    if (overwrite) {
      // Append HTTP header PUT for UPDATE scenario
      options.headers.append('X-HTTP-Method', 'PUT');
      url += ')/AttachmentFiles(\'' + fileName + '\')/$value';
    } else {
      // CREATE scenario
      url += ')/AttachmentFiles/add(FileName=\'' + fileName + '\')';
    }
    return this.http.post(url, options, binary);
  };

  // Get attachment for item
  getAttach(listName: string, id: string): Observable<any> {
    const url = this.baseUrl + '/_api/web/lists/GetByTitle(\'' + listName + '\')/items(' + id + ')/AttachmentFiles';
    return this.http.get(url, this.options);
  };

  // Copy file
  copyFile(sourceUrl: string, destinationUrl: string): Observable<any> {
    const url = this.baseUrl + '/_api/web/GetFileByServerRelativeUrl(\'' + sourceUrl + '\')/copyto(strnewurl=\'' + destinationUrl + '\',boverwrite=false)';
    return this.http.post(url, this.options);
  };

  // ----------SHAREPOINT LIST CORE----------

  // CREATE item - SharePoint list name, and JS object to stringify for save
  create(listName: string, jsonBody: any): Observable<any> {
    const url = this.apiUrl.replace('{0}', listName);
    // append metadata
    if (!jsonBody.__metadata) {
      jsonBody.__metadata = {
        'type': 'SP.ListItem'
      };
    }
    const data = JSON.stringify(jsonBody);
    return this.http.post(url, data, this.options);
  };

  // Build URL string with OData parameters
  readBuilder(url: string, options: any): string {
    if (options) {
      if (options.filter) {
        url += ((this.endsWith(url, 'items')) ? '?' : '&') + '$filter=' + options.filter;
      }
      if (options.select) {
        url += ((this.endsWith(url, 'items')) ? '?' : '&') + '$select=' + options.select;
      }
      if (options.orderby) {
        url += ((this.endsWith(url, 'items')) ? '?' : '&') + '$orderby=' + options.orderby;
      }
      if (options.expand) {
        url += ((this.endsWith(url, 'items')) ? '?' : '&') + '$expand=' + options.expand;
      }
      if (options.top) {
        url += ((this.endsWith(url, 'items')) ? '?' : '&') + '$top=' + options.top;
      }
      if (options.skip) {
        url += ((this.endsWith(url, 'items')) ? '?' : '&') + '$skip=' + options.skip;
      }
    }
    return url;
  };

  // READ entire list - needs $http factory and SharePoint list name
  read(listName: string, options: any): Observable<any> {
    // Build URL syntax
    // https://msdn.microsoft.com/en-us/library/office/fp142385.aspx#bk_support
    let url = this.apiUrl.replace('{0}', listName);
    url = this.readBuilder(url, options);
    return this.http.get(url, this.options);
  };


  // READ single item - SharePoint list name, and item ID number
  readItem(listName: string, id: string): Observable<any> {
    let url = this.apiUrl.replace('{0}', listName) + '(' + id + ')';
    url = this.readBuilder(url, null);
    return this.http.get(url, this.options);
  };

  // UPDATE item - SharePoint list name, item ID number, and JS object to stringify for save
  update(listName: string, id: string, jsonBody: any) {
    // Append HTTP header MERGE for UPDATE scenario
    const options = JSON.parse(JSON.stringify(this.options));
    options.headers.append('X-HTTP-Method', 'MERGE');
    options.headers.append('If-Match', '*');

    // Append metadata
    if (!jsonBody.__metadata) {
      jsonBody.__metadata = {
        'type': 'SP.ListItem'
      };
    }
    const data = JSON.stringify(jsonBody);
    const url = this.apiUrl.replace('{0}', listName) + '(' + id + ')';
    return this.http.post(url, data, options);
  };

  // DELETE item - SharePoint list name and item ID number
  del(listName: string, id: string) {
    // append HTTP header DELETE for DELETE scenario
    const options = JSON.parse(JSON.stringify(this.options));
    options.headers.append('X-HTTP-Method', 'DELETE');
    options.headers.append('If-Match', '*');
    const url = this.apiUrl.replace('{0}', listName) + '(' + id + ')';
    return this.http.post(url, options);
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


  // JSON blob upsert write to SharePoint list - SharePoint list name and JS object to stringify for save
  jsonWrite(listName: string, jsonBody: any) {
    return this.refreshDigest().then(function (res: Response) {
      return this.jsonRead(listName).then(function (item: any) {
        // HTTP 200 OK
        if (item) {
          // update if found
          item.JSON = JSON.stringify(jsonBody);
          return this.update(listName, item.Id, item);
        } else {
          // create if missing
          item = {
            '__metadata': {
              'type': 'SP.ListItem'
            },
            'Title': this.login,
            'JSON': JSON.stringify(jsonBody)
          };
          return this.create(listName, item);
        }
      });
    });
  };
  // **
}
