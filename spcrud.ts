imports(Http, Response, Headers) from '@angular/http';

// SPCrud - Angular 2 TypeScript Module
export module SPCrud {
    class SPCrud {
        // private properties
        baseUrl = '';
        apiUrl = '';
        headers;

        constructor(http: Http) {
            setBaseUrl();
        }

        //----------SHAREPOINT GENERAL----------

        //initialize
        init() {
            //default to local web URL
            this.apiUrl = this.baseUrl + '/_api/web/lists/GetByTitle(\'{0}\')/items';

            //globals
            const jsonHeader = 'application/json;odata=verbose';
            this.headers = {
                'Content-Type': jsonHeader,
                'Accept': jsonHeader
            };

            //request digest
            var el = document.querySelector('#__REQUESTDIGEST');
            if (el) {
                //digest local to ASPX page
                this.headers['X-RequestDigest'] = el.value;
            }
        };

        //change target web URL
        setBaseUrl(webUrl) {
            if (webUrl) {
                //user provided target Web URL
                spcrud.baseUrl = webUrl;
            } else {
                //default local SharePoint context
                if (window._spPageContextInfo) {
                    this.baseUrl = window._spPageContextInfo.webAbsoluteUrl;
                }
            }
            init();
        };


        //string ends with
        endsWith(str, suffix) {
            return str.indexOf(suffix, str.length - suffix.length) !== -1;
        };

        //digest refresh worker
        refreshDigest() {
            var config = {
                url: this.baseUrl + '/_api/contextinfo',
                headers: this.headers
            };
            return this.http.post(config).subscribe((res: Response) => res.json()).then(function (response) {
                //parse JSON and save
                this.headers['X-RequestDigest'] = response.data.d.GetContextWebInformation.FormDigestValue;
            });

        };

        //send email
        sendMail = function (to, ffrom, subj, body) {
            //append metadata
            to = to.split(",");
            var recip = (to instanceof Array) ? to : [to],
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
                    url: spcrud.baseUrl + '/_api/SP.Utilities.Utility.SendEmail',
                    headers: spcrud.headers,
                    data: message.json()
                };
            return this.http.post(config).subscribe((res: Response) => res.json());
        };

        //----------SHAREPOINT USER PROFILES----------

        //lookup SharePoint current web user
        getCurrentUser() {
            if (!spcrud.currentUser) {
                var url = spcrud.baseUrl + '/_api/web/currentuser?$expand=Groups',
                    config = {
                        url: url,
                        cache: true,
                        headers: spcrud.headers
                    };
                return this.http.get(config).subscribe((res: Response) => res.json());
            }
        };

        //lookup my SharePoint profile
        getMyProfile() {
            if (!spcrud.myProfile) {
                var url = spcrud.baseUrl + '/_api/SP.UserProfiles.PeopleManager/GetMyProperties?select=*',
                    config = {
                        url: url,
                        cache: true,
                        headers: spcrud.headers
                    };
                return this.http.get(config).subscribe((res: Response) => res.json());
            }
        };

        //lookup any SharePoint profile
        getProfile(login) {
            var url = spcrud.baseUrl + '/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v=\'' + login + '\'&select=*',
                config = {
                    url: url,
                    headers: spcrud.headers
                };
            return this.http.get(config).subscribe((res: Response) => res.json());
        };

        //lookup any SharePoint UserInfo
        getUserInfo(Id) {
            var url = spcrud.baseUrl + '/_api/web/getUserById(' + Id + ')',
                config = {
                    url: url,
                    headers: spcrud.headers
                };
            return this.http.get(config).subscribe((res: Response) => res.json());
        };

        //ensure SPUser exists in target web
        ensureUser(login) {
            var url = spcrud.baseUrl + '/_api/web/ensureuser',
                config = {
                    url: url,
                    headers: spcrud.headers,
                    data: login
                };
            return this.http.post(config).subscribe((res: Response) => res.json());
        };


        //----------SHAREPOINT LIST AND FIELDS----------
        //create list
        createList(title, baseTemplate, description) {
            var data = {
                '__metadata': { 'type': 'SP.List' },
                'BaseTemplate': baseTemplate,
                'Description': description,
                'Title': title
            },
                url = spcrud.baseUrl + '/_api/web/lists',
                config = {
                    url: url,
                    headers: spcrud.headers,
                    data: data
                };
            return this.http.post(config).subscribe((res: Response) => res.json());
        };

        //create list
        createField(listTitle, fieldName, fieldType) {
            var data = {
                '__metadata': { 'type': 'SP.Field' },
                'Type': fieldType,
                'Title': fieldName
            },
                url = spcrud.baseUrl + '/_api/web/lists/GetByTitle(\'' + listTitle + '\')/fields',
                config = {
                    url: url,
                    headers: spcrud.headers,
                    data: data
                };
            return this.http.post(config).subscribe((res: Response) => res.json());
        };

        //----------SHAREPOINT FILES AND FOLDERS----------

        //create folder
        createFolder(folderUrl) {
            var data = {
                '__metadata': {
                    'type': 'SP.Folder'
                },
                'ServerRelativeUrl': folderUrl
            },
                url = spcrud.baseUrl + '/_api/web/folders',
                config = {
                    url: url,
                    headers: spcrud.headers,
                    data: data
                };
            return this.http.post(config).subscribe((res: Response) => res.json());
        };

        // upload file to folder
        // https://kushanlahiru.wordpress.com/2016/05/14/file-attach-to-sharepoint-2013-list-custom-using-angular-js-via-rest-api/
        // http://stackoverflow.com/questions/17063000/ng-model-for-input-type-file
        // var binary = new Uint8Array(FileReader.readAsArrayBuffer(file[0]));
        uploadFile(folderUrl, fileName, binary) {
            var url = spcrud.baseUrl + '/_api/web/GetFolderByServerRelativeUrl(\'' + folderUrl + '\')/files/add(overwrite=true, url=\'' + fileName + '\')',
                config = {
                    url: url,
                    headers: spcrud.headers,
                    data: binary,
                    transformRequest: []
                };
            return this.http.post(config).subscribe((res: Response) => res.json());
        };

        //upload attachment to item
        uploadAttach(listName, id, fileName, binary, overwrite) {
            var url = spcrud.baseUrl + '/_api/web/lists/GetByTitle(\'' + listName + '\')/items(' + id,
                headers = JSON.parse(JSON.stringify(spcrud.headers));

            if (overwrite) {
                //append HTTP header PUT for UPDATE scenario
                headers['X-HTTP-Method'] = 'PUT';
                url += ')/AttachmentFiles(\'' + fileName + '\')/$value';
            } else {
                //CREATE scenario
                url += ')/AttachmentFiles/add(FileName=\'' + fileName + '\')';
            }

            var config = {
                url: url,
                headers: headers,
                data: binary
            };
            return this.http.post(config).subscribe((res: Response) => res.json());
        };

        //get attachment for item
        getAttach(listName, id) {
            var url = spcrud.baseUrl + '/_api/web/lists/GetByTitle(\'' + listName + '\')/items(' + id + ')/AttachmentFiles',
                config = {
                    method: 'GET',
                    url: url,
                    headers: spcrud.headers
                };
            return this.http.get(config).subscribe((res: Response) => res.json());
        };

        //copy file
        copyFile(sourceUrl, destinationUrl) {
            var url = spcrud.baseUrl + '/_api/web/getfilebyserverrelativeurl(\'' + sourceUrl + '\')/copyto(strnewurl=\'' + destinationUrl + '\',boverwrite=false)',
                config = {
                    method: 'POST',
                    url: url,
                    headers: spcrud.headers
                };
            return this.http.post(config).subscribe((res: Response) => res.json());
        };

        //----------SHAREPOINT LIST CORE----------

        //CREATE item - SharePoint list name, and JS object to stringify for save
        create(listName, jsonBody) {
            //append metadata
            if (!jsonBody.__metadata) {
                jsonBody.__metadata = {
                    'type': 'SP.ListItem'
                };
            }
            var data = jsonBody.json(),
                config = {
                    url: spcrud.apiUrl.replace('{0}', listName),
                    data: data,
                    headers: spcrud.headers
                };
            return this.http.post(config).subscribe((res: Response) => res.json());
        };

        readBuilder(url, options) {
            if (options) {
                if (options.filter) {
                    url += ((spcrud.endsWith(url, 'items')) ? "?" : "&") + "$filter=" + options.filter;
                }
                if (options.select) {
                    url += ((spcrud.endsWith(url, 'items')) ? "?" : "&") + "$select=" + options.select;
                }
                if (options.orderby) {
                    url += ((spcrud.endsWith(url, 'items')) ? "?" : "&") + "$orderby=" + options.orderby;
                }
                if (options.expand) {
                    url += ((spcrud.endsWith(url, 'items')) ? "?" : "&") + "$expand=" + options.expand;
                }
                if (options.top) {
                    url += ((spcrud.endsWith(url, 'items')) ? "?" : "&") + "$top=" + options.top;
                }
                if (options.skip) {
                    url += ((spcrud.endsWith(url, 'items')) ? "?" : "&") + "$skip=" + options.skip;
                }
            }
            return url;
        }

        //READ entire list - needs $http factory and SharePoint list name
        read(listName, options) {
            //build URL syntax
            //https://msdn.microsoft.com/en-us/library/office/fp142385.aspx#bk_support
            var url = this.apiUrl.replace('{0}', listName);
            url = this.readBuilder(url, options);

            //config
            var config = {
                url: url,
                headers: spcrud.headers
            };
            return this.http.get(config).subscribe((res: Response) => res.json());
        };

        //READ single item - SharePoint list name, and item ID number
        readItem = function (listName, id) {
            var url = spcrud.apiUrl.replace('{0}', listName) + '(' + id + ')'
            url = spcrud.readBuilder(url, options);

            //config
            var config = {
                url: url,
                headers: spcrud.headers
            };
            return this.http.get(config).subscribe((res: Response) => res.json());
        };

        //UPDATE item - SharePoint list name, item ID number, and JS object to stringify for save
        update (listName, id, jsonBody) {
            //append HTTP header MERGE for UPDATE scenario
            var headers = JSON.parse(JSON.stringify(spcrud.headers));
            headers['X-HTTP-Method'] = 'MERGE';
            headers['If-Match'] = '*';

            //append metadata
            if (!jsonBody.__metadata) {
                jsonBody.__metadata = {
                    'type': 'SP.ListItem'
                };
            }
            var data = jsonBody.json(),
                config = {
                    url: spcrud.apiUrl.replace('{0}', listName) + '(' + id + ')',
                    data: data,
                    headers: headers
                };
            return this.http.post(config).subscribe((res: Response) => res.json());
        };

        //DELETE item - SharePoint list name and item ID number
        del = function (listName, id) {
            //append HTTP header DELETE for DELETE scenario
            var headers = JSON.parse(JSON.stringify(spcrud.headers));
            headers['X-HTTP-Method'] = 'DELETE';
            headers['If-Match'] = '*';
            var config = {
                url: spcrud.apiUrl.replace('{0}', listName) + '(' + id + ')',
                headers: headers
            };
            return this.http.post(config).subscribe((res: Response) => res.json());
        };


        //... JSON ... Social

    }
}