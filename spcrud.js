/**
 * Library with AngularJS operations for CRUD operations to SharePoint 2013 lists over REST api
 *
 * Contains 6 core functions
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
 * where one row is created per user (Title = current user Login) and JSON multi-text field holds the JSON blob.  Simple and flexible way to save data
 * for many scenarios.
 *
 * @spjeff
 * spjeff@spjeff.com
 * http://spjeff.com
 *
 * last updated 12-10-2015
 */

//namespace
var spcrud = spcrud || {};

//----------SHARED----------

//initialize
spcrud.init = function () {
    //default to local web URL
    spcrud.apiUrl = spcrud.baseUrl + '/_api/web/lists/getbytitle(\'{0}\')/items';
    
    //globals
    spcrud.jsonHeader = 'application/json;odata=verbose';
    spcrud.headers = {
        'Content-Type': spcrud.jsonHeader,
        'Accept': spcrud.jsonHeader
    };

    //request digest
    var el = document.querySelector('#__REQUESTDIGEST');
    if (el) {
        //digest local to ASPX page
        spcrud.headers['X-RequestDigest'] = el.value;
    }
};

//change target web URL
spcrud.setBaseUrl = function(webUrl) {
    if (webUrl) {
        //user provided target Web URL
        spcrud.baseUrl = webUrl;
    } else {
        //default local SharePoint context
        if (_spPageContextInfo) {
            spcrud.baseUrl = _spPageContextInfo.webAbsoluteUrl;
        }
    }
    spcrud.init();
};
spcrud.setBaseUrl();

//string helper function - endsWith()
spcrud.endsWith(str, suffix) {
    return str.indexOf(suffix, str.length - suffix.length) !== -1;
}

//digest refresh worker
spcrud.refreshDigest = function ($http) {
    var config = {
        method: 'POST',
        url: spcrud.baseUrl + '/_api/contextinfo';
        headers: spcrud.headers
    };
    return $http(config).then(function (response) {
        //parse JSON and save
        spcrud.headers['X-RequestDigest'] = response.data.d.GetContextWebInformation.FormDigestValue;
    });

};

//lookup SharePoint current web user
spcrud.getCurrentUser = function ($http) {
    var url = spcrud.baseUrl + '/_api/web/currentuser';
    var config = {
        method: 'GET',
        url: url,
        cache: true,
        headers: spcrud.headers
    };
    return $http(config);
};

//lookup SharePoint my profile
spcrud.getMyProfile = function ($http) {
    var url = spcrud.baseUrl + '/_api/SP.UserProfiles.PeopleManager/GetMyProperties';
    var config = {
        method: 'GET',
        url: url,
        cache: true,
        headers: spcrud.headers
    };
    return $http(config);
};

//ensure SPUser exists in target web
spcrud.ensureUser = function($http, login) {
    var url = spcrud.baseUrl + '/_api/web/ensureuser';
    var config = {
        method: 'POST',
        url: url,
        headers: spcrud.headers,
        data: login
    };
    return $http(config);
};

//----------CORE----------

//CREATE item - needs $http factory, SharePoint list name, and JS object to stringify for save
spcrud.create = function ($http, listName, jsonBody) {
    //append metadata
    if (!jsonBody['__metadata']) {
        jsonBody['__metadata'] = {
            'type': 'SP.ListItem'
        };
    }
    var data = angular.toJson(jsonBody);

    var config = {
        method: 'POST',
        url: spcrud.apiUrl.replace('{0}', listName),
        data: data,
        headers: spcrud.headers
    };
    return $http(config);
};

//READ entire list - needs $http factory and SharePoint list name
spcrud.read = function ($http, listName, filter, selectt, orderby) {
    //build URL syntax
    var url = spcrud.apiUrl.replace('{0}', listName);
    if (filter) {
        url += ((spcrud.endsWith(url, 'items')) ? "?" : "&") + "$filter=" + filter;
    }
    if (selectt) {
        url += ((spcrud.endsWith(url, 'items')) ? "?" : "&") + "$select=" + selectt;
    }
    if (orderby) {
        url += ((spcrud.endsWith(url, 'items')) ? "?" : "&") + "$orderby=" + orderby;
    }

    //config
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//READ single item - needs $http factory, SharePoint list name, and item ID number
spcrud.readItem = function ($http, listName, id) {
    var config = {
        method: 'GET',
        url: spcrud.apiUrl.replace('{0}', listName) + '(' + id + ')',
        headers: spcrud.headers
    };
    return $http(config);
};

//UPDATE item - needs $http factory, SharePoint list name, item ID number, and JS object to stringify for save
spcrud.update = function ($http, listName, id, jsonBody, webUrl) {
    //append HTTP header MERGE for UPDATE scenario
    var headers = JSON.parse(JSON.stringify(spcrud.headers));
    headers['X-HTTP-Method'] = 'MERGE';
    headers['If-Match'] = '*';

    //append metadata
    if (!jsonBody['__metadata']) {
        jsonBody['__metadata'] = {
            'type': 'SP.ListItem'
        };
    }
    var data = angular.toJson(jsonBody);

    var config = {
        method: 'POST',
        url: spcrud.apiUrl.replace('{0}', listName) + '(' + id + ')',
        data: data,
        headers: headers
    };
    return $http(config);
};

//DELETE item - needs $http factory, SharePoint list name and item ID number
spcrud.del = function ($http, listName, id) {
    //append HTTP header DELETE for DELETE scenario
    var headers = JSON.parse(JSON.stringify(spcrud.headers));
    headers['X-HTTP-Method'] = 'DELETE';
    headers['If-Match'] = '*';

    var config = {
        method: 'POST',
        url: spcrud.apiUrl.replace('{0}', listName) + '(' + id + ')',
        headers: headers
    };
    return $http(config);
};

//JSON blob read from SharePoint list - needs $http factory and SharePoint list name
spcrud.jsonRead = function ($http, listName, cache) {
    return spcrud.getCurrentUser($http).then(function (response) {
        //GET SharePoint Current User
        spcrud.currentUser = response.data.d;
        spcrud.login = response.data.d.LoginName.toLowerCase();
        if (spcrud.login.indexOf('\\')) {
            //parse domain prefix
            spcrud.login = spcrud.login.split('\\')[1];
        }

        //default no caching
        if (!cache) {
            cache = false;
        }

        //GET SharePoint list item(s)
        var config = {
            method: 'GET',
            url: spcrud.apiUrl.replace('{0}', listName) + '?$select=JSON,Id,Title&$filter=Title+eq+\'' + spcrud.login + '\'',
            cache: cache,
            headers: spcrud.headers
        };

        //GET SharePoint Profile
        spcrud.getMyProfile($http).then(function (response) {
            spcrud.myProfile = response.data.d;
        });

        //parse single SPListItem only
        return $http(config).then(function (response) {
            if (response.data.d.results) {
                return response.data.d.results[0];
            } else {
                return null;
            }
        });
    });
};

//JSON blob upsert write to SharePoint list - needs $http factory, SharePoint list name, and JS object to stringify for save
spcrud.jsonWrite = function ($http, listName, jsonBody) {
    return spcrud.refreshDigest($http).then(function (response) {
        return spcrud.jsonRead($http, listName).then(function (item) {
            //HTTP 200 OK
            if (item) {
                //update if found
                item.JSON = angular.toJson(jsonBody);
                return spcrud.update($http, listName, item.Id, item);
            } else {
                //create if missing
                var item = {
                    '__metadata': {
                        'type': 'SP.ListItem'
                    },
                    'Title': spcrud.login,
                    'JSON': angular.toJson(jsonBody)
                };
                return spcrud.create($http, listName, item);
            }
        });
    });
};