/**
* Library with AngularJS operations for CRUD operations to SharePoint 2013 lists over REST api
*
* Contains 6 core functions
*
* 1) Create    - add item to List
* 2) Read      - find all items or single item from List
* 3) Update    - update item in List
* 4) Delete    - delete item in List
* 5) jsonRead  - save JSON to List
* 6) jsonWrite - read JSON to List
*
* NOTE - 5 and 6 require the target SharePoint List to have two columns: "Title" (indexed) and "JSON" (mult-text).   These are
* intendend to save JSON objects for JS internal application needs.   For example, saving user preferences to a "JSON-Settings" list
* where one row is created per user (Title = current user Login) and JSON multi-text field holds the blob.  Simple and flexible to save data
* for many scenarios.
*
* @spjeff
* spjeff@spjeff.com
* http://spjeff.com
*
* last updated 10-30-2015
*/

//namespace
var spcrud = spcrud || {};

//configuration
spcrud.apiUrl = _spPageContextInfo.webAbsoluteUrl + '/_api/web/lists/getbytitle(\'{0}\')/items';
spcrud.json = 'application/json;odata=verbose';
spcrud.headers = {
    'Content-Type': spcrud.json,
    'Accept': spcrud.json,
    'X-RequestDigest': document.querySelector('#__REQUESTDIGEST').value
};

//CREATE item - needs $http factory, SharePoint list name, and JS object to stringify for save
spcrud.create = function ($http, listName, jsonBody) {
    //append metadata
    jsonBody['__metadata'] = {
        'type': 'SP.ListItem'
    };
    var data = JSON.stringify(jsonBody);

    var config = {
        method: 'POST',
        url: spcrud.apiUrl.replace('{0}', listName),
        data: data,
        headers: spcrud.headers
    };
    return $http(config);
};

//READ entire list - needs $http factory and SharePoint list name
spcrud.read = function ($http, listName) {
    var config = {
        method: 'GET',
        url: spcrud.apiUrl.replace('{0}', listName),
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
spcrud.update = function ($http, listName, id, jsonBody) {
    //append HTTP header MERGE for UPDATE scenario
    var headers = JSON.parse(JSON.stringify(spcrud.headers));
    headers['X-HTTP-Method'] = 'MERGE';
    headers['If-Match'] = '*';

    //append metadata
    jsonBody['__metadata'] = {
        'type': 'SP.ListItem'
    };
    var data = JSON.stringify(jsonBody);

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

//parse SharePoint item ID from generic HTTP response
spcrud.getId = function (resp) {
    var item;
    if (resp.data.d.results) {
        //EDIT - nest JSON deeper in a 'results' object
        if (resp.data.d.results.constructor === Array) {
            item = resp.data.d.results[0];
        } else {
            item = resp.data.d.results;
        }
    } else {
        //READ - flatter JSON response format
        if (resp.data.d.constructor === Array) {
            item = resp.data.d[0];
        } else {
            item = resp.data.d;
        }
    }
    return (item) ? item.Id : null;
};

//lookup SharePoint Current User
spcrud.getCurrentUser = function ($http) {
    var url = _spPageContextInfo.webAbsoluteUrl + '/_api/web/currentuser';
    var config = {
        method: 'GET',
        url: url,
        cache: true,
        headers: spcrud.headers
    };
    return $http(config);
};

//JSON blob read from SharePoint list - needs $http factory and SharePoint list name
spcrud.jsonRead = function ($http, listName) {
    return spcrud.getCurrentUser($http).then(function (response) {
        //login name
        spcrud.login = response.data.d.LoginName;

        //GET SharePoint list
        var config = {
            method: 'GET',
            url: spcrud.apiUrl.replace('{0}', listName) + '?$select=JSON,Id,Title&$filter=Title+eq+\'' + spcrud.login + '\'',
            headers: spcrud.headers
        };

        //clean JSON response to SPListItem only
        return $http(config).then(function (response) {
            if (response.data.d.results) {
                return response.data.d.results[0];
            } else {
                return null;
            }
        });
    });
};

//JSON blob write to SharePoint list - needs $http factory, SharePoint list name, and JS object to stringify for save
spcrud.jsonWrite = function ($http, listName, jsonBody) {
    return spcrud.jsonRead($http, listName).then(function (item) {
        //HTTP 200 OK
        if (item) {
            //update if found
            item.JSON = JSON.stringify(jsonBody);
            return spcrud.update($http, listName, item.Id, item);
        } else {
            //create if missing
            var item = {
                '__metadata': {
                    'type': 'SP.ListItem'
                },
                'Title': spcrud.login,
                'JSON': JSON.stringify(jsonBody)
            };
            return spcrud.create($http, listName, item);
        }
    });
};
