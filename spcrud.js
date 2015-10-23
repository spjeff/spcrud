//library with AngularJS operations for CRUD to SharePoint 2013 over REST api

//namespace
var spcrud = spcrud || {};

//globals
spcrud.apiUrl = _spPageContextInfo.webAbsoluteUrl + '/_api/web/lists/getbytitle(\'{0}\')/items';
spcrud.json = 'application/json;odata=verbose';
spcrud.headers = {
    'Content-Type': spcrud.json,
    'Accept': spcrud.json,
    'X-RequestDigest': document.querySelector('#__REQUESTDIGEST').value
};

//CREATE item
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

//READ entire list
spcrud.read = function ($http, listName) {
    var config = {
        method: 'GET',
        url: spcrud.apiUrl.replace('{0}', listName),
        headers: spcrud.headers
    };
    return $http(config);
};

//READ single item
spcrud.readItem = function ($http, listName, id) {
    var config = {
        method: 'GET',
        url: spcrud.apiUrl.replace('{0}', listName) + '(' + id + ')',
        headers: spcrud.headers
    };
    return $http(config);
};

//UPDATE item
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

//DELETE item
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

//parse SharePoint list item Id
spcrud.getId = function (resp) {
    var item;
    if (resp.data.d) {
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
    }
    return (item) ? item.Id : null;
};
