//namespace
'use strict';
var spcrud = spcrud || {};

//----------ACCESS WEB DATABASE----------

//helper functions
spcrud.accInit = function () {
	spcrud.accUrl = spcrud.baseUrl + '/_vti_bin/accsvc/accessportal.json/GetData';
}
spcrud.accBody = function (table, isRead) {
    if (isRead) {
        //READ
        return {
            "dataBaseInfo": {
                "AllowAdditions": true,
                "AllowDeletions": true,
                "AllowEdits": true,
                "DataEntry": false,
                "DoNotPrefetchImages": false,
                "InitialPage": "0",
                "SelectCommand": table,
                "FetchSchema": true,
                "NewImageStorage": true
            },
            "pagingInfo": {
                "FirstRow": 0,
                "PageSize": 50,
                "RetrieveExactRowCount": true,
                "SortExpression": null,
                "UseCache": false,
                "SessionId": null
            }
        };
    } else {
        //WRITE
        return {
            "dataBaseInfo": {
                "AllowAdditions": true,
                "AllowDeletions": true,
                "AllowEdits": true,
                "DataEntry": false,
                "DoNotPrefetchImages": false,
                "InitialPage": "0",
                "SelectCommand": table,
                "FetchSchema": true,
                "NewImageStorage": true
            },
            "updateRecord": {
                "Paging": {
                    "FirstRow": 0,
                    "PageSize": 50,
                    "RetrieveExactRowCount": true,
                    "UseCache": true,
                    "SessionId": null,
                    "CacheCommands": 0,
                    "Filter": null,
                    "RowKey": 0,
                    "TotalRows": 0
                },
                "ReturnDataMacroIds": false
            }
        };
    }
};
spcrud.accWorker = function ($http, data, stem) {
        var url = (stem) ? spcrud.accUrl.replace('GetData', stem) : spcrud.accUrl;
        var config = {
            method: 'POST',
            url: url,
            headers: spcrud.headers,
            data: data
        };
        return $http(config);
    }
    // CREATE row - SQL Azure table name, values[], and fields[] arrays
spcrud.accCreate = function ($http, table, values, fields) {
    var data = spcrud.accBody(table);
    data.updateRecord.NewValues = [values];
    data.dataBaseInfo.FieldNames = fields;
    return spcrud.accWorker($http, data, 'InsertRecords');
};
// READ all - SQL Azure table name
spcrud.accRead = function ($http, table) {
    var data = spcrud.accBody(table, true);
    return spcrud.accWorker($http, data);
};
// READ row - SQL Azure table name and ID#
spcrud.accReadID = function ($http, table, id) {
    var data = spcrud.accBody(table, true);
    data.dataBaseInfo.Restriction = "<Expression xmlns='http://schemas.microsoft.com/office/accessservices/2010/12/application'><FunctionCall Name='='><Identifier Name='ID' Index= '0' /><StringLiteral Value='" + id + "' Index='1' /></FunctionCall></Expression>";
    return spcrud.accWorker($http, data);
};
// UPDATE row - SQL Azure table name, values[], and fields[] arrays
spcrud.accUpdate = function ($http, table, values, fields) {
    var data = spcrud.accBody(table);
    data.updateRecord.OriginalValues = [values];
    data.updateRecord.NewValues = [values];
    data.dataBaseInfo.FieldNames = fields;
    return spcrud.accWorker($http, data, 'UpdateRecords');
};
// DELETE row - SQL Azure table name and ID#
spcrud.accDelete = function ($http, table, id) {
    return spcrud.accReadID($http, table, id).then(function (response) {
        var data = spcrud.accBody(table);
        data.dataBaseInfo.FieldNames = [];
        angular.forEach(response.data.d.Result.Fields, function (f) {
            data.dataBaseInfo.FieldNames.push(f.ColumnName);
        });
        data.updateRecord.OriginalValues = response.data.d.Result.Values;
        return spcrud.accWorker($http, data, 'DeleteRecords');
    });
};