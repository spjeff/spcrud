//demo AngularJS controller to test run CRUD operations

function spcrudCtl($scope, $http) {
    //default data
    var vm = $scope;
    vm.status = 'OK';
    vm.listName = 'Test';
    vm.accTable = 'Employees';
    vm.accSubweb = '/Timesheet';

    //click events
    vm.create = function () {
        vm.status = ''
        spcrud.create($http, vm.listName, {
            'Title': 'Hello World'
        }).then(function (resp) {
            vm.itemId = resp.data.d.ID;
            vm.done('CREATE Id=' + vm.itemId);
        });
    };
    vm.read = function () {
        spcrud.read($http, vm.listName).then(function (resp) {
            vm.itemId = resp.data.d.results[0].ID;
            vm.done('READ Id=' + vm.itemId);
        });
    };
    vm.update = function () {
        spcrud.update($http, vm.listName, vm.itemId, {
            'Title': 'Hello Update'
        }).then(function () {
            vm.done('UPDATE Id=' + vm.itemId);
        });
    };
    vm.del = function () {
        spcrud.del($http, vm.listName, vm.itemId).then(function () {
            vm.itemId = null;
            vm.done('DELETE Id=' + vm.itemId);
        });
    };

    //status display
    vm.done = function (operation) {
        vm.status = operation + ' complete ' + (new Date()).toTimeString();
    };

    //==========
    //JSON blob read
    vm.settingsRead = function () {
        spcrud.jsonRead($http, 'Test').then(function (item) {
            if (item) {
                //success
                console.log('settingsRead');
                console.log(item);
                vm.settings = item.JSON;
            }
        });
    };

    //JSON blob write
    vm.settingsWrite = function () {
        spcrud.jsonWrite($http, 'Test', vm.settings).then(function (response) {
            //response
            console.log('settingsWrite');
            console.log(response);
        });
    };
    
    
    //==========
    //SendMail
    vm.send = function () {
      vm.mailResult = '';
      spcrud.sendMail($http, vm.mailTo, vm.mailFrom, vm.mailSubject, vm.mailBody).then(function(resp) {
          vm.mailResult = angular.toJson(resp);
      });
    };
    
    
    //==========
    //ACCDW
    //CREATE
    vm.accCreate = function() {
    	spcrud.setBaseUrl(_spPageContextInfo.webAbsoluteUrl + vm.accSubweb);
    	spcrud.refreshDigest($http).then(function() {
	    	spcrud.accCreate($http, vm.accTable, [null, "222", "john smith"], ["ID","Employee Number","First Name"]).then(function (response) {
	            //response
	            console.log('accCreate');
	            console.log(response);
	            vm.accResult = JSON.stringify(response.data.d.Result.Values);
   	            vm.accId = response.data.d.Result.Values[0][0];
	        });
        });
    };
    
    //READ
    vm.accRead = function() {
    	spcrud.setBaseUrl(_spPageContextInfo.webAbsoluteUrl + vm.accSubweb);
    	spcrud.refreshDigest($http).then(function() {
	    	spcrud.accRead($http, vm.accTable).then(function (response) {
	            //response
	            console.log('accRead');
	            console.log(response);
	            vm.accResult = JSON.stringify(response.data.d.Result.Values);
	            vm.accId = response.data.d.Result.Values[0][0];
	        });
        });
    };
    
    //READ
    vm.accReadID = function() {
    	spcrud.setBaseUrl(_spPageContextInfo.webAbsoluteUrl + vm.accSubweb);
    	spcrud.refreshDigest($http).then(function() {
	    	spcrud.accReadID($http, vm.accTable, vm.accId).then(function (response) {
	            //response
	            console.log('accReadID');
	            console.log(response);
	            vm.accResult = JSON.stringify(response.data.d.Result.Values);
	            vm.accId = response.data.d.Result.Values[0][0];
	        });
        });
    };


	//UPDATE
    vm.accUpdate = function() {
   		spcrud.setBaseUrl(_spPageContextInfo.webAbsoluteUrl + vm.accSubweb);
    	spcrud.refreshDigest($http).then(function() {
	    	spcrud.accUpdate($http, vm.accTable, [vm.accId, "111","john update"], ["ID","Employee Number","First Name"]).then(function (response) {
	            //response
	            console.log('accUpdate');
	            console.log(response);
	            vm.accResult = JSON.stringify(response.data.d.Result.Values);
	        });
        });
    };

	//DELETE
    vm.accDelete = function() {
   		spcrud.setBaseUrl(_spPageContextInfo.webAbsoluteUrl + vm.accSubweb);
    	spcrud.refreshDigest($http).then(function() {
	    	spcrud.accDelete($http, vm.accTable, vm.accId).then(function (response) {
	            //response
	            console.log('accDel');
	            console.log(response);
	            vm.accResult = JSON.stringify(response.data.d.Result.Values);
	        });
        });
    };
}

//load
angular.module('spcrudApp', []).controller('spcrudCtl', spcrudCtl);
