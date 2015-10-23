//demo AngularJS controller to test run CRUD operations

function spcrudCtl($scope, $http) {
    //default data
    var vm = $scope;
    vm.status = 'OK';
    vm.listName = 'Test';

    //click events
    vm.create = function () {
        vm.status = ''
        spcrud.create($http, vm.listName, {
            'Title': 'Hello World'
        }).then(function (resp) {
            vm.itemId = spcrud.getId(resp);
            vm.done('CREATE Id=' + vm.itemId);
        });
    };
    vm.read = function () {
        spcrud.read($http, vm.listName).then(function (resp) {
            vm.itemId = spcrud.getId(resp);
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
}

//load
angular.module('spcrudApp', []).controller('spcrudCtl', spcrudCtl);
