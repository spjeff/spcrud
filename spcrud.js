/**
 * Library with AngularJS operations for CRUD operations to SharePoint 2013 lists over REST api
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
 * version 0.1.15
 * last updated 10-18-2016
 *
 */

//namespace
'use strict';
var spcrud = spcrud || {};
var spsocial = spsocial || {};

//----------SHAREPOINT GENERAL----------

//initialize
spcrud.init = function() {
    //default to local web URL
    spcrud.apiUrl = spcrud.baseUrl + '/_api/web/lists/GetByTitle(\'{0}\')/items';

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

//string ends with
spcrud.endsWith = function(str, suffix) {
    return str.indexOf(suffix, str.length - suffix.length) !== -1;
};

//digest refresh worker
spcrud.refreshDigest = function($http) {
    var config = {
        method: 'POST',
        url: spcrud.baseUrl + '/_api/contextinfo',
        headers: spcrud.headers
    };
    return $http(config).then(function(response) {
        //parse JSON and save
        spcrud.headers['X-RequestDigest'] = response.data.d.GetContextWebInformation.FormDigestValue;
    });

};

//send email
spcrud.sendMail = function($http, to, ffrom, subj, body) {
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
            method: 'POST',
            url: spcrud.baseUrl + '/_api/SP.Utilities.Utility.SendEmail',
            headers: spcrud.headers,
            data: angular.toJson(message)
        };
    return $http(config);
};

//----------SHAREPOINT USER PROFILES----------

//lookup SharePoint current web user
spcrud.getCurrentUser = function($http) {
    if (!spcrud.currentUser) {
        var url = spcrud.baseUrl + '/_api/web/currentuser?$expand=Groups',
            config = {
                method: 'GET',
                url: url,
                cache: true,
                headers: spcrud.headers
            };
        return $http(config);
    } else {
        return {
            then: function(a) {
                a();
            }
        };
    }
};

//lookup my SharePoint profile
spcrud.getMyProfile = function($http) {
    if (!spcrud.myProfile) {
        var url = spcrud.baseUrl + '/_api/SP.UserProfiles.PeopleManager/GetMyProperties?select=*',
            config = {
                method: 'GET',
                url: url,
                cache: true,
                headers: spcrud.headers
            };
        return $http(config);
    } else {
        return {
            then: function(a) {
                a();
            }
        };
    }
};

//lookup any SharePoint profile
spcrud.getProfile = function($http, login) {
    var url = spcrud.baseUrl + '/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v=\'' + login + '\'&select=*',
        config = {
            method: 'GET',
            url: url,
            headers: spcrud.headers
        };
    return $http(config);
};

//ensure SPUser exists in target web
spcrud.ensureUser = function($http, login) {
    var url = spcrud.baseUrl + '/_api/web/ensureuser',
        config = {
            method: 'POST',
            url: url,
            headers: spcrud.headers,
            data: login
        };
    return $http(config);
};

//----------SHAREPOINT FILES AND FOLDERS----------

//create folder
spcrud.createFolder = function($http, folderUrl) {
    var data = {
            '__metadata': {
                'type': 'SP.Folder'
            },
            'ServerRelativeUrl': folderUrl
        },
        url = spcrud.baseUrl + '/_api/web/folders',
        config = {
            method: 'POST',
            url: url,
            headers: spcrud.headers,
            data: data
        };
    return $http(config);
};

// upload file to folder
// https://kushanlahiru.wordpress.com/2016/05/14/file-attach-to-sharepoint-2013-list-custom-using-angular-js-via-rest-api/
// http://stackoverflow.com/questions/17063000/ng-model-for-input-type-file
spcrud.uploadFile = function($http, folderUrl, fileName, binary) {
    var url = spcrud.baseUrl + '/_api/web/GetFolderByServerRelativeUrl(\'' + folderUrl + '\')/files/add(overwrite=true, url=\'' + fileName + '\')',
        config = {
            method: 'POST',
            url: url,
            headers: spcrud.headers,
            data: binary,
            transformRequest: []
        };
    return $http(config);
};

//upload attachment to item
spcrud.uploadAttach = function($http, listName, id, fileName, binary, overwrite) {
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
        method: 'POST',
        url: url,
        headers: headers,
        data: binary
    };
    return $http(config);
};

//get attachment for item
spcrud.getAttach = function($http, listName, id) {
    var url = spcrud.baseUrl + '/_api/web/lists/GetByTitle(\'' + listName + '\')/items(' + id + ')/AttachmentFiles',
        config = {
            method: 'GET',
            url: url,
            headers: spcrud.headers
        };
    return $http(config);
};

//----------SHAREPOINT LIST CORE----------

//CREATE item - SharePoint list name, and JS object to stringify for save
spcrud.create = function($http, listName, jsonBody) {
    //append metadata
    if (!jsonBody.__metadata) {
        jsonBody.__metadata = {
            'type': 'SP.ListItem'
        };
    }
    var data = angular.toJson(jsonBody),
        config = {
            method: 'POST',
            url: spcrud.apiUrl.replace('{0}', listName),
            data: data,
            headers: spcrud.headers
        };
    return $http(config);
};

//READ entire list - needs $http factory and SharePoint list name
spcrud.read = function($http, listName, options) {
    //build URL syntax
    //https://msdn.microsoft.com/en-us/library/office/fp142385.aspx#bk_support
    var url = spcrud.apiUrl.replace('{0}', listName);
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

    //config
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//READ single item - SharePoint list name, and item ID number
spcrud.readItem = function($http, listName, id) {
    var config = {
        method: 'GET',
        url: spcrud.apiUrl.replace('{0}', listName) + '(' + id + ')',
        headers: spcrud.headers
    };
    return $http(config);
};

//UPDATE item - SharePoint list name, item ID number, and JS object to stringify for save
spcrud.update = function($http, listName, id, jsonBody) {
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
    var data = angular.toJson(jsonBody),
        config = {
            method: 'POST',
            url: spcrud.apiUrl.replace('{0}', listName) + '(' + id + ')',
            data: data,
            headers: headers
        };
    return $http(config);
};

//DELETE item - SharePoint list name and item ID number
spcrud.del = function($http, listName, id) {
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

//JSON blob read from SharePoint list - SharePoint list name
spcrud.jsonRead = function($http, listName, cache) {
    return spcrud.getCurrentUser($http).then(function(response) {
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
        spcrud.getMyProfile($http).then(function(response) {
            spcrud.myProfile = response.data.d;
        });

        //parse single SPListItem only
        return $http(config).then(function(response) {
            if (response.data.d.results) {
                return response.data.d.results[0];
            } else {
                return null;
            }
        });
    });
};

//JSON blob upsert write to SharePoint list - SharePoint list name and JS object to stringify for save
spcrud.jsonWrite = function($http, listName, jsonBody) {
    return spcrud.refreshDigest($http).then(function(response) {
        return spcrud.jsonRead($http, listName).then(function(item) {
            //HTTP 200 OK
            if (item) {
                //update if found
                item.JSON = angular.toJson(jsonBody);
                return spcrud.update($http, listName, item.Id, item);
            } else {
                //create if missing
                item = {
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

//----------SHAREPOINT SOCIAL NEWSFEED----------

//Returns newsfeed thread info
spsocial.getNewsFeed = function($http, utc) {
    var url = "";
    if (utc) {
        url = spcrud.baseUrl + '/_api/social.feed/my/news(MaxThreadCount=100,OlderThan=@v)?@v=datetime' + "%27" + utc + "%27";
    } else {
        url = spcrud.baseUrl + '/_api/social.feed/my/news';
    }
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Gets the feed of activity by the current user and by people and content the user is following, sorted by created date
spsocial.getTimeLineFeed = function($http, olderThan) {
    var url = "";
    if (olderThan) {
        url = spcrud.baseUrl + "/_api/social.feed/my/timelinefeed(MaxThreadCount=25,SortOrder=0,NewerThan=@v)?@v=datetime'2016-03-11T21:48:45.000Z'";
    } else {
        url = spcrud.baseUrl + '/_api/social.feed/my/timelinefeed';
    }

    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Get feed of personal activity
spsocial.getPersonalActivityFeed = function($http) {
    var url = spcrud.baseUrl + '/_api/social.feed/my/feed';
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Get feed of like activity
spsocial.getLikesFeed = function($http) {
    var url = spcrud.baseUrl + '/_api/social.feed/my/likes';
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Get list of followers
spsocial.getFollowers = function($http) {
    var url = spcrud.baseUrl + '/_api/social.following/my/followed(types=1)';
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Gets the feed of activity by the current user and by people and content the user is following, sorted by created date
spsocial.getMentionsFeed = function($http) {
    var url = spcrud.baseUrl + '/_api/social.feed/my/mentionfeed';
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Get all replies for a given post ID
spsocial.getAllReplies = function($http, id) {
    var data = {
        ID: id
    };
    var url = spcrud.baseUrl + "/_api/social.feed/post";
    var config = {
        method: 'POST',
        data: data,
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Social data counts
//1 - Users
//2 - Documents
//4 - Sites
//8 - Tags
//value enum: https://msdn.microsoft.com/en-us/library/office/dn194080.aspx#bk_FollowedCount​
spsocial.getSocialCounts = function($http, types) {
    var url = spcrud.baseUrl + '/_api/social.following/my/followedcount(types=' + types + ')';
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Returns sites being followed by the user
spsocial.getFollowedSites = function($http) {
    var url = spcrud.baseUrl + '/_api/social.following/my/followed(types=4)';
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Return Trending Tags Data
spsocial.getTrendingTags = function($http, utc) {
    var url = spcrud.baseUrl + "/_api/search/query?querytext='ContentTypeId:0x01FD* write>=\"" + utc + "\" -ContentClass=urn:content-class:SPSPeople'&refiners='Tags'";
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Get all social tags used
spsocial.getAllTags = function($http) {
    var url = spcrud.baseUrl + "/_api/search/query?querytext='ContentTypeId:0x01FD* -ContentClass=urn:content-class:SPSPeople'&refiners='Tags'&rowlimit=500";
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Get list of recent documents
spsocial.getRecentDocs = function($http) {
    var url = spcrud.baseUrl + "/_api/search/query?querytext='*'&querytemplate='(AuthorOwsUser:{User.AccountName} OR EditorOwsUser:{User.AccountName}) AND ContentType:Document AND IsDocument:1 AND -Title:OneNote_DeletedPages AND -Title:OneNote_RecycleBin NOT(FileExtension:mht OR FileExtension:aspx OR FileExtension:html OR FileExtension:htm OR FileExtension:one OR FileExtension:bin)'&rowlimit=100&bypassresulttypes=false&selectproperties='Title,Path,Filename,FileExtension,Created,Author,LastModifiedTime,ModifiedBy,LinkingUrl,SiteTitle,ParentLink,DocumentPreviewMetadata,ListID,ListItemID,SPSiteURL,SiteID,WebId,UniqueID,SPWebUrl'&sortlist='LastModifiedTime:descending'&enablesorting=true";
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Follow another user
spsocial.follow = function($http, account) {
    var data = {
        'actor': {
            '__metadata': { "type": "SP.Social.SocialActorInfo" },
            'ActorType': 0,
            'AccountName': account,
            'Id': null
        }
    };
    var url = spcrud.baseUrl + "/_api/social.following/follow";
    var config = {
        method: 'POST',
        data: data,
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Unfollow another user
spsocial.unfollow = function($http, account) {
    var data = {
        'actor': {
            '__metadata': { "type": "SP.Social.SocialActorInfo" },
            'ActorType': 0,
            'AccountName': account,
            'Id': null
        }
    };
    var url = spcrud.baseUrl + "/_api/social.following/stopfollowing";
    var config = {
        method: 'POST',
        data: data,
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Get list of community sites
spsocial.getCommunitySites = function($http) {
    var url = spcrud.baseUrl + "/_api/search/query?querytext='WebTemplate=COMMUNITY OR WebTemplate=STS OR WebTemplate=PROJECTSITE'&rowlimit=1000&sortlist='LastModifiedTime:descending'&trimduplicates=false";
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};