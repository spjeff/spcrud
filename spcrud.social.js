/**
 * Library with AngularJS operations for SharePoint 2013 Newsfeed operations over REST api
 *
 *
 * @spjeff
 * spjeff@spjeff.com
 * http://spjeff.com
 *
 * version 0.1.20
 * last updated 06-06-2017
 *
 */

//namespace
'use strict';
var spsocial = spsocial || {};

//----------SHAREPOINT SOCIAL NEWSFEED----------

//Returns newsfeed thread info
spsocial.getNewsFeed = function ($http, utc) {
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
spsocial.getTimeLineFeed = function ($http, olderThan) {
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
spsocial.getPersonalActivityFeed = function ($http) {
    var url = spcrud.baseUrl + '/_api/social.feed/my/feed';
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Get feed of like activity
spsocial.getLikesFeed = function ($http) {
    var url = spcrud.baseUrl + '/_api/social.feed/my/likes';
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Get list of followers
spsocial.getFollowers = function ($http) {
    var url = spcrud.baseUrl + '/_api/social.following/my/followed(types=1)';
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Gets the feed of activity by the current user and by people and content the user is following, sorted by created date
spsocial.getMentionsFeed = function ($http) {
    var url = spcrud.baseUrl + '/_api/social.feed/my/mentionfeed';
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Get all replies for a given post ID
spsocial.getAllReplies = function ($http, id) {
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
spsocial.getSocialCounts = function ($http, types) {
    var url = spcrud.baseUrl + '/_api/social.following/my/followedcount(types=' + types + ')';
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Returns sites being followed by the user
spsocial.getFollowedSites = function ($http) {
    var url = spcrud.baseUrl + '/_api/social.following/my/followed(types=4)';
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Return Trending Tags Data
spsocial.getTrendingTags = function ($http, utc) {
    var url = spcrud.baseUrl + "/_api/search/query?querytext='ContentTypeId:0x01FD* write>=\"" + utc + "\" -ContentClass=urn:content-class:SPSPeople'&refiners='Tags'";
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Get all social tags used
spsocial.getAllTags = function ($http) {
    var url = spcrud.baseUrl + "/_api/search/query?querytext='ContentTypeId:0x01FD* -ContentClass=urn:content-class:SPSPeople'&refiners='Tags'&rowlimit=500";
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Get list of recent documents
spsocial.getRecentDocs = function ($http) {
    var url = spcrud.baseUrl + "/_api/search/query?querytext='*'&querytemplate='(AuthorOwsUser:{User.AccountName} OR EditorOwsUser:{User.AccountName}) AND ContentType:Document AND IsDocument:1 AND -Title:OneNote_DeletedPages AND -Title:OneNote_RecycleBin NOT(FileExtension:mht OR FileExtension:aspx OR FileExtension:html OR FileExtension:htm OR FileExtension:one OR FileExtension:bin)'&rowlimit=100&bypassresulttypes=false&selectproperties='Title,Path,Filename,FileExtension,Created,Author,LastModifiedTime,ModifiedBy,LinkingUrl,SiteTitle,ParentLink,DocumentPreviewMetadata,ListID,ListItemID,SPSiteURL,SiteID,WebId,UniqueID,SPWebUrl'&sortlist='LastModifiedTime:descending'&enablesorting=true";
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};

//Follow another user
spsocial.follow = function ($http, account) {
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
spsocial.unfollow = function ($http, account) {
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
spsocial.getCommunitySites = function ($http) {
    var url = spcrud.baseUrl + "/_api/search/query?querytext='WebTemplate=COMMUNITY OR WebTemplate=STS OR WebTemplate=PROJECTSITE'&rowlimit=1000&sortlist='LastModifiedTime:descending'&trimduplicates=false";
    var config = {
        method: 'GET',
        url: url,
        headers: spcrud.headers
    };
    return $http(config);
};