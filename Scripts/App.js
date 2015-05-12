﻿(function ($, SP, window) {
    function getQueryStringParameters() {
        var params = document.URL.split("?")[1].split("&");
        var obj = {};

        for (var i = 0; i < params.length; i = i + 1) {
            var singleParam = params[i].split("=");
            obj[singleParam[0]] = decodeURIComponent(singleParam[1]);
        }

        return obj;
    }

    function getFileServerRelativeUrl() {
        var deferred = $.Deferred();

        var queryStringParameters = getQueryStringParameters();

        if (!queryStringParameters.SPListId && !queryStringParameters.SPListItemId) {
            return deferred.reject('This app is a custom menu item action, please visit this page from menu item.');
        }

        var appWebUrl = queryStringParameters.SPAppWebUrl;
        var hostWebUrl = queryStringParameters.SPHostUrl;

        var clientContext = SP.ClientContext.get_current();
        var appContextSite = new SP.AppContextSite(clientContext, hostWebUrl);
        var web = appContextSite.get_web();
        var list = web.get_lists().getById(queryStringParameters.SPListId);
        var listItem = list.getItemById(queryStringParameters.SPListItemId);
        var file = listItem.get_file();

        clientContext.load(file, 'ServerRelativeUrl');
        clientContext.executeQueryAsync(function (sender, args) {
            var serverRelativeUrl = file.get_serverRelativeUrl();

            deferred.resolve(serverRelativeUrl, appWebUrl, hostWebUrl);
        }, function (sender, args) {
            var message = args.get_message();

            deferred.reject(message);
        });

        return deferred.promise();
    }

    function getFileServerRelativeUrlOnFail(message) {
        $('.spinner').hide();
        $('.error').text(message).show();
    }

    function getFileExtension(fileServerRelativeUrl) {
        if (!fileServerRelativeUrl) {
            return '';
        }

        var matches = fileServerRelativeUrl.match(/\.(\w+)$/);

        if (matches.length > 1) {
            return matches[1];
        } else {
            return '';
        }
    }

    function isSupportedFileExtension(fileExtension) {
        return /txt|md|xml|js|css|html/i.test(fileExtension);
    }

    function readFileContents(fileServerRelativeUrl, appWebUrl, hostWebUrl) {
        var deferred = $.Deferred();

        var fileExtension = getFileExtension(fileServerRelativeUrl);

        if (!isSupportedFileExtension(fileExtension)) {
            return deferred.reject('Unsupported file extension.');
        }

        var executor = new SP.RequestExecutor(appWebUrl);
        var options = {
            url: appWebUrl + "/_api/SP.AppContextSite(@target)/web/GetFileByServerRelativeUrl('" + fileServerRelativeUrl + "')/$value?@target='" + hostWebUrl + "'",
            type: "GET",
            success: function (response) {
                if (response.statusCode == 200) {
                    alert(response.body);
                } else {
                    deferred.reject(response.statusCode + ": " + response.statusText);
                }
            },
            error: function (response) {
                deferred.reject(response.statusCode + ": " + response.statusText);
            }
        };

        executor.executeAsync(options);

        return deferred.promise();
    }

    function readFileContentsOnFail(message) {
        $('.spinner').hide();
        $('.error').text(message).show();
    }

    function render(fileContents) {

    }

    var App = {
        run: function () {
            getFileServerRelativeUrl().then(readFileContents, getFileServerRelativeUrlOnFail).then(render, readFileContentsOnFail);
        }
    };

    window.App = App;
})(jQuery, SP, window);