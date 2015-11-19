/* Common EWS functionality */

var ews = (function () {
    "use strict";

    var ews = {};

    /* Models */
    ews.folder = function () {
        this.Id = null;
        this.ParentId = null;
        this.DisplayName = "";
        this.DisplayNameAllLower = "";
        this.TotalCount = 0;
        this.ChildFolderCount = 0;
        this.UnreadCount = 0;
    };

    /* Cache */
    ews.folders = [];

    /* Private Functions */
    var __getEWSEnvelope = function (request) {
        // Wrap an Exchange Web Services request in a SOAP envelope. 
        var result =

        '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '  <soap:Header>' +
        '  <t:RequestServerVersion Version="Exchange2013"/>' +
        '  </soap:Header>' +
        '  <soap:Body>' +

        request +

        '  </soap:Body>' +
        '</soap:Envelope>';

        return result;
    }

    /* Public Functions */
    ews.getFolderList = function (mailbox, callback) {
        // Return a GetItem EWS operation request for the list of folders  
        var call =
            '<FindFolder' +
            '   Traversal="Deep"' +
            '   xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '   xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '   <FolderShape>' +
            '       <t:BaseShape>AllProperties</t:BaseShape>' +
            '   </FolderShape>' +
            '   <ParentFolderIds>' +
            '       <t:DistinguishedFolderId Id="inbox"/>' +
            '   </ParentFolderIds>' +
            '</FindFolder>';
           
        mailbox.makeEwsRequestAsync(__getEWSEnvelope(call), function(asyncResult) {
            // convert from XML to plain js objects
            var results = jQuery.parseXML(asyncResult.value);
            var xmlDoc = $(results);
            var folderNodes = xmlDoc.find("Folder");
            var folderModel, folderNode;
            for (var i = 0; i < folderNodes.length; i++) {
                folderNode = $(folderNodes[i]);

                folderModel = new ews.folder();
                folderModel.Id = folderNode.find("FolderId").attr("Id");
                folderModel.ParentId = folderNode.find("ParentFolderId").attr("Id");
                folderModel.DisplayName = folderNode.find("DisplayName").text();
                folderModel.DisplayNameAllLower = folderModel.DisplayName.toLowerCase();
                folderModel.TotalCount = folderNode.find("TotalCount").text();
                folderModel.ChildFolderCount = folderNode.find("ChildFolderCount").text();
                folderModel.UnreadCount = folderNode.find("UnreadCount").text();

                ews.folders.push(folderModel)
            }
            callback(ews.folders);
        });
    }

    ews.GetItem = function (mailbox, item, callback) {
        var call = 
            '<GetItem' +
            '   xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '   xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '   <ItemShape>' +
            '       <t:BaseShape>Default</t:BaseShape>' +
            '       <t:IncludeMimeContent>true</t:IncludeMimeContent>' +
            '   </ItemShape>' +
            '   <ItemIds>' +
            '       <t:ItemId Id="' + item.itemId + '" />' +
            '   </ItemIds>' +
            '</GetItem>';

        mailbox.makeEwsRequestAsync(__getEWSEnvelope(call), function (asyncResult) {
            var item = $.parseXML(asyncResult.value);
            callback(item);
        });
    }

    ews.Move = function (mailbox, item, folderId) {
        var call = 
        '<MoveItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '    <ToFolderId>' +
        '        <t:FolderId Id="' + folderId + '" />' +
        '    </ToFolderId>' +
        '    <ItemIds>' +
        '        <t:ItemId Id="' + item.itemId + '" ChangeKey="' + item.changeKey + '"/>' +
        '    </ItemIds>' +
        '</MoveItem>';

        mailbox.makeEwsRequestAsync(__getEWSEnvelope(call), function (asyncResult) {
        });
    }

    return ews;
})();