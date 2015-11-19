/// <reference path="../App.js" />
/// <reference path="../../Scripts/microsoft.ews.js" />
/// <reference path="../../Scripts/search.js" />

(function () {
    "use strict";

    var folderCache = null,
        mailbox,
        item,
        config =  {
            searchResultsLimit: 10
        };


    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            init();
        });
    };
    
    function init() {
        var loadingContainer = $("#loading");
        var mainContainer = $("main").hide();

        mailbox = Office.context.mailbox;
        item = Office.cast.item.toItemRead(mailbox.item);
        var user = mailbox.userProfile.displayName;
        var from = Office.cast.item.toMessageRead(item).from;

        ews.getFolderList(mailbox, function (folders) {
            folderCache = folders;
            search.initialize("#appBody", onSearch);
            $("#searchInput").focus();
            loadingContainer.hide();
            mainContainer.show();
        });

        $('#title').text("Where shall we file this email " + user.split(' ')[0] + "?");

    }

    function onSearch(data) {
        var keyword = data.keyword.toLowerCase();

        $("#keyword").text(keyword);

        if (data.action == "clear" || keyword.length == 0) {
            updateFolderOptions(null);
        }
        else if (data.action == "search") {
            var matchingFolders = [],
                        folder,
                        length,
                        i,
                        matchIndex,
                        matchCount;

            length = folderCache.length;

            for (i = 0; i < length; i++) {
                folder = folderCache[i];
                if ((matchIndex = folder.DisplayNameAllLower.indexOf(keyword)) != -1) {
                    folder.MatchIndex = matchIndex;
                    folder.DisplayNameHighlight =
                        folder.DisplayName.slice(0, matchIndex) +
                        "<strong>" +
                        folder.DisplayName.slice(matchIndex, matchIndex + keyword.length) +
                        "</strong>" +
                        folder.DisplayName.slice(matchIndex + keyword.length);

                    matchCount = matchingFolders.push(folder);
                }

            }

            // sort by nearest match (position of the start of the match)
            matchingFolders.sort(function (a, b) {
                if (a.MatchIndex < b.MatchIndex) {
                    return -1;
                }

                if (a.MatchIndex > b.MatchIndex) {
                    return 1;
                }

                return 0;
            });

            updateFolderOptions(matchingFolders);
        }
        else if (data.action == "select") {
            // was a specific index specific?
            if (data.index >= 0) {
                var item = $("#searchResults > ul > li")[data.index];
                if (item) {
                    $(item).click();
                }
            }
            else {
                // select current highlighted element (or the first one)
                var list = $("#searchResults > ul > li:active");
                if (list.length == 0) {
                    // if no active li, trigger off the first
                    $("#searchResults > ul > li").first().click();
                }
            }
        }
        else if (data.action == "navigate") {
            // navigate
        }


        
    }

    function updateFolderOptions(folders) {
        var containerElement = $("#searchResults");
        containerElement.empty();

        if (!folders) { return; }

        var list = $("<ul>").appendTo(containerElement);
        var length = folders.length;
        var i;

        var stop = length >= config.searchResultsLimit ? config.searchResultsLimit : length;
        for (i = 0; i < stop; i++) {
            list.append(
                $('<li data-folderid="' + folders[i].Id + '">' +
                  ' <button class="ms-Button ms-fontColor-neutralTertiary ms-font-xl">' +
                  '     <span class="ms-Button-label"> <span class="ms-borderColor-neutralTertiary ms-font-xs">' + (i + 1) + '</span> ' + folders[i].DisplayNameHighlight + '</span>' +
                  ' </button>' +
                  '</li>')
                .click(function (e) {
                    moveItem($(e.currentTarget).data("folderid"))
                })
            );
        }

        if (stop != length) {
            containerElement.append('<div class="info ms-font-mi"><i class="ms-Icon ms-Icon--filter" aria-hidden="true"></i>Showing top ' + config.searchResultsLimit + ' matches out of ' + length + ' possible.</div>');
        }
    }

    function moveItem(toFolderId) {
        ews.GetItem(mailbox, item, function (ewsItem) {
            item.changeKey = $(ewsItem).find("ItemId").attr("ChangeKey");
            ews.Move(mailbox, item, toFolderId);
        });
    }

})();