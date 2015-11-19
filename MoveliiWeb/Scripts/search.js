/* Common EWS functionality */

var search = (function () {
    "use strict";

    var search = {};
    var __lastSearch = null;

    search.currentSearch = "";

    search.clearAfter = 1500;

    search.initialize = function (elementSelector, callback) {
        $(elementSelector).keydown(function (e) {
            var meta = {};
            var now = new Date();

            var keyCode = e.which || e.keyCode;

            if (keyCode == 8 || keyCode == 46) { // BACKSPACE or DELETE
                search.currentSearch = search.currentSearch.substring(0, search.currentSearch.length - 1);
                meta.action = "search";
            }
            else if (keyCode == 27) { //ESC
                search.currentSearch = "";
                meta.action = "clear";
            }
            else if (keyCode == 37 || keyCode == 39 || keyCode == 38 || keyCode == 40) { // LEFT, RIGHT, UP, DOWN
                meta.action = "navigate";
                meta.direction = (keyCode == 37 || keyCode == 38) ? "-1" : 1;

            }
            else if (keyCode == 13) { // ENTER
                meta.action = "select";
            }
            else if (keyCode >= 48 && keyCode <= 57) { // 0-9
                meta.action = "select";
                meta.index = (keyCode - 48) - 1;
            }
            else if ((keyCode >= 65 && keyCode <= 90) || keyCode == 32) { // a-z and SPACE
                if (__lastSearch == null || (now.getTime() - __lastSearch.getTime() >= search.clearAfter)) {
                    search.currentSearch = "";
                }

                meta.action = "search";
                search.currentSearch += e.char;
            }
            else {
                return;
            }

            e.preventDefault();
            e.stopPropagation();


            __lastSearch = now;

            meta.keyword = search.currentSearch;
            if (callback) {
                callback(meta);
            }
        });
    }
    
    return search;
})();