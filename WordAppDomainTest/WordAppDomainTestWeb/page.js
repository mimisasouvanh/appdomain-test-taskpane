/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information. */

(function () {
    "use strict";

    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#ok-button').click(messageParent);
            $('#bing').click(navigateBing);
            $('#open-bing').click(openNewWindow);
        });
    };

    function messageParent() {
        Office.context.ui.messageParent('ok');
    }

    function navigateBing() {
        window.location.replace("https://www.bing.com");
    }

    function openNewWindow() {
        var myWindow = window.open("https://www.bing.com");
    }

}());
