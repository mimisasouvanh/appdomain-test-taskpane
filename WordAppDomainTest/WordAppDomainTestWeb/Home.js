/// <reference path="/Scripts/FabricUI/MessageBanner.js" />


(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample provides links to test AppDomains.");
                $('#button-text').text("Navigate with window.location");
                $('#button-desc').text("Click to navigate within task pane");
                
                $('#highlight-button').click(
                    displaySelectedText);
                return;
            }

            $("#template-description").text("This sample provides links to test AppDomains.");
            $('#button-text').text("Navigate with window.location");
            $('#button-desc').text("Click to navigate within task pane");
                  

            // Add a click event handler for the highlight button.
            $('#dialog-open-window').click(
                navigatetoBing);
            // Add a click event handler for the dialog button.
            $('#dialog-open-bing').click(
                openDialogtoBing);
            $('#dialog-open-page').click(
                openDialogtoPage);
        });
    };
    

    function openDialogtoBing() {
        app.openDialog("https://www.bing.com", 50, 50);
    }
    function openDialogtoPage() {
        app.openDialog("https://localhost:44344/page.html", 50, 50);
    }
    function navigatetoBing() {
        window.location.replace("https://www.bing.com");
       // console.log(error);        
    } 


    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
