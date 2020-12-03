'use strict';

(function () {
    // Test
    var contentUrl = "https://taptools.ngrok.io/";

    // Prod
    //var contentUrl = "https://provisioningbulkrequest.azurewebsites.net/"

    var internalUrl = contentUrl + "internal";
    var r3Url = contentUrl + "r3";


    microsoftTeams.initialize();

    function setValid() {
        console.log("onClick called");
        microsoftTeams.settings.setValidityState(true);
    }

    
    microsoftTeams.settings.registerOnSaveHandler(function (saveEvent) {
        console.log("calling registerOnSaveHandler");
        var selected = document.querySelector('.formSelect:checked').value;

        var url;
        if (selected == "Internal") {
            url = internalUrl;
        } else if (selected == "R3") {
            url = r3Url;
        } else {
            url = contentUrl;
        }

        var settings = {
            entityId: "I dunno",
            contentUrl: url,
            suggestedDisplayName: "Bulk Provisioning"
        }
        microsoftTeams.settings.setSettings(settings);
        saveEvent.notifySuccess();
    });

})();