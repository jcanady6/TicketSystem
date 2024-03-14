(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var appId = "eb06e508-6e7f-463f-b35d-aa3855ac6981";
            var item = Office.context.mailbox.item;
            var bodyText = "";
            var body = item.body;

            body.getAsync(Office.CoercionType.Text, function (asyncResult) {
                if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                }
                else {
                    bodyText = asyncResult.value;
                }
                var parameters = "&subject=" + item.subject + "&bodyText=" + bodyText;
                $('#canvas-iframe').attr("src", "https://apps.powerapps.com/play/" + appId + "?source=iframe" + parameters);
            });
        });
    }
})();