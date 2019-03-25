var config;
var btnEvent;

// The initialize function must be run each time a new page is loaded
Office.initialize = function (reason) {};

var settingsDialog;

function ForwardToZohoDesk(event) {

    config = getConfig();

    // Check if the add-in has been configured
    if (config && config.zohodeskemail) {

        var originalSenderAddress = Office.context.mailbox.item.sender.emailAddress;
        var emailsubject = Office.context.mailbox.item.subject;

        Office.context.mailbox.item.body.getAsync(
            "html", {
                asyncContext: 'To Zoho Desk'
            },
            function callback(result) {
                var emailbody = result.value;
                Office.context.mailbox.displayNewMessageForm({
                    toRecipients: [config.zohodeskemail],
                    subject: emailsubject,
                    htmlBody: '#original_sender {' + originalSenderAddress + '} <br/><hr><br/>' + emailbody
                });
            });

    } else {
        // Save the event object so we can finish up later
        btnEvent = event;
        // Not configured yet, display settings dialog with
        // warn=1 to display warning.
        var url = new URI('../settings/dialog.html?warn=1').absoluteTo(window.location).toString();
        var dialogOptions = {
            width: 20,
            height: 40,
            displayInIframe: true
        };

        Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
            settingsDialog = result.value;
            settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, receiveMessage);
            settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, dialogClosed);
        });
    }
}

function ShowSettings(event) {

    config = getConfig();

    var url = new URI('../settings/dialog.html').absoluteTo(window.location).toString();
    if (config) {
        // If the add-in has already been configured, pass the existing values
        // to the dialog
        console.log("Send to email address: " + config.zohodeskemail);
        url = url + '?email=' + config.zohodeskemail;

    }

    var dialogOptions = {
        width: 20,
        height: 40,
        displayInIframe: true
    };

    Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
        settingsDialog = result.value;
        settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, receiveMessage);
        settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, dialogClosed);
    });

}

function receiveMessage(message) {
    config = JSON.parse(message.message);
    setConfig(config, function (result) {
        settingsDialog.close();
        settingsDialog = null;
        btnEvent.completed();
        btnEvent = null;
    });
}

function dialogClosed(message) {
    settingsDialog = null;
    btnEvent.completed();
    btnEvent = null;
}