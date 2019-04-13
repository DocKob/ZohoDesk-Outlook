var config;

// The initialize function must be run each time a new page is loaded
Office.initialize = function () {}

function Forward(event) {
    config = getConfig();
    if (config && config.zohodeskemail) {
        config = getConfig();

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
                event.completed();
            });
    } else {
        $('#app-body').show();
        event.completed();
    }
}

function getConfig() {
    var config = {};

    config.zohodeskemail = Office.context.roamingSettings.get('zohodesk-email');

    return config;
}