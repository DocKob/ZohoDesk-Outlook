var config;

// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {

};

function Forward() {
    config = getConfig();
    if (config && config.zohodeskemail) {
        console.log('OUI');
        ForwardToDesk();
    } else {
        console.log('NON');
        $('#app-body').show();
    }
}

function ForwardToDesk() {
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
        });
}

function getConfig() {
    var config = {};

    config.zohodeskemail = Office.context.roamingSettings.get('zohodesk-email');

    return config;
}