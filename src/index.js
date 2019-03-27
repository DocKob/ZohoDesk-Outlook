var config;

$(document).ready(() => {
    $('#desksettings').click(desksettings);
});

// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    config = getConfig();
    if (config && config.zohodeskemail) {
        var user = config.zohodeskemail;
        $('#zoho-email').val(user);
    } else {}
    $('#app-body').show();
};

async function desksettings() {
    config = getConfig();
    config.zohodeskemail = $('#zoho-email').val();
    Office.context.roamingSettings.set('zoho-email', config.zohodeskemail);
    Office.context.roamingSettings.saveAsync();
    ForwardToDesk();
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

    config.zohodeskemail = Office.context.roamingSettings.get('zoho-email');

    return config;
}