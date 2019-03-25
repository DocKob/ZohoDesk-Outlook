function getConfig() {
    var config = {};

    config.zohodeskemail = Office.context.roamingSettings.get('zohodeskemail');

    return config;
}

function setConfig(config, callback) {
    Office.context.roamingSettings.set('zohodeskemail', config.zohodeskemail);

    Office.context.roamingSettings.saveAsync(callback);
}