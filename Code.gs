/*
     ## Gmail Auto Forwarder ##

     Version  :  2
     Author   :  Shan Eapen Koshy
     Date.    :  4th November 2017
     Github   :  http://github.com/shaneapen/Gmail-Auto-Forwarder
     Website  :  https://codegena.com

*/

function onInstall() {
    onOpen();
}

/* What should the add-on do when a document is opened */
function onOpen() {
    SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Create New Rule', 'create_new_rule')
        .addItem('Manage Rules', 'manage_rules')
        .addSeparator()
        .addItem('View Remaining Quota', 'view_remaining_quota')
        .addItem('Quick Start Guide', 'showDialog')
        .addSeparator()
        .addItem('Reset', 'deleteAll')

        .addToUi();
}

function createTimeDrivenTriggers() {
    ScriptApp.newTrigger('runAllRules')
        .timeBased()
        .everyHours(3)
        .create();
}

/* Function that runs when create new rule is clicked */
function create_new_rule() {
    var html = HtmlService.createHtmlOutputFromFile('newRule_Dialog')
        .setWidth(600)
        .setHeight(400);
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .showModalDialog(html, 'Create New Rule');
}

/* Function that runs when manage rule is clicked */
function manage_rules() {
    var html = HtmlService.createHtmlOutputFromFile('manageRules_Dialog')
        .setWidth(500)
        .setHeight(300);
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .showModalDialog(html, 'Manage Forwarding Rules');
}

function view_remaining_quota() {
    var html = HtmlService.createHtmlOutputFromFile('remainingMailQuota')
        .setWidth(500)
        .setHeight(300);
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .showModalDialog(html, 'Remaining Mail Quota');
}

// Assistant function that passes list of labels and alias emails to the new_rule html page
function getData() {

    //Getting the user labels
    var labels = GmailApp.getUserLabels();
    for (i = 0; i < labels.length; ++i) {
        labels[i] = labels[i].getName()
    }

    //Getting current email + available alias email addresses
    var aliasEmails = GmailApp.getAliases();
    aliasEmails.push(Session.getActiveUser().getEmail())

    var data = {
        "labels": labels,
        "emailAddresses": aliasEmails
    }
    return (data);
}

/*  Function that gets called when user clicks on the create new rule button from the form */
function addRule(formInput) {
    var userProperties = PropertiesService.getUserProperties();
    var count;
    if (userProperties.getProperty('count') == null) {
        userProperties.setProperty('count', '0');
    }
    //getting count again- IMP
    var count = parseInt(userProperties.getProperty('count'));
    count += 1;
    userProperties.setProperty('count', count.toString());

    userProperties.setProperty('rule_' + userProperties.getProperty('count'), JSON.stringify(formInput));

    if (count == 1) {
        startForwarding();
    }
}

// Assistant function used to pass currently stored rules to the manage_rules html page
function getRules() {
    Logger.log(PropertiesService.getUserProperties().getProperties());
    var count = PropertiesService.getUserProperties().getProperty('count');
    var rulesList = []
    for (var i = 1; i <= count; ++i) {
        var rule = PropertiesService.getUserProperties().getProperty('rule_' + i);
        rule = JSON.parse(rule);

        var toEmail = 'forwarded to:' + rule.forwardTo;
        var string = search_filter_string(i) + ' ' + toEmail;
        Logger.log("string#### " + string);
        rulesList.push({
            'rule': i,
            'str': string
        });
    }
    return rulesList;

}

function deleteAll() {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Delete all rules?', ui.ButtonSet.YES_NO);

    // Process the user's response.
    if (response == ui.Button.YES) {
        Logger.log('The user clicked "Yes."');
        PropertiesService.getUserProperties().deleteAllProperties();
        PropertiesService.getUserProperties().setProperty('count', '0');
    } else {
        Logger.log('The user clicked "No" or the close button in the dialog\'s title bar.');
    }

}

function deleteRule(ruleNum) {
    /*  Case : Suppose the rules contains [rule1,rule2,rule3] and the user opts to delete rule1,
               then the whole array has to be shifted by 1 index to right and count value should be count-1
    */
    var userProperties = PropertiesService.getUserProperties();
    var count = parseInt(userProperties.getProperty('count'));

    //shifting the array to right
    for (var i = ruleNum; i <= count - 1; ++i) {
        userProperties.setProperty('rule_' + i, userProperties.getProperty('rule_' + (i + 1)));
    }

    // after deleting rule1 from [rule1,rule2,rule3] the rule2 and rule3 becomes same => delete last rule after each shift
    userProperties.deleteProperty('rule_' + count);

    if (count == 1) {
        userProperties.setProperty('count', '0');
        userProperties.deleteProperty('dateAsEpoch');
        deleteTriggers();
    } else {
        userProperties.setProperty('count', (count - 1).toString());
    }

}

function startForwarding() {
    var dateAsEpoch = Number(new Date().getTime() / 1000.0).toFixed(0);
    PropertiesService.getUserProperties().setProperty("dateAsEpoch", dateAsEpoch.toString());
    createTimeDrivenTriggers();
}

function deleteTriggers() {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() == "runAllRules") {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }
}

function runAllRules() {

    if (MailApp.getRemainingDailyQuota() > 1) {
        var lastCheckedDate = PropertiesService.getScriptProperties().getProperty("dateAsEpoch");
        var count = parseInt(PropertiesService.getUserProperties().getProperty('count'));
        //loops through all the rules
        for (var i = 1; i <= count; ++i) {
            var rule = JSON.parse(PropertiesService.getUserProperties().getProperty('rule_' + i));
            var threads = GmailApp.search('after:' + lastCheckedDate + ' ' + search_filter_string(i));

            for (i = 0; i < threads.length; ++i) {
                var message = GmailApp.getMessagesForThread(threads[i]);
                for (j = 0; j < message.length; ++j) {
                    var messageDate = Number(message[j].getDate().getTime() / 1000.0).toFixed(0);
                    if (messageDate > lastCheckedDate && rule.aliasEmail) {
                      message[j].forward(rule.forwardTo,{'from':rule.aliasEmail});
                    } else if(messageDate > lastCheckedDate){
                        message[j].forward(rule.forwardTo);
                    }else {
                        break; //break out of loop if current thread doesn't contain anymore new message
                    }
                }
            }
        }
       //Updating the lasChecked status
       var currentDate = new Date();
       var currentDateAsEpoch = Number(currentDate.getTime()/1000.0).toFixed(0);
       PropertiesService.getScriptProperties().setProperty("dateAsEpoch", currentDateAsEpoch);

    }

}


function runRule(ruleNum) {
    if (MailApp.getRemainingDailyQuota() > 1) {
        var lastCheckedDate = PropertiesService.getScriptProperties().getProperty("dateAsEpoch");
        var rule = JSON.parse(PropertiesService.getUserProperties().getProperty('rule_' + ruleNum));
        var threads = GmailsApp.search('after:' + lastCheckedDate + ' ' + search_filter_string(ruleNum));
        for (i = 0; i < threads.length; ++i) {
            var message = GmailApp.getMessagesForThread(threads[i]);
            for (j = 0; j < message.length; ++j) {
                var messageDate = Number(message[j].getDate().getTime() / 1000.0).toFixed(0);
                if (messageDate > lastCheckedDate) {
                    message[j].forward(rule.forwardTo);
                } else {
                    break; //break out of loop if current thread doesn't contain anymore new message
                }
            }
        }
    }
}


/*
        ## Assistant functions ##

        1. search_filter_string(ruleNumber) - > String
            Returns a string of search parameters (label + from + subject + advancedQueries)



*/

function search_filter_string(i) {
    var rule = JSON.parse(PropertiesService.getUserProperties().getProperty('rule_' + i));

    var label = (rule.label) ? 'label:' + rule.label + ' ' : '';
    var from = (rule.from) ? 'from:' + rule.from + ' ' : '';
    var subject = (rule.subject) ? 'subject:' + rule.subject + ' ' : '';
    var advancedSearch = (rule.advancedSearch) ? rule.advancedSearch : '';
    var string = label + from + subject + advancedSearch;

    return string;

}

function remainingQuota() {
    return MailApp.getRemainingDailyQuota();
}




function debug() {
//    Logger.log("DATE: " +PropertiesService.getUserProperties().getProperty('dateAsEpoch'));
    Logger.log(PropertiesService.getUserProperties().getProperties())
//    Logger.log("Remaining quota: " + MailApp.getRemainingDailyQuota());
//      var threads = GmailApp.search('after: 1509889580' + ' ' + search_filter_string(1));
//      Logger.log(threads[0].getMessages()[0].getSubject());
}
