function onInstall() {
  onOpen();
}

/* What should the add-on do when a document is opened */
function onOpen() {
  SpreadsheetApp.getUi()
  .createAddonMenu()
      .addItem('Create New Rule', 'create_new_rule')
      .addItem('Manage Forwarding Rules', 'manage_rules')
      .addSeparator()
      .addItem('View Remaining Quota', 'remaining_quota')
      .addItem('Quick Start Guide', 'showDialog')

  .addToUi();  // Run the showSidebar function when someone clicks the menu
}

function createTimeDrivenTriggers() {
  // Trigger every 6 hours.
  ScriptApp.newTrigger('gmailAutoForwarderStatus')
      .timeBased()
      .everyHours(1)
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
function manage_rules(){
 var html = HtmlService.createHtmlOutputFromFile('manageRules_Dialog')
      .setWidth(500)
      .setHeight(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Manage Forwarding Rules');
}

// Assistant function that passes list of labels and alias emails to the new_rule html page
function getData() {

  //Getting the user labels
  var labels = GmailApp.getUserLabels();
  for(i=0;i<labels.length;++i){
   labels[i]=labels[i].getName()
  }

  //Getting current email + available alias email addresses
  var aliasEmails = GmailApp.getAliases();
  aliasEmails.push(Session.getActiveUser().getEmail())

  var data = {"labels":labels,"emailAddresses":aliasEmails}
  return(data);
}

/*  Function that gets called when user clicks on the create new rule button from the form */
function addRule(formInput){
  var userProperties = PropertiesService.getUserProperties();
  var count;
  if(userProperties.getProperty('count') == null){
   userProperties.setProperty('count', '0');
  }
  //getting count again- IMP
  var count = parseInt(userProperties.getProperty('count'));
  Logger.log(count)

  count += 1;
  userProperties.setProperty('count', count.toString());

  userProperties.setProperty('rule_'+ userProperties.getProperty('count'),JSON.stringify(formInput));
}

// Assistant function used to pass currently stored rules to the manage_rules html page
function getRules(){
 Logger.log(PropertiesService.getUserProperties().getProperties());
 var count = PropertiesService.getUserProperties().getProperty('count');
 var rulesList = []
 for(var i=1;i<=count;++i){
   var rule = PropertiesService.getUserProperties().getProperty('rule_'+i);
   rule = JSON.parse(rule)
   var label = (rule.label)?'Label:'+rule.label : '';
   var from = (rule.from)?'From:'+rule.from : '';
   var subject = (rule.subject)?'Subject:' + rule.subject : '';
   var advancedSearch = (rule.advancedSearch)?'Advanced Search:'+rule.advancedSearch : '';
   var toEmail = 'forwarded to:' + rule.forwardTo;
   var string = label + ' ' + from + ' ' +subject + ' ' + advancedSearch + ' '+ toEmail ;
   rulesList.push({'rule':i,'str':string});
 }
 return rulesList

}

function deleteAll(){
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

function deleteRule(ruleNum){
  /*  Case : Suppose the rules contains [rule1,rule2,rule3] and the user opts to delete rule1,
             then the whole array has to be shifted by 1 index to right and count value should be count-1
  */
  var userProperties = PropertiesService.getUserProperties();
  var count = parseInt(userProperties.getProperty('count'));

  //shifting the array to right
  for(var i=ruleNum;i<=count-1;++i){
   userProperties.setProperty('rule_'+i, userProperties.getProperty('rule_'+(i+1)));
  }

  // after deleting rule1 from [rule1,rule2,rule3] the rule2 and rule3 becomes same => delete last rule after each shift
  userProperties.deleteProperty('rule_'+ count);

  if(count==1){
    userProperties.setProperty('count', '0');
  }else{
    userProperties.setProperty('count', (count-1).toString());
  }

}


function debug(){
//  PropertiesService.getUserProperties().deleteAllProperties();
  Logger.log(PropertiesService.getUserProperties().getProperties())

}
