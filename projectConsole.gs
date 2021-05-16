// GLOBAL var for this project console
// GLOBAL var from historyData.gs will still be used on this gs.
var ss = SpreadsheetApp.getActiveSpreadsheet();
var consoleSheet = ss.getSheetByName("Project List");
var projectList = ss.getRangeByName("ProjectListArea");

/**
* Simple Trigger: onOpen that will create a custom menu 
*/

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("TeamGantt")
  .addItem("Sign in", "showLoginDialog")
  .addSeparator()
  .addItem("Load Projects", "loadProjects")
  .addSubMenu(ui.createMenu('Scheduler')
              .addItem('Schedule the script', 'launchScheduleBox')
              .addItem('Remove all schedule', 'clearTrigger'))

  .addItem("Calculate Summary", "displayHistory")
  .addItem("Check Out of Rule Task", "checkTask")

  .addSeparator()
  .addItem("Sign out", "logOut")
  .addToUi();
}

/**
 * Creates a time-driven trigger.
 */
function launchScheduleBox() {
  if (PropertiesService.getUserProperties().getProperty('LOGIN') === null || PropertiesService.getUserProperties().getProperty('PASSWORD') === null || PropertiesService.getUserProperties().getProperty('authKey') === null){
    SpreadsheetApp.getUi().alert("Please Login first with complete information!");
    return;
  }
  var html = HtmlService.createHtmlOutputFromFile('scheduleBox')
    .setWidth(300)
    .setHeight(260);
  SpreadsheetApp.getUi()
    .showModalDialog(html, 'Schedule the script');
}

function createTimeDrivenTrigger(array){
  clearTrigger();
  var firstFunc = 'displayHistory';
  var secondFunc = 'loadProjects';
  var weekday;
  switch (array[1]){
    case 'MONDAY':
      weekday = ScriptApp.WeekDay.MONDAY 
      break;
    case 'TUESDAY':
      weekday = ScriptApp.WeekDay.TUESDAY 
      break;
    case 'WEDNESDAY':
      weekday = ScriptApp.WeekDay.WEDNESDAY 
      break;
    case 'THURSDAY':
      weekday = ScriptApp.WeekDay.THURSDAY 
      break;
    case 'FRIDAY':
      weekday = ScriptApp.WeekDay.FRIDAY 
      break;
    case 'SATURDAY':
      weekday = ScriptApp.WeekDay.SATURDAY 
      break;
    case 'SUNDAY':
      weekday = ScriptApp.WeekDay.SUNDAY 
      break;
  }
  var time = array[0].slice(0,2);
  if (time.includes(":")){
    time = time.slice(0,1);
  }
  // if the trigger hasnt been created, then create a displayHistory trigger at given time
  // and the load project trigger 1 HOUR EARLIER
  if(!(isTrigger(firstFunc) && isTrigger(secondFunc))) {
    ScriptApp.newTrigger(firstFunc)
    .timeBased()
    .onWeekDay(weekday)
    .atHour(time)
    .create();
    
    ScriptApp.newTrigger(secondFunc)
    .timeBased()
    .onWeekDay(weekday)
    .atHour(time-1)
    .create();
  }
  
  SpreadsheetApp.getUi().alert("The script has been scheduled on " + array[1] + " at " + array[0]);
}



/**
 * Function to prevent a duplicated trigger
 */

function isTrigger(funcName){
  var r=false;
  if(funcName){
    var allTriggers=ScriptApp.getProjectTriggers();
    for(var i=0;i<allTriggers.length;i++){
      if(funcName==allTriggers[i].getHandlerFunction()){
        r=true;
        break;
      }
    }
  }
  return r;
}

function clearTrigger(){
  if (PropertiesService.getUserProperties().getProperty('LOGIN') === null || PropertiesService.getUserProperties().getProperty('PASSWORD') === null || PropertiesService.getUserProperties().getProperty('authKey') === null){
    SpreadsheetApp.getUi().alert("Please Login first with complete information!");
    return;
  }
 // Deletes all displayHistory trigger in the current project.
  var func = 'displayHistory';
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === func){
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
                     
           
/**
 * Function to perform authentication and load project list
 * @return 2D array that consist of four data: project name, project id , project status and project public key
 */

function loadAllProjects() {
  // perform authentication
  authenticate();
  // initiate URL, header, options
  var getAllProjectURL = "https://api.teamgantt.com/v1/projects/all?"
  var header = {
    'TG-Authorization' : 'Bearer ' + PropertiesService.getUserProperties().getProperty('authKey'),
    'TG-Api-Key' : api_key,
    'TG-User-Token' : user_token
  }
  var options = {
      'method' : 'get',
      'contentType' : 'application/json',
      'headers' : header
  };  
  // call the API and parse the response
  var response = UrlFetchApp.fetch(getAllProjectURL, options);
  var json = response.getContentText();
  var data = JSON.parse(json);
  var myMap = [];
  // map name, id, status, and public key
  for (i = 0; i < data.projects.length; i++) {
    myMap.push([data.projects[i].name, data.projects[i].id.toString(), data.projects[i].status, data.projects[i].public_key]);
  }
  return myMap;
}

/**
 * Function to write data that are returned by loadAllProjects to Google Sheet "Project List" tab
 */

function loadProjects(){
  // check if user already login or not, otherwise prompt ui alert
  if (PropertiesService.getUserProperties().getProperty('LOGIN') === null || PropertiesService.getUserProperties().getProperty('PASSWORD') === null 
  || PropertiesService.getUserProperties().getProperty('authKey') === null){
    SpreadsheetApp.getUi().alert("Please Login first with complete information!");
    return;
  }
  // load the project data and put it on Project List tab

  var myMap = loadAllProjects();
  consoleSheet.getRange(3, 1, consoleSheet.getLastRow(), consoleSheet.getLastColumn()).clearContent();
  for (var i = 0; i < myMap.length; i++){
    for (var j = 0; j < myMap[i].length; j++){
      consoleSheet.getRange(projectList.getRow()+i, projectList.getColumn()+j).setValue(myMap[i][j]);
    }
  }

}

/**
 * To show login dialog everytime user press Sign In menu
 * Please check html for the layout
 */

function showLoginDialog() {
  
  if (PropertiesService.getUserProperties().getProperty('LOGIN') != null ){
    SpreadsheetApp.getUi().alert("You are currently login as " + PropertiesService.getUserProperties().getProperty('LOGIN'));
  } else {
  
    var html = HtmlService.createHtmlOutputFromFile('popup')
    .setWidth(300)
    .setHeight(360);
    
    SpreadsheetApp.getUi()
    .showModalDialog(html, 'TeamGantt Sign in Information');
  }
  
}

/**
 * When user pressed submit, this function will grab all the user filled information and store it on PropertiesService
 * @arg {dataArray} array of submitted data from form_data() js function on html doc
 */


function grabLoginInformation(dataArray){
  // data from user and store on PropertiesService object 
  var newProperties = {LOGIN: dataArray[0], PASSWORD: dataArray[1], authKey: dataArray[2]};
  PropertiesService.getUserProperties().setProperties(newProperties, true);
  // check if the data is valid
  checkLogin();

}

/**
 * To validate Login, Password and app token entered by user by performing authentication
 * If the authentication is not succesful, return the message to user and delete all saved PropertiesService
 */

function checkLogin(){
  // perform authentication, if successful return ui alert, otherwise return error message and clear all PropertiesService
  try{
    authenticate();
    SpreadsheetApp.getUi().alert('Login Successful!');
  } catch (e){
    SpreadsheetApp.getUi().alert(e.message);
    PropertiesService.getUserProperties().deleteAllProperties();
  }
}

/**
 * To logout and delete all PropertiesService 
 */

function logOut(){
  PropertiesService.getUserProperties().deleteAllProperties();
  SpreadsheetApp.getUi().alert('You have logged out. Thank you.');
}

