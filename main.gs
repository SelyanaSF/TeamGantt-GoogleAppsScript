// GLOBAL variable to authenticate API 
var ss = SpreadsheetApp.getActiveSpreadsheet();
var summarySheet = ss.getSheetByName("Usage Summary");
var dataSheet = ss.getSheetByName("Data");
var archiveSheet = ss.getSheetByName("Archive");
var dataRange = ss.getRangeByName("data");
var dateRange = ss.getRangeByName("dateofWeeks");
var actionRange = ss.getRangeByName("actionRange");
var api_key = '';
var user_token = '';

/**
 * Authentication process that grab the api key and user token
 * Nothing is returned but the GLOBAL VAR api_key and user_token will be filled
 */

function authenticate(){
  // prepare URL, header, payload(body), and options
  var urlAUTH = 'https://api.teamgantt.com/v1/authenticate';
  var header = {
    'TG-Authorization' : 'Bearer ' + PropertiesService.getUserProperties().getProperty('authKey'),
  }
  var body = {
    'user' : PropertiesService.getUserProperties().getProperty('LOGIN'),
    'pass' : PropertiesService.getUserProperties().getProperty('PASSWORD')
  }
  var options = {
    'method' : 'post',
    'contentType' : 'application/json',
    'headers' : header,
    'payload' : JSON.stringify(body)
  };
  // call the API and parse the data
  var response = UrlFetchApp.fetch(urlAUTH, options);
  var json = response.getContentText();
  var data = JSON.parse(json);
  // store api_key and user_token as global var
  if (data != null) {
    api_key = data["api_key"];
    user_token = data["user_token"];
  } else {
    Logger.log("Wrong user/pass");
  }
}

/**
 * @arg {url} url to fetch
 * @arg {method} HTTP method (GET/POST)
 * @arg {body} Payload. could be null if no need body, but still need to be passed
 * @arg {projectID} TeamGantt project id 
 * @arg {projectPublicKey} TeamGantt project API key
 * @return {data} a parsed data
 */

function authenticateWithProjectID(url, method, body, projectID, projectPublicKey){
  // perform authentication 
  authenticate();
  // prepare header and options
  var header = {
    'TG-Authorization' : 'Bearer ' + PropertiesService.getUserProperties().getProperty('authKey'),
    'TG-Api-Key' : api_key,
    'TG-User-Token' : user_token,
    'TG-Authorization-ProjectIds' : projectID,
    'TG-Authorization-PublicKeys' : projectPublicKey
  }
  if (body != null){

    var options = {
      'method' : method,
      'contentType' : 'application/json',
      'headers' : header,
      'payload' : JSON.stringify(body)
    };
  } else {

    var options = {
    'method' : method,
    'contentType' : 'application/json',
    'headers' : header
    };
  }
  // call the API and parse the data
  var response = UrlFetchApp.fetch(url, options);
  var json = response.getContentText();
  var data = JSON.parse(json);
  return data;
}

/**
 * Main function for this program. 
 */

function displayHistory(){
  // check if user has login or not, otherwise call ui alert and dont continue
  if (PropertiesService.getUserProperties().getProperty('LOGIN') === null || PropertiesService.getUserProperties().getProperty('PASSWORD') === null || PropertiesService.getUserProperties().getProperty('authKey') === null){
    SpreadsheetApp.getUi().alert("Please Login first with complete information!");
    return;
  }
  // get the project ID map with key
  var projectIDmap = mapProjectIDwithKey();
  var sortedOutput = [];

    // each[0] is project ID and each[1] is project key
  projectIDmap.forEach(function (idMap) {
    
    // for each map, get the all its tasks IDs (both sub-group and normal task)
    var tasksURL = "https://api.teamgantt.com/v1/tasks?project_ids=" + idMap[0] + "&project_status[]=Active&project_status[]=On+Hold&project_status[]=Complete";
    var taskArrays = getTaskIDs(tasksURL,idMap[0],idMap[1]);
    var groupURL = "https://api.teamgantt.com/v1/groups?project_ids=" + idMap[0] + "&flatten_children=true&project_status[]=Active&project_status[]=On+Hold&project_status[]=Complete";
    var groupTaskArrays = getTaskIDs(groupURL, idMap[0], idMap[1]);
    
    // sorted entry for task and group task. 
    //1. Get history for each ID 
    //2. After, put each history to an output array var (sortedOutput)
    if (taskArrays.length > 0){
      taskArrays.forEach(function(taskArray){
        var taskHistoryURL = "https://api.teamgantt.com/v1/tasks/" + taskArray[0] + "/history"
        getHistoryById(taskHistoryURL, idMap[0], idMap[1],taskArray[1], sortedOutput);
      });
    }
    if (groupTaskArrays.length > 0){
      groupTaskArrays.forEach(function(groupTaskArray){
        var groupTaskHistoryURL = "https://api.teamgantt.com/v1/groups/" + groupTaskArray[0] + "/history";
        getHistoryById(groupTaskHistoryURL, idMap[0], idMap[1], groupTaskArray[1], sortedOutput);
      });
    }
  });
  //Logger.log(sortedOutput);
  // length of row = arrayoutput length and length of column = array[0] length.
  // after that clear content on the sheet and write the data
  var rowLen = sortedOutput.length;
  var columnLen = sortedOutput[0].length;
  if (dataSheet.getLastRow() > 0){
    dataSheet.getRange(3,1,dataSheet.getLastRow(),columnLen).clearContent();
  }
  dataSheet.getRange(3,1,rowLen,columnLen).setValues(sortedOutput);
  
  // after the data is written, make the calculation
  writeCalculation();
}



/**
 * This function would be a request back to the TG API and return an array of IDs of task/sub-group
 * @arg {url} url to fetch
 * @arg {projectID} TeamGantt project id 
 * @arg {projectPublicKey} TeamGantt project API key
 * @return {IDarray} an array of task IDs
 */
function getTaskIDs(url,projectID, projectPublicKey){
  // call function authenticatewithprojectid and grab the data
  var body;
  var data = authenticateWithProjectID(url,'get',body,projectID, projectPublicKey);
  var IDarrays = [];
  var dateCompare = formatDate(addMonths(new Date(), -1));
  // filter which task ID needs to be returned
  data.forEach(function(elem){
    // no need to return completed task that has been completed more than 2 months ago to reduce output and runtime...
    if (!(elem["end_date"] <= dateCompare && parseInt(elem["percent_complete"]) == 100)){
      IDarrays.push([elem["id"].toString(), elem["percent_complete"]]);
    };
  });
  Logger.log(IDarrays);
  return IDarrays;
}

/**
 * This function would be a request back to the TG API and return an array of history events. 
 * Events are filtered by 3 actions: updated, added, rescheduled
 * @arg {url} url to fetch, url should contain task id
 * @arg {projectId} TeamGantt task id 
 * @arg {projectPublicKey} TeamGantt project API key
 * @return {historyArray} an array of histories
 */
function getHistoryById(url,projectId, projectPublicKey, percentComplete, outputArray){
  // call function authenticatewithprojectid and grab the data
  var body;
  var data = authenticateWithProjectID(url, 'get', body, projectId, projectPublicKey);
  //Logger.log(data);
  // only grab timestamp, action, and projectid
  // Filter those timestamp.
  
  // var historyArray = data.map(function(elem){
  //   return ([elem["time_stamp"], elem["details"]["action"], elem["project_id"].toString(), percentComplete]);
  // });
  data.map(function(elem){
    outputArray.push([elem["time_stamp"], elem["details"]["action"], elem["project_id"].toString(), percentComplete]);
  });

  //return historyArray;
}

/**
 * @return {newMap} a 2d array of project id with key 
 */

function mapProjectIDwithKey(){
  var newMap = [];
  // grab data from ProjectListArea range and push to newMap var
  var lookupValues = ss.getRangeByName("ProjectListArea").getValues();
  lookupValues.forEach(function(each){
    if (each[3]){
      newMap.push([each[1].toString(),each[3]]);
    }
  });
  var newMap = newMap.filter(function (el) {
    return el != null;
  });
  return newMap;
}

/**
 * Function to write calculation for each action / date combination 
 */

function writeCalculation(){
  archiveData();
  // get date from date of weeks range
  var tempp = dateRange.getValues();
  var arrayofDates = [];
  // filter out empty entry
  tempp.forEach(function(eachDate){
    if (eachDate[0]){
      if (eachDate[0] != 'Week of'){
        arrayofDates.push(eachDate);
      }
    };
  });
  arrayofDates = arrayofDates[arrayofDates.length-1];
  
  // get actions from actionRange range
  var arrayofActions = actionRange.getValues()[0];
  var result = [];
  var resultArray = [];

  //Loop the result (i and j) and perform a calculate function for the combination of date and action
  var date = arrayofDates[0];
  Logger.log(date);
  for (var j = 0; j < arrayofActions.length; j++){
    result.push(calculate(arrayofActions[j].toLowerCase(), date));
  }
  resultArray.push(result);
  
  Logger.log(resultArray);
  // store the result on summarySheet
  //var row = actionRange.getRow()+1;
  var row = summarySheet.getLastRow();
  var col = actionRange.getColumn();
  var maxcol = resultArray[0].length;
  var maxrow = resultArray.length;
  summarySheet.getRange(row, col, summarySheet.getLastRow(), summarySheet.getLastColumn()).clearContent();
  summarySheet.getRange(row, col, maxrow, maxcol).setValues(resultArray);
}

/**
 * Function to count how many data that matches action & within the date on Data tab
 * @arg {action} action name 
 * @arg {date} date of the week to filter
 */

function calculate(action, date) {
  
  // get all the data from "Data" tab
  var values = dataRange.getValues();
  var count = 0;
  
  // start date and end date to create filter
  var startDate = new Date(date);
  var endDate = addDays(startDate, 7);
  
  // loop through every data and if the data matches the condition, add count
  // condition 1a: action (each[1]) matches the argument
  // condition 1b: if action == update percent complete of , look at percent_complete, count if only its 100%
  // condition 2: timestamp (each[0]) is between the start date and end date

  values.forEach(function(each){
    if ((action.includes(each[1]) || each[1].includes(action)) && (new Date(each[0]) >= startDate && new Date(each[0]) < endDate)){
      if (action.includes("updated percent complete of")){
        if (each[3] == 100){
          count++;
        }
      } else {
        count++;
      }
    }
  });
  return count;
}

/**
 * Function to keep the latest row data on Usage Summary and move the rest to archive tab
 * @arg {action} action name 
 * @arg {date} date of the week to filter
 */

function archiveData(){
  // select area to archive by offseting from actionRange
  // offset 1 row  and -1 column, with numofrows: 1 and numofcolumn : 4
  var dataForArchive = actionRange.offset(1, -1, 1, 4).getValues();
  
  // remove the last entry bcs we dont want to archive this and move the entry to the 'Archive' tab
  archiveSheet.getRange(archiveSheet.getLastRow()+1,1,1,dataForArchive[0].length).setValues(dataForArchive);
  
  // remove the first row, then populate a new date
  summarySheet.deleteRow(actionRange.getRow()+1);
  var lastDate = summarySheet.getRange(summarySheet.getLastRow(), 1).getValue();
  summarySheet.getRange(summarySheet.getLastRow()+1,1).setValue(addDays(lastDate,7));
  
  // copy the format
  summarySheet.getRange(summarySheet.getLastRow()-1, 1, 1, 4).copyTo(summarySheet.getRange(summarySheet.getLastRow(), 1, 1, 4), {formatOnly:true});
  archiveSheet.getRange(archiveSheet.getLastRow()-1, 1, 1, 4).copyTo(archiveSheet.getRange(archiveSheet.getLastRow(), 1, 1, 4), {formatOnly:true});
}

function addDays(date, days) {
  var result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function addMonths(date, months) {
  var result = new Date(date);
  var month = (result.getMonth() + months) % 12;
  //create a new Date object that gets the last day of the desired month
  var last = new Date(result.getFullYear(), month + 1, 0);

  //compare dates and set appropriately
  if (result.getDate() <= result.getDate()) {
    result.setMonth(month);
  }

  else {
    result.setMonth(month, last.getDate());
  }

  return result;
}

function formatDate(date) {
  var d = new Date(date);
  var month = '' + (d.getMonth() + 1);
  var day = '' + d.getDate();
  var year = d.getFullYear();

  if (month.length < 2){
    month = '0' + month;
  }
  if (day.length < 2){ 
    day = '0' + day;
  }
  return [year, month, day].join('-');
}

                        
                        



