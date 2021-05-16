var OTRSheet = ss.getSheetByName("Out of Rule Task");

function checkTask(){
  // check if user has login or not, otherwise call ui alert and dont continue
  if (PropertiesService.getUserProperties().getProperty('LOGIN') === null || PropertiesService.getUserProperties().getProperty('PASSWORD') === null || PropertiesService.getUserProperties().getProperty('authKey') === null){
    SpreadsheetApp.getUi().alert("Please Login first with complete information!");
    return;
  }
  // get the project ID map with key
  var projectIDmap = mapActiveProjectIDwithKey();
  var listOfTask = [];

  // each[0] is project ID and each[1] is project key
  projectIDmap.forEach(function (each) {
    
    // for each map, get the all its tasks details 
    var tasksURL = "https://api.teamgantt.com/v1/tasks?project_ids=" + each[0] + "&project_status[]=Active&project_status[]=On+Hold&project_status[]=Complete";
    var taskDetails = getTaskDetails(tasksURL,each[0],each[1]);
    taskDetails.forEach((each) => {
      listOfTask.push(each);
    })
  });
  listOfTask = mergeEntry(listOfTask);
  OTRSheet.getRange(3, 1, OTRSheet.getLastRow(), OTRSheet.getLastColumn()).clearContent();
  OTRSheet.getRange(3, 1, listOfTask.length, OTRSheet.getLastColumn()).setValues(listOfTask);

}

function mapActiveProjectIDwithKey(){
  var newMap = [];
  // grab data from ProjectListArea range and push to newMap var
  var lookupValues = ss.getRangeByName("ProjectListArea").getValues();
  lookupValues.forEach(function(each){
    if (each[3]){
      if (each[2] === 'Active'){
        newMap.push([each[1].toString(),each[3]]);
      }
    }
  });
  var newMap = newMap.filter(function (el) {
    return el != null;
  });
  return newMap;
}

function getTaskDetails(url,projectID, projectPublicKey){
  // call function authenticatewithprojectid and grab the data
  var body;
  var url = "https://api.teamgantt.com/v1/tasks?project_ids=" + projectID + "&project_status[]=Active&project_status[]=On+Hold";
  var data = authenticateWithProjectID(url,'get',body,projectID, projectPublicKey);
  var arrayOfTaskDetails = [];
  var dateCompare = formatDate(new Date());
  var listOfTask = [];
  
  data.forEach(function(elem){

    var taskLength = (new Date(elem["end_date"]).getTime() - new Date(elem["start_date"]).getTime()) / (1000 * 3600 * 24);
    
    // if the task is not complete 
    if (parseInt(elem["percent_complete"]) < 100 && elem["end_date"] >= dateCompare){
      
      // this is to prevent a duplicated entry
      if (!listOfTask.includes(elem["name"])){
        
        // return task that end_date - start_date is more than 5 days 
        if (taskLength > 4){
          var assignee = getResourceName(elem).join(', ');
          arrayOfTaskDetails.push([elem["name"],elem["project_name"],elem["start_date"],elem["end_date"],elem["estimated_hours"],assignee,"Task is scheduled more than 1 week"]);
          listOfTask.push(elem["name"].toString());
        }
        // return task that has more than 24 hours
        if (parseInt(elem["estimated_hours"]) > 24){
          var assignee = getResourceName(elem).join(', ');
          arrayOfTaskDetails.push([elem["name"],elem["project_name"],elem["start_date"],elem["end_date"],elem["estimated_hours"],assignee,"Task has more than 24 hours"]);
          listOfTask.push(elem["name"].toString());
        }
        
      }
    }
  });

  return arrayOfTaskDetails;
}

function getResourceName(json){
  var name = [];
  var resources = json["resources"];

  if (resources != null){
    for (var i = 0; i<resources.length; i++){
      name.push(resources[i].name)
    }
  }
  return name;
}

function mergeEntry(arrayOfTask){

  
  // filter out those empty row
  arrayOfTask = arrayOfTask.filter(function(r){ 
    return r.join('').length>0;
  });
  var taskNameList = [];
  var cleanArray = [];
  arrayOfTask.forEach((eachTask) => {
    if (taskNameList.includes(eachTask[0])){
      cleanArray.forEach((cleanTask) => {
        if (cleanTask[0] == eachTask[0]){
          cleanTask[6] = cleanTask[6] + ", " + eachTask[6];
          
        } 
      })
  
    } else {
      taskNameList.push(eachTask[0]);
      cleanArray.push(eachTask);
    }
  })

  return cleanArray;

}




