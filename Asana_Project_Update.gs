function GET_ASANA_UPDATE(PROJECT_ID) {
  
  var token = "INSERT ASANA BEARER TOKEN HERE";

  // GET a list of the Asana project status updates
  var options = {
     "method" : "GET",
     "headers" : {
       "contentType": "application/json",
       "Authorization": "Bearer " + token,
     }
  };
  var response = UrlFetchApp.fetch('https://app.asana.com/api/1.0/projects/' + PROJECT_ID + '/project_statuses?opt_pretty', options);
  
  // Make request to API and get response before this point.
  var json = response.getContentText();
  var data = JSON.parse(json);
  // Logger.log(data.title);
   
   
  var status_updates = JSON.parse(response.getContentText());
  // Logger.log(`status id: ${status_updates.data[0].gid}`);
  // Logger.log(`status title: ${status_updates.data[0].title}`);

  var last_update = status_updates.data[0].gid;

  // GET the most recent Asana project status update
  var options = {
     "method" : "GET",
     "headers" : {
       "contentType": "application/json",
       "Authorization": "Bearer " + token,
     }
  };
  var response = UrlFetchApp.fetch('https://app.asana.com/api/1.0/project_statuses/'+ last_update + '?opt_fields=color,title,text', options);
  
  // Make request to API and get response before this point.
  var json = response.getContentText();
  var current_update = JSON.parse(json);

  var colour = current_update.data.color + '\n';
  var title = current_update.data.title + '\n';
  var status = current_update.data.text + '\n';
  
  // Logger.log(`colour: ${colour}`);
  // Logger.log(`title: ${title}`);
  // Logger.log(`text: ${status}`);
  
  var message = colour.concat(title.concat(status));
  
  // Logger.log(`message: ${message}`);
  
  
 // GET the Asana project name
  var options = {
     "method" : "GET",
     "headers" : {
       "contentType": "application/json",
       "Authorization": "Bearer " + token,
     }
  };
  var response = UrlFetchApp.fetch('https://app.asana.com/api/1.0/projects/' + PROJECT_ID + '?opt_fields=name', options);
  
  // Make request to API and get response before this point.
  var json = response.getContentText();
  var blueprint_name = JSON.parse(json);
  // Logger.log(`blueprint name: ${blueprint_name.data.name}`);
  
  var blueprint_name = blueprint_name.data.name + '\n';
  var response = blueprint_name.concat(message);
  // Logger.log(`Response: ${response}`);

  return response;
}
