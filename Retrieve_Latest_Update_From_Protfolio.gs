function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Asana BluePrints')
      .addItem('Import Portfolio', 'GET_ASANA_PORTFOLIO_UPDATE')
  .addToUi();
}

function GET_ASANA_PORTFOLIO_UPDATE() {
  
  var PORTFOLIO_ID = "Portfolio_ID Goes Here";
  var token = "<Asana Token Goes Here>";
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.setName("OppBluePrints");
  sheet.clearContents();
  sheet.appendRow([`ACCOUNT`, `STATUS`, `PRODUCT`, `DATE`, `TAM UPDATE`]);
  
  // GET a list of the Asana project ids
  var options = {
    "method" : "GET",
    "headers" : {
      "contentType": "application/json",
      "Authorization": "Bearer " + token,
    }
 };
 var response = UrlFetchApp.fetch('https://app.asana.com/api/1.0/portfolios/' + PORTFOLIO_ID + '/items?opt_fields=name', options);
 
 // Make request to API and get response before this point.
 var detail = JSON.parse(response.getContentText());
 //Logger.log(data);
  
  for(var i = 0; i < detail.data.length; i++)
  {
    Logger.log("Company: " + detail.data[i].name + " Project Id: " + detail.data[i].gid); 
    var trimmedSummary = `Summary Missing. TAM to Update.`;
    var colour = `unset`;
    var title = `unset`;
    var status = `Summary Missing. TAM to Update.`;

    var PROJECT_ID = (JSON.stringify(detail.data[i].gid)).replace(/['"]+/g, '');
    Logger.log(`project id: ${PROJECT_ID}`);
    var BLUEPRINT_DETAILS = (detail.data[i].name).split(/:/);
    var ACCOUNT_NAME = BLUEPRINT_DETAILS[0];
    var PRODUCT = BLUEPRINT_DETAILS[3];
    
    // get the list of project update ids
    var response = UrlFetchApp.fetch('https://app.asana.com/api/1.0/projects/' + PROJECT_ID + '/project_statuses?opt_pretty', options);
    
    var status_updates = JSON.parse(response.getContentText());
    
    if (status_updates.data.length > 0) {
      
      var last_update = status_updates.data[0].gid;
      // get the last project update text
      var response = UrlFetchApp.fetch('https://app.asana.com/api/1.0/project_statuses/'+ last_update + '?opt_fields=color,title,text', options);
      var current_update = JSON.parse(response.getContentText());
  
      colour = current_update.data.color;
      title = current_update.data.title;
      status = current_update.data.text;
      var summary = status.split(/\n\nWhat/);
      
      var length = 400;
      if (summary[0].length > length) {
        trimmedSummary = summary[0].substring(0, length);

      } else {
        trimmedSummary = summary[0];
        
      }
      
      if (trimmedSummary.lastIndexOf(".") > 0){
        trimmedSummary = trimmedSummary.substring(0, trimmedSummary.lastIndexOf(".") + 1);
      }
      trimmedSummary = trimmedSummary.replace(/\n/g, " ");
     
    } 
    sheet.appendRow([ACCOUNT_NAME, colour, PRODUCT, title, trimmedSummary]);


  }
  sheet.autoResizeColumns(1, 5);

}
