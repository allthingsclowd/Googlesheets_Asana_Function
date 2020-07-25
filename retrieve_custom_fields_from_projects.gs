function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Asana BluePrints')
      .addItem('Import Portfolio', 'GET_ASANA_PORTFOLIO_UPDATE')
  .addToUi();
}

function GET_ASANA_PORTFOLIO_UPDATE() {
  
  var PORTFOLIO_ID = "Portfolio_ID";
  var token = "ASANA_BEARER_TOKEN";

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BlueprintFlags");
  if (sheet == null) {
    sheet = SpreadsheetApp.getActiveSpreadsheet();
    sheet.insertSheet("BlueprintFlags",0,);
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = SpreadsheetApp.getActiveRange();
  sheet.clear();
  sheet.appendRow([`Account`,
                  `Status`,
                  `Product`,
                  `Date`,
                  `Update`,
                  `TAM`,
                  `Journey`,
                  `SDLC`,
                  `Engagement`,
                  `Communication`,
                  `Deployment`,
                  `Sticky Features`,
                  `Usecase Fit`,
                  `Sponsorship`,
                  `Success Definition`,
                  `Customer Ability`,
                  `Adoption`,
                  `Sentiment Product`,
                  `Sentiment Services`,
                  `Sentiment Support`,
                  `Summary`]).setFrozenRows(1);
  
  sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn()).setBackground('#C0C0C0').setFontSize(10).setFontWeight("bold");
  
  // GET a list of the Asana project ids
  var options = {
    "method" : "GET",
    "headers" : {
      "contentType": "application/json",
      "Authorization": "Bearer " + token,
    }
 };
 //var custom_fields_res = UrlFetchApp.fetch('https://app.asana.com/api/1.0/portfolios/' + PORTFOLIO_ID + '?opt_fields=custom_field_settings', options);
 var projects_res = UrlFetchApp.fetch('https://app.asana.com/api/1.0/portfolios/' + PORTFOLIO_ID + '/items?opt_fields=name', options);
  
 // Make request to API and get response before this point.
 //var custom_fields = JSON.parse(custom_fields_res.getContentText());
 var projects = JSON.parse(projects_res.getContentText());

 //var custom_fields_string = JSON.stringify(custom_fields);
 var projects_string = JSON.stringify(projects);
  
 //Logger.log(custom_fields_string);
//Logger.log(projects_string); 
 //sheet.appendRow([`custom fields`,custom_fields_string]);
 //sheet.appendRow([`projects`,projects_string]);
  for(var i = 0; i < projects.data.length; i++)
  {
    //Logger.log("Company: " + projects.data[i].name + " Project Id: " + projects.data[i].gid); 
    var trimmed_summary = `Summary Missing. TAM to Update.`;
    var colour = `unset`;
    var title = `unset`;
    var status = `Summary Missing. TAM to Update.`;

    var project_id = (JSON.stringify(projects.data[i].gid)).replace(/['"]+/g, '');
    //Logger.log(`project id: ${project_id}`);
    var blueprint_details = (projects.data[i].name).split(/:/);
    var account_name = blueprint_details[0];
    var product = blueprint_details[3];
    
    // get the list of project update ids
    var project_details_res = UrlFetchApp.fetch('https://app.asana.com/api/1.0/projects/' + project_id + '?opt_pretty', options);
    var project_details = JSON.parse(project_details_res.getContentText());
    var project_details_string = JSON.stringify(project_details);
    //sheet.appendRow([`project details`,project_details_string]);
    //Logger.log("project details:",project_details);
    var project_status = project_details.data.current_status;
    var project_owner = project_details.data.owner.name;
    var custom_fields = project_details.data.custom_fields;
    var engagement_flag = `unset`;
    var sponsor_flag = `unset`;
    var successplan_flag = `unset`;
    var ability_flag = `unset`;
    var adoption_flag = `unset`;
    var usecase_flag = `unset`;
    var sticky_flag = `unset`;
    var lifecycle_flag = `unset`;
    var support_flag = `unset`;
    var services_flag = `unset`;
    var product_flag = `unset`;
    var communication_flag = `unset`;
    var sdlc_flag = `unset`;
    var deployment_flag = `unset`;
    
    //Logger.log(`project status: ${project_status}`);
    
    //sheet.appendRow([`Custom Fields`, JSON.stringify(custom_fields)]);
    
    custom_fields.forEach(function(item){
    
      var flag = `unset`;
    
      //sheet.appendRow([`Custom Field Name`, JSON.stringify(item.name)]);
      
      if (item.enum_value) {
        flag = item.enum_value.name;
      }
      
      //sheet.appendRow([`Custom Field Value`, flag]);
      
      switch(item.name) {
        case "Engagement":
          engagement_flag = flag;
          break;
          
        case "Sponsorship":
          sponsor_flag = flag;
          break;
          
        case "Success Definition":
          successplan_flag = flag;
          break;
          
        case "Customer Ability":
          ability_flag = flag;
          break;
          
        case "Adoption":
          adoption_flag = flag;
          break;
          
        case "Usecase Fit":
          usecase_flag = flag;
          break;
          
        case "Sticky Features":
          sticky_flag = flag;
          break;
          
        case "Journey Stage":
          lifecycle_flag = flag;
          break;
          
        case "Sentiment Support":
          support_flag = flag;
          break;
          
        case "Sentiment Services":
          services_flag = flag;
          break;
          
        case "Sentiment Product":
          product_flag = flag;
          break;
          
        case "Communication":
          communication_flag = flag;
          break;
         
        case "SDLC":
          sdlc_flag = flag;
          break;
          
        case "Deployment":
          deployment_flag = flag;
          break;
          
        default:
          break;
          
      }

      
    })

    
    
    if (project_status) {
      
     // var last_update = project_stat;
      // get the last project update text
      //var response = UrlFetchApp.fetch('https://app.asana.com/api/1.0/project_statuses/'+ last_update + '?opt_fields=color,title,text', options);
      //var current_update = JSON.parse(response.getContentText());
  
      colour = project_status.color;
      title = project_status.title;
      status = project_status.text;
      var summary = status.split(/\n\nWhat/);
      
      var length = 400;
      if (summary[0].length > length) {
        trimmed_summary = summary[0].substring(0, length);

      } else {
        trimmed_summary = summary[0];
        
      }
      
      if (trimmed_summary.lastIndexOf(".") > 0){
        trimmed_summary = trimmed_summary.substring(0, trimmed_summary.lastIndexOf(".") + 1);
      }
      trimmed_summary = trimmed_summary.replace(/\n/g, " ");
     
    } 
    
    sheet.appendRow([account_name,
                      colour,
                      product,
                      title,
                      status,
                      project_owner,
                      lifecycle_flag,
                      sdlc_flag,
                      engagement_flag,
                      communication_flag,
                      deployment_flag,
                      sticky_flag,
                      usecase_flag,
                      sponsor_flag,
                      successplan_flag,
                      ability_flag,
                      adoption_flag,
                      product_flag,
                      services_flag,
                      support_flag,
                      trimmed_summary]);

    var RED = "#FF0000";
    var YELLOW = "#FFFF00";
    var GREEN = "#00FF00";

    var bgColor = GREEN;
    // This changes font color
    if (colour == 'red') {
      bgColor = RED;
    } else if (colour == 'yellow') {
      bgColor = YELLOW;
    }
    
    sheet.getRange(sheet.getLastRow(), 2, 1, 1).setBackground(bgColor);
    sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn()).setFontSize(8);
  }
  sheet.autoResizeColumns(1, 21);
  sheet.sort(6).sort(2, false);
}
