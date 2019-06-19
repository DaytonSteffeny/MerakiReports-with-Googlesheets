// Settings - Modify this with your values
// *************************

// User Defined in the Script
var API_KEY = '';
var ORG_ID = '';
var NET_ID = '';
var TIMESPAN = '';

// User Defined in a Sheet
var SHEET_NAME = "settings"
var API_KEY_SHEET_CELL = "B3";
var API_KEY_SHEET_CELL_LABEL = "A3";
var ORG_ID_SHEET_CELL = "B4";
var ORG_ID_SHEET_CELL_LABEL = "A4";
var NET_ID_SHEET_CELL = "B5";
var NET_ID_SHEET_CELL_LABEL = "A5";
var TIMESPAN_SHEET_CELL = "B6";
var TIMESPAN_SHEET_CELL_LABEL = "A6";




// *************************
// Initialize Settings Sheet and Environment Variables

// find or create settings sheet
var ss = SpreadsheetApp.getActiveSpreadsheet();
if (ss.getSheetByName(SHEET_NAME) == null){
  ss.insertSheet(SHEET_NAME); 
  ss.getRange(API_KEY_SHEET_CELL_LABEL).setValue('API KEY:');
  ss.getRange(ORG_ID_SHEET_CELL_LABEL).setValue('Org ID:');
  ss.getRange(NET_ID_SHEET_CELL_LABEL).setValue('Net ID:');
  ss.getRange(API_KEY_SHEET_CELL).setValue('YourAPIKey');
  ss.getRange(ORG_ID_SHEET_CELL).setValue('YourOrgId');
  ss.getRange(NET_ID_SHEET_CELL).setValue('YourNetId (optional)');
  ss.getRange(TIMESPAN_SHEET_CELL_LABEL).setValue('Timespan:');
  ss.getRange(TIMESPAN_SHEET_CELL).setValue(7200);
}
var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

// assign settings
var settings = {};
settings.apiKey = settingsSheet.getRange(API_KEY_SHEET_CELL).getValue() || API_KEY; 
settings.orgId = settingsSheet.getRange(ORG_ID_SHEET_CELL).getValue() || ORG_ID;
settings.netId = settingsSheet.getRange(NET_ID_SHEET_CELL).getValue() || NET_ID;
settings.timespan = settingsSheet.getRange(TIMESPAN_SHEET_CELL).getValue() || TIMESPAN;


function setApiKey(apiKey){
  settings.apiKey = apiKey;
  settingsSheet.getRange(API_KEY_SHEET_CELL).setValue(apiKey)
}

function setOrgId(orgId){
  settings.orgId = orgId;
  settingsSheet.getRange(ORG_ID_SHEET_CELL).setValue(orgId)
}

function setNetId(netId){
  settings.netId = netId;
  settingsSheet.getRange(NET_ID_SHEET_CELL).setValue(netId)
}

function setTimespan(timespan){
  settings.timespan = timespan;
  settingsSheet.getRange(TIMESPAN_SHEET_CELL).setValue(timespan)
}

// Toolbar Menu Items
function onOpen() {

  loadMenu();
}  

//var orgFunctions = {};

function loadMenu() {
  // Main Menu
  var ui = SpreadsheetApp.getUi();
  var mainMenu = ui.createMenu('Meraki-Reports');
      
  
  // Reports
  mainMenu.addItem('Organizations','callOrgs')
  mainMenu.addItem('Networks','callNetworks')
      
  

  
  mainMenu.addSubMenu(SpreadsheetApp.getUi().createMenu('Org-Wide')    
      .addItem('Admins','callAdmins')
      .addItem('Clients','callClientsOfOrg')
      .addItem('Clients Details','callClientsOfOrgDetails')
      .addItem('Configuration Templates','callConfigTemplates')
      .addItem('Devices Details','callDevices')
      .addItem('Devices Status','callDevicesStatuses')
      .addItem('Devices Uplink Details','callUplinkInfos')
      .addItem('License State','callLicenseState')  
      .addItem('License State all Orgs','callLicenseStateForOrgs')
      .addItem('License State Details','callLicenseStateDetails')  
      .addItem('Group Policies','callGroupPoliciesOfOrg')
      .addItem('Inventory','callInventory')
      .addItem('Networks','callNetworks')
      .addItem('Organizations','callOrgs')
      .addItem('Organization','callOrg')
      .addItem('SSIDs','callSsidsOfOrg')
      .addItem('StaticRoutes','callStaticRoutes')
      .addItem('Wireless Health Connection Stats','callConnectionStatsOfOrg')
      .addItem('Wireless Health Latency Stats','callLatencyStatsOfOrg')
      .addItem('Traffic Analysis','callTrafficOfOrg') // Testing 
      .addItem('VLANS','callVlansOfOrg') 
      .addItem('VPN','callSiteToSiteVpn'));
      
  mainMenu.addSubMenu(SpreadsheetApp.getUi().createMenu('Network-Wide')
      .addItem('Clients','callClientsOfNet')
      .addItem('Clients Details','callClientsOfNetDetails')    
      .addItem('Device Loss and Latency History','callDeviceLossAndLatencyHistory')
      .addItem('Wireless Health Failed Connections','callFailedConnections')                  
      .addItem('Wireless Health Connection Stats by Device','callConnectionStatsByNode')
      .addItem('Wireless Health Connection Stats by Client','callConnectionStatsByClient')
      .addItem('Wireless Health Latency Stats by Device','callLatencyStatsByNode')
      .addItem('Wireless Health Latency Stats by Client','callLatencyStatsByClient'));
  
  mainMenu.addSubMenu(SpreadsheetApp.getUi().createMenu('Settings')
      .addItem('Set API Key','promptApiKey')
      .addItem('Set Org ID','promptOrg')
      .addItem('Set Net ID','promptNet')
      .addItem('Set Timespan','promptTimespan'));
  
  mainMenu.addToUi() ;

}

/*
function selectOrg(id){
  var ui = SpreadsheetApp.getUi();
  settings.orgId = id;
  ui.alert('Organization Set To: '+settings.orgId, ui.ButtonSet.YES_NO);
}
*/


function promptApiKey(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Meraki API Key', 'Required to run reports.', ui.ButtonSet.OK_CANCEL);
  
  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.OK) {
    
    // save key and refresh menu
    setApiKey(response.getResponseText());
    //ui.alert('API Key Set: '+settings.apiKey, ui.ButtonSet.YES_NO);

    loadMenu();
  } else if (response.getSelectedButton() == ui.Button.CANCEL) {
    Logger.log('The user didn\'t want to provide an API key.');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}

function promptOrg(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Organization ID', 'Run the Organizations report to get this info.', ui.ButtonSet.OK_CANCEL);
  
  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.OK) {
    Logger.log('The user\'s Org Id is %s.', response.getResponseText());
    // save key and refresh menu
    setOrgId(response.getResponseText());

    loadMenu();
  } else if (response.getSelectedButton() == ui.Button.CANCEL) {
    Logger.log('The user didn\'t want to provide an Org Id.');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}

function promptNet(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Network ID', 'Sets the network for "network-wide" reports. Run the org-wide "Networks" report to get this info.', ui.ButtonSet.OK_CANCEL);
  
  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.OK) {
    Logger.log('The user\'s Net Id is %s.', response.getResponseText());
    // save key and refresh menu
    setNetId(response.getResponseText());

    loadMenu();
  } else if (response.getSelectedButton() == ui.Button.CANCEL) {
    Logger.log('The user didn\'t want to provide an Net Id.');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}

function promptTimespan(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Timespan', 'Used by some reports to get range of data', ui.ButtonSet.OK_CANCEL);
  
  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.OK) {
    Logger.log('The user\'s timespan is %s.', response.getResponseText());
    // save key and refresh menu
    setTimespan(response.getResponseText());

    loadMenu();
  } else if (response.getSelectedButton() == ui.Button.CANCEL) {
    Logger.log('The user didn\'t want to provide a timespan');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}

function isObject(obj) {
  return obj === Object(obj);
}
  
function displayJSON(json, keys){
  if (!Array.isArray(json) || !keys.length) {
  // array does not exist, is not an array, or is empty
    Logger.log('displayJson did not receive json data');
    return;
  }
  if (!Array.isArray(keys) || !keys.length) {
  // array does not exist, is not an array, or is empty
    Logger.log('displayJson did not receive keys');
    return;
  }
  //json = [{"id":"1234","name":"sample"},{"id":"9876","name":"sample 2", "extra":"more info"}];
  Logger.log('displayJSON'+ JSON.stringify(json));
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheet = ss.getActiveSheet();
  var values = [];
  var row = sheet.getActiveCell().getLastRow();
  var column = sheet.getActiveCell().getColumn();
  var numRows = 1;
  var numColumns = 1;
  
  // Display Headers (keys)
    sheet.getRange(sheet.getActiveCell().getLastRow(), sheet.getActiveCell().getColumn(), 1, keys.length).setValues([keys]);
  
  // Parse JSON Object
  if(!Array.isArray(json)){
       
    // Get Values
    var v = [];    
    keys.forEach(
      function (k){
        v.push(json[k]);
      }
    );
    values.push(v);
    Logger.log('Parse Object values '+values);
    
  } else {
  
  // Parse JSON Array of Objects
    for (i = 0; i < json.length; i++) { 
      var data = json[i];
      
      // Get Values
      var v = [];
      keys.forEach(
        function (k){
          v.push(data[k]);
        }
      );
      values.push(v);
      Logger.log('Parse Array of Object values '+values);
    } 
  }
  
    // Display Data
    numRows = values.length;
    numColumns = keys.length;
  /*
    Logger.log('display numRows: '+ numRows);
    Logger.log('display numColumns: ' + numColumns);
    Logger.log('values '+values);
  */
    sheet.getRange(row+1,column,numRows,numColumns).setValues(values);

}

function flattenObject(ob) {
   var toReturn = {};
	
	for (var i in ob) {
		if (!ob.hasOwnProperty(i)) continue;
		
		if ((typeof ob[i]) == 'object') {
			var flatObject = flattenObject(ob[i]);
			for (var x in flatObject) {
				if (!flatObject.hasOwnProperty(x)) continue;
				
				toReturn[i + '.' + x] = flatObject[x];
			}
		} else {
			toReturn[i] = ob[i];
		}
	}
	return toReturn;
};

// **************************
// Meraki API 
// **************************

function getAdmins(apiKey, orgId) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/organizations/"+orgId+"/admins", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getOrgs(apiKey) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/organizations", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  data = data.replace(/([\[:])?(\d+)([,\}\]])/g, "$1\"$2\"$3");
  var json = JSON.parse(data);
  return json; 
}

function getOrg(apiKey, orgId) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/organizations/"+orgId, {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  data = data.replace(/([\[:])?(\d+)([,\}\]])/g, "$1\"$2\"$3");
  var json = JSON.parse(data)
  return json; 
}

function getNetworks(apiKey, orgId) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/organizations/"+orgId+"/networks", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getDevicesStatuses(apiKey, orgId) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/organizations/"+orgId+"/deviceStatuses", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getDevices(apiKey, netId) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/devices", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getDeviceLossAndLatencyHistory(apiKey, netId, serial, timespan) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/devices/"+serial+"/lossAndLatencyHistory?timespan="+timespan+"&ip=8.8.8.8&uplink=wan1", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getLicenseState(apiKey, orgId) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/organizations/"+orgId+"/licenseState", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json;
}

function getConfigTemplates(apiKey, orgId) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/organizations/"+orgId+"/configTemplates", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getClient(apiKey,netId, clientMac) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/clients/"+clientMac, {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getClients(apiKey,serial) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/devices/"+serial+"/clients?timespan="+settings.timespan, {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getInventory(apiKey, orgId) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/organizations/"+orgId+"/inventory", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getGroupPolicies(apiKey, netId) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/groupPolicies", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getSsids(apiKey, netId) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/ssids", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getStaticRoutes(apiKey, netId) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/staticRoutes", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getUplinkInfo(apiKey, netId, serial) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/devices/"+serial+"/uplink", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getVlans(apiKey, netId) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/vlans", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data); 
  return json; 
}

function getSiteToSiteVpn(apiKey, netId) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/siteToSiteVpn", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}


// Requires Meraki Network to be configured with Hostname Visibility to work
function getTraffic(apiKey, netId) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/traffic?timespan=7200", {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

// Wireless Health 
function getConnectionStats(apiKey, netId, t0, t1) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/connectionStats?t0="+t0+"&t1="+t1, {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getConnectionStatsByNode(apiKey, netId, t0, t1) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/devices/connectionStats?t0="+t0+"&t1="+t1, {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getConnectionStatsByClient(apiKey, netId, t0, t1) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/clients/connectionStats?t0="+t0+"&t1="+t1, {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getLatencyStats(apiKey, netId, t0, t1) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/latencyStats?t0="+t0+"&t1="+t1, {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getLatencyStatsByNode(apiKey, netId, t0, t1) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/devices/latencyStats?t0="+t0+"&t1="+t1, {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getLatencyStatsByClient(apiKey, netId, t0, t1) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/clients/latencyStats?t0="+t0+"&t1="+t1, {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}

function getFailedConnections(apiKey, netId, t0, t1) {
  var response = UrlFetchApp.fetch("https://api.meraki.com/api/v0/networks/"+netId+"/failedConnections?t0="+t0+"&t1="+t1, {headers:{'X-Cisco-Meraki-API-Key': apiKey}});
  var data = response.getContentText();
  var json = JSON.parse(data);
  return json; 
}


// **************************
// Reports 
// **************************

function callOrgs(){ 
  var data = []; 
  var keys = [];
  var result = getOrgs(settings.apiKey, settings.orgId); 
  
  if(!Array.isArray(result)){return}   
  result.forEach(function(obj){ 
    var flat = flattenObject(obj);
    data.push(flat);
    
    // set keys
    Object.keys(flat).forEach(function(value){
      if (keys.indexOf(value)==-1) keys.push(value);
    });
  });
 
  displayJSON(data,keys);
}

function callOrg(){ 
  var data = []; 
  var keys = [];
  var result = getOrg(settings.apiKey, settings.orgId);

  var flat = flattenObject(result);
  data.push(flat);
  
  // set keys
  Object.keys(flat).forEach(function(value){
    if (keys.indexOf(value)==-1) keys.push(value);
  });

  displayJSON(data,keys);
}

function callAdmins(){ 
  var data = []; 
  var keys = [];
  var result = getAdmins(settings.apiKey, settings.orgId); 
  
  if(!Array.isArray(result)){return}   
  result.forEach(function(obj){ 
    var flat = flattenObject(obj);
    data.push(flat);
    
    // set keys
    Object.keys(flat).forEach(function(value){
      if (keys.indexOf(value)==-1) keys.push(value);
    });
  });
 
  displayJSON(data,keys);
}

function callNetworks(){
  var data = []; 
  var keys = [];
  if(!settings.orgId){
    promptOrg();
  }
  var result = getNetworks(settings.apiKey, settings.orgId);
  
  if(!Array.isArray(result)){return}   
  result.forEach(function(obj){ 
    var flat = flattenObject(obj);
    data.push(flat);
    
    // set keys
    Object.keys(flat).forEach(function(value){
      if (keys.indexOf(value)==-1) keys.push(value);
    });
  });
 
  displayJSON(data,keys);
}

function callDevices(){
  var data = [];
  var keys = [];
  var nets = getNetworks(settings.apiKey, settings.orgId); 
  
  // Get Devices for each network in the organization
  for (var i = 0; i <= nets.length; i++){
    try{
      var result = getDevices(settings.apiKey, nets[i].id);
      if(!Array.isArray(result)){continue}
      
      result.forEach(function(obj){ 
        // flatten object and add network info
        var flat = flattenObject(obj);
        flat['networkId'] = nets[i].id;
        flat['networkName'] = nets[i].name;
        data.push(flat);
        
        // set keys
        Object.keys(flat).forEach(function(value){
          if (keys.indexOf(value)==-1) keys.push(value);
        });
      });
          
    }catch(e){
      Logger.log('error'+e);
      continue;
    }
  }
  displayJSON(data,keys);
}

function callDeviceLossAndLatencyHistory(){
  var data = [];
  var keys = [];
  var devices = getDevices(settings.apiKey, settings.netId); 
  
  // Get Devices for each network in the organization
  for (var i = 0; i <= devices.length; i++){
    try{
      var result = getDeviceLossAndLatencyHistory(settings.apiKey, settings.netId, devices[i].serial, settings.timespan);
      if(!Array.isArray(result)){continue}
      
      result.forEach(function(obj){ 
        // flatten object and add network info
        var flat = flattenObject(obj);
        flat['deviceSerial'] = devices[i].serial;
        flat['deviceName'] = devices[i].name;
        data.push(flat);
        
        // set keys
        keys = ['deviceName', 'deviceSerial'];
        Object.keys(flat).forEach(function(value){
          if (keys.indexOf(value)==-1) keys.push(value);
        });
      });
          
    }catch(e){
      Logger.log('error'+e);
      continue;
    }
  }
  displayJSON(data,keys);
}
function callDevicesStatuses(){
  var data = []; 
  var keys = [];
  var result = getDevicesStatuses(settings.apiKey, settings.orgId);
  
  if(!Array.isArray(result)){return}   
  result.forEach(function(obj){ 
    var flat = flattenObject(obj);
    data.push(flat);
    
    // set keys
    Object.keys(flat).forEach(function(value){
      if (keys.indexOf(value)==-1) keys.push(value);
    });
  });
	
displayJSON(data,keys);
 
}

function callInventory(){
  var data = []; 
  var keys = [];
  var result = getInventory(settings.apiKey, settings.orgId);
  
  if(!Array.isArray(result)){return}   
  result.forEach(function(obj){ 
    var flat = flattenObject(obj);
    data.push(flat);
    
    // set keys
    Object.keys(flat).forEach(function(value){
      if (keys.indexOf(value)==-1) keys.push(value);
    });
  });
 
  displayJSON(data,keys);
}
  
function callLicenseState(){
  var data = []; 
  var keys = [];
  var result = getLicenseState(settings.apiKey, settings.orgId);

  var flat = flattenObject(result);
  data.push(flat);
  
  keys = ['status','expirationDate'];

  displayJSON(data,keys);
}

function callLicenseStateForOrgs(){
  var data = []; 
  var keys = [];
  var orgs = getOrgs(settings.apiKey);
  

   // Get License State for each org for API key
  for (var i = 0; i <= orgs.length; i++){
    try{
      var license = getLicenseState(settings.apiKey, orgs[i].id);
      
      // flatten object and add org info
        var flat = flattenObject(license);
        flat['orgId'] = orgs[i].id;
        flat['orgName'] = orgs[i].name;
        data.push(flat);
      
      // set keys
      keys = ['orgName', 'orgId', 'status','expirationDate'];
      
      // device count details
        Object.keys(flat).forEach(function(value){
          if (keys.indexOf(value)==-1) keys.push(value);
        });
     
          
    }catch(e){
      Logger.log('error'+e);
      continue;
    }
  }
  displayJSON(data,keys);
}

function callLicenseStateDetails(){
  var data = []; 
  var keys = [];
  var result = getLicenseState(settings.apiKey, settings.orgId);

  var flat = flattenObject(result);
  data.push(flat);
  
  // set keys
  Object.keys(flat).forEach(function(value){
    if (keys.indexOf(value)==-1) keys.push(value);
  });

  displayJSON(data,keys);
}

function callConfigTemplates(){
  var data = []; 
  var keys = [];
  var result = getConfigTemplates(settings.apiKey, settings.orgId);
  
  if(!Array.isArray(result)){return}   
  result.forEach(function(obj){ 
    var flat = flattenObject(obj);
    data.push(flat);
    
    // set keys
    Object.keys(flat).forEach(function(value){
      if (keys.indexOf(value)==-1) keys.push(value);
    });
  });
 
  displayJSON(data,keys);
}

function callVlansOfOrg(){ 
  var data = [];
  var keys = [];
  var nets = getNetworks(settings.apiKey, settings.orgId); 
  
  // Get VLANs for each network in the organization
  for (var i = 0; i <= nets.length; i++){
    try{
      var result = getVlans(settings.apiKey, nets[i].id);
      if(!Array.isArray(result)){continue}
      
      result.forEach(function(obj){ 
        // flatten object and add network info
        var flat = flattenObject(obj);
        flat['networkId'] = nets[i].id;
        flat['networkName'] = nets[i].name;
        data.push(flat);
        
        // set keys
        Object.keys(flat).forEach(function(value){
          if (keys.indexOf(value)==-1) keys.push(value);
        });
      });
          
    }catch(e){
      Logger.log('error'+e);
      continue;
    }
  }
  displayJSON(data,keys);
}


function callTrafficOfOrg(){
  var data = [];
  var keys = [];
  var nets = getNetworks(settings.apiKey, settings.orgId); 
  
  // Get traffic analysis for each network in the organization
  for (var i = 0; i < nets.length; i++){
    try{
      var traffic = getTraffic(settings.apiKey, nets[i].id);
      
      // flatten object and add network info
      var flat = flattenObject(traffic);
      flat.networkId = nets[i].id;
      flat.networkName = nets[i].name;
      
      // set data
      data.push(flat);
      
      // set keys
      Object.keys(flat).forEach(function(value){
        if (keys.indexOf(value)==-1) keys.push(value);
      });
      
    }catch(e){
      Logger.log('error'+e);
      continue;
    }     
  }
  displayJSON(data, keys);
}


function callSiteToSiteVpn(){
  var data = [];
  var keys = [];
  var nets = getNetworks(settings.apiKey, settings.orgId);
 
  // Get vpn info for each network in the organization
  for (var i = 0; i < nets.length; i++){
    try{
      var vpn = getSiteToSiteVpn(settings.apiKey, nets[i].id);   
      
      // flatten object and add network info
      var flat = flattenObject(vpn);
      flat.networkId = nets[i].id;
      flat.networkName = nets[i].name;
      
      // set data
      data.push(flat);
      
      // set keys
      Object.keys(flat).forEach(function(value){
        if (keys.indexOf(value)==-1) keys.push(value);
      });
      
    }catch(e){
      Logger.log('error'+e);
      continue;
    }     
  }
  displayJSON(data, keys);
}


function callSsidsOfOrg(){
  var data = [];
  var keys = [];
  var nets = getNetworks(settings.apiKey, settings.orgId); 
  
  // Get SSIDs for each network in the organization
  for (var i = 0; i <= nets.length; i++){
    try{
      var ssids = getSsids(settings.apiKey, nets[i].id);
      if(!Array.isArray(ssids)){continue}
      
      ssids.forEach(function(obj){ 
        // flatten object and add network info
        var flat = flattenObject(obj);
        flat['networkId'] = nets[i].id;
        flat['networkName'] = nets[i].name;
        data.push(flat);
        
        // set keys
        Object.keys(flat).forEach(function(value){
          if (keys.indexOf(value)==-1) keys.push(value);
        });
      });
          
    }catch(e){
      Logger.log('error'+e);
      continue;
    }
  }
  displayJSON(data,keys);
}



function callGroupPoliciesOfOrg(){
  var data = [];
  var keys = [];
  var nets = getNetworks(settings.apiKey, settings.orgId); 

  // Get Policies for each network in the organization
  for (var i = 0; i < nets.length; i++){
    try{
      var policies = getGroupPolicies(settings.apiKey, nets[i].id);    
      if(!Array.isArray(policies)){continue}
      
      policies.forEach(function(obj){ 
        // flatten object and add network info
        var flat = flattenObject(obj);
        flat['networkId'] = nets[i].id;
        flat['networkName'] = nets[i].name;
        data.push(flat);
        
        // set keys
        Object.keys(flat).forEach(function(value){
          if (keys.indexOf(value)==-1) keys.push(value);
        });
      });
          
    }catch(e){
      Logger.log('error'+e);
      continue;
    }
  }
  displayJSON(data,keys);
}

function callClientsOfOrg(){
  var data = [];
  var keys = [];
  const devices = getDevicesStatuses(settings.apiKey, settings.orgId);
  const nets = getNetworks(settings.apiKey, settings.orgId);
  
  // Get clients for each device in the organization
  for (var i = 0; i < devices.length; i++){
    try{
      var clients = getClients(settings.apiKey, devices[i].serial);
      
      if(!Array.isArray(clients)){continue}
      
      clients.forEach(function(obj){  
        // flatten object and add network info
        var flat = flattenObject(obj);
        flat.deviceName = devices[i].name;
        flat.deviceLanIp = devices[i].lanIp;
        flat.deviceWan1IP = devices[i].wan1Ip;
        flat.deviceWan2IP = devices[i].wan2Ip;
        flat.deviceMac = devices[i].mac;
        flat.deviceName = devices[i].name;
        flat.networkId = devices[i].networkId;    
        flat.networkName = nets.filter(function(obj){
           return devices[i].networkId == obj['id'];
         })[0]['name'];

        data.push(flat);
        
        // set keys
        keys = ['description','dhcpHostname','mac','ip','usage.sent','usage.recv','id','networkId','networkName']
        Object.keys(flat).forEach(function(value){
          if (keys.indexOf(value)==-1) keys.push(value);
        });       
      });
    } catch(e){
        Logger.log('error'+e);
        continue;
    }  
  }
  displayJSON(data,keys);
}

function callClientsOfOrgDetails(){
  var data = [];
  var keys = [];
  const devices = getDevicesStatuses(settings.apiKey, settings.orgId);
  const nets = getNetworks(settings.apiKey, settings.orgId);
  
  // Get clients for each device in the organization
  for (var i = 0; i < devices.length; i++){
    try{
      var clients = getClients(settings.apiKey, devices[i].serial);
      
      if(!Array.isArray(clients)){continue}
      
      clients.forEach(function(obj){  
        // flatten object
        var flat = flattenObject(obj);
        
        // add network info
        flat.deviceName = devices[i].name;
        flat.deviceLanIp = devices[i].lanIp;
        flat.deviceWan1IP = devices[i].wan1Ip;
        flat.deviceWan2IP = devices[i].wan2Ip;
        flat.deviceMac = devices[i].mac;
        flat.deviceName = devices[i].name;
        flat.networkId = devices[i].networkId;    
        flat.networkName = nets.filter(function(obj){
           return devices[i].networkId == obj['id'];
         })[0]['name'];

        // add client info
        var client = getClient(settings.apiKey, devices[i].networkId, obj.mac);
        for (var attrname in client) { flat[attrname] = client[attrname]; }
        
        data.push(flat);
        
        // set keys
        keys = ['description','dhcpHostname','mac','ip','usage.sent','usage.recv','id','networkId','networkName']
        Object.keys(flat).forEach(function(value){
          if (keys.indexOf(value)==-1) keys.push(value);
        });       
      });
    } catch(e){
        Logger.log('error'+e);
        continue;
    }  
  }
  displayJSON(data,keys);
}

function callClientsOfNet(){
  var data = [];
  var keys = [];
  var deviceStatuses = getDevicesStatuses(settings.apiKey, settings.orgId);

  var devices = deviceStatuses.filter(function(d){
    return d.networkId == settings.netId;
  });

  // Get clients for each device in the organization
  for (var i = 0; i < devices.length; i++){
    try{
      var clients = getClients(settings.apiKey, devices[i].serial);
      
      if(!Array.isArray(clients)){continue}
      
      clients.forEach(function(obj){  
        // flatten object and add network info
        var flat = flattenObject(obj);
        flat.deviceName = devices[i].name;
        flat.deviceLanIp = devices[i].lanIp;
        flat.deviceWan1IP = devices[i].wan1Ip;
        flat.deviceWan2IP = devices[i].wan2Ip;
        flat.deviceMac = devices[i].mac;
        flat.deviceName = devices[i].name;
        flat.networkId = devices[i].networkId;   
        data.push(flat);
        
        // set keys
        keys = ['description','dhcpHostname','mac','ip','usage.sent','usage.recv','id']
        Object.keys(flat).forEach(function(value){
          if (keys.indexOf(value)==-1) keys.push(value);
        });       
      });
    } catch(e){
        Logger.log('error'+e);
        continue;
    }  
  }
  displayJSON(data,keys);
}

function callClientsOfNetDetails(){
  var data = [];
  var keys = [];
  var deviceStatuses = getDevicesStatuses(settings.apiKey, settings.orgId);

  var devices = deviceStatuses.filter(function(d){
    return d.networkId == settings.netId;
  });

  // Get clients for each device in the organization
  for (var i = 0; i < devices.length; i++){
    try{
      var clients = getClients(settings.apiKey, devices[i].serial);
      // Logger.log(JSON.parse(clients));
      if(!Array.isArray(clients)){continue}
      
      clients.forEach(function(obj){  
        var client = getClient(settings.apiKey, devices[i].networkId, obj.mac);
        Logger.log(client);
        // flatten object 
        var flat = flattenObject(obj);
        
        // Copy device info
        flat.deviceName = devices[i].name;
        flat.deviceLanIp = devices[i].lanIp;
        flat.deviceWan1IP = devices[i].wan1Ip;
        flat.deviceWan2IP = devices[i].wan2Ip;
        flat.deviceMac = devices[i].mac;
        flat.deviceName = devices[i].name;
        flat.networkId = devices[i].networkId; 
        
        // Copy over all client details
        for (var attrname in client) { flat[attrname] = client[attrname]; }
        

        Logger.log("flat"+JSON.stringify(flat));
        data.push(flat);
        
        // set keys
        keys = ['description','dhcpHostname','mac','ip','usage.sent','usage.recv','id']
        Object.keys(flat).forEach(function(value){
          if (keys.indexOf(value)==-1) keys.push(value);
        });       
      });
    } catch(e){
        Logger.log('error'+e);
        continue;
    }  
  }
  displayJSON(data,keys);
}

function callStaticRoutes(){
  var data = [];
  var keys = [];
  var nets = getNetworks(settings.apiKey, settings.orgId); 
  
  // Get routes for each network in the organization
  for (var i = 0; i <= nets.length; i++){
    try{
      var routes = getStaticRoutes(settings.apiKey, nets[i].id);
      
      if(!Array.isArray(routes)){continue}
      routes.forEach(function(obj){ 
        // flatten object and add network info
        var flat = flattenObject(obj);
        flat['networkId'] = nets[i].id;
        flat['networkName'] = nets[i].name;
        data.push(flat);
        
        // set keys
        Object.keys(flat).forEach(function(value){
          if (keys.indexOf(value)==-1) keys.push(value);
        });
      });
          
    }catch(e){
      Logger.log('error'+e);
      continue;
    }
  }
  displayJSON(data,keys);
}



function callConnectionStatsOfOrg(){
  var data = [];
  var keys = [];
  var nets = getNetworks(settings.apiKey, settings.orgId); 
  /*
  var nets = [
    {
      "name":"test",
      "id":"L_643451796760561218"
    }
  ];
  */
  var t1 = Math.floor((new Date).getTime() / 1000);
  var t0 = t1 - settings.timespan;
  
  // Get routes for each network in the organization
  for (var i = 0; i <= nets.length; i++){
    try{
      var stats = getConnectionStats(settings.apiKey, nets[i].id, t0, t1);
      Logger.log('stats ' + JSON.stringify(stats));
      typeof val === 'object'
      if(!isObject(stats)){
        stats = {};
        //continue;
      }
      
      
      // flatten object and add network info
      var flat = flattenObject(stats);
      flat['networkId'] = nets[i].id;
      flat['networkName'] = nets[i].name;
      Logger.log('stats parsed ' + JSON.stringify(flat));
      data.push(flat);
      
      // set keys
      Object.keys(flat).forEach(function(value){
        if (keys.indexOf(value)==-1) keys.push(value);
      });
 
          
    }catch(e){
      Logger.log('error'+e);
      continue;
    }
  }
  displayJSON(data,keys);
}

function callLatencyStatsOfOrg(){
  var data = [];
  var keys = [];
  var nets = getNetworks(settings.apiKey, settings.orgId); 
  /*
  var nets = [
    {
      "name":"test",
      "id":"L_643451796760561218"
    }
  ];
  */
  var t1 = Math.floor((new Date).getTime() / 1000);
  var t0 = t1 - settings.timespan;
  
  // Get routes for each network in the organization
  for (var i = 0; i <= nets.length; i++){
    try{
      var stats = getLatencyStats(settings.apiKey, nets[i].id, t0, t1);
      Logger.log('stats ' + JSON.stringify(stats));
      typeof val === 'object'
      if(!isObject(stats)){
        stats = {};
        //continue;
      }
        
      // flatten object and add network info
      var flat = flattenObject(stats);
      flat['networkId'] = nets[i].id;
      flat['networkName'] = nets[i].name;
      Logger.log('stats parsed ' + JSON.stringify(flat));
      data.push(flat);
      
      // set keys
      Object.keys(flat).forEach(function(value){
        if (keys.indexOf(value)==-1) keys.push(value);
      });
 
          
    }catch(e){
      Logger.log('error'+e);
      continue;
    }
  }
  displayJSON(data,keys);
}

function callConnectionStatsByNode(){
  var data = [];
  var keys = []; 
  
  var t1 = Math.floor((new Date).getTime() / 1000);
  var t0 = t1 - settings.timespan;
 
  var stats = getConnectionStatsByNode(settings.apiKey, settings.netId, t0, t1);
 
  stats.forEach(function(s) {   
      // flatten object and add network info
      var flat = flattenObject(s);
      flat['networkId'] = settings.netId;
      data.push(flat);
  
      // set keys
      Object.keys(flat).forEach(function(value){
        if (keys.indexOf(value)==-1) keys.push(value);
      });
   });
  
  displayJSON(data,keys);
}

function callConnectionStatsByClient(){
  var data = [];
  var keys = []; 
  
  var t1 = Math.floor((new Date).getTime() / 1000);
  var t0 = t1 - settings.timespan;
 
  var stats = getConnectionStatsByClient(settings.apiKey, settings.netId, t0, t1);
 
  stats.forEach(function(s) {   
      // flatten object and add network info
      var flat = flattenObject(s);
      flat['networkId'] = settings.netId;
      data.push(flat);
  
      // set keys
      Object.keys(flat).forEach(function(value){
        if (keys.indexOf(value)==-1) keys.push(value);
      });
   });
  
  displayJSON(data,keys);
}

function callLatencyStatsByNode(){
  var data = [];
  var keys = []; 
  
  var t1 = Math.floor((new Date).getTime() / 1000);
  var t0 = t1 - settings.timespan;
 
  var stats = getLatencyStatsByNode(settings.apiKey, settings.netId, t0, t1);
 
  stats.forEach(function(s) {   
      // flatten object and add network info
      var flat = flattenObject(s);
      flat['networkId'] = settings.netId;
      data.push(flat);
  
      // set keys
      Object.keys(flat).forEach(function(value){
        if (keys.indexOf(value)==-1) keys.push(value);
      });
   });
  
  displayJSON(data,keys);
}

function callLatencyStatsByClient(){
  var data = [];
  var keys = []; 
  
  var t1 = Math.floor((new Date).getTime() / 1000);
  var t0 = t1 - settings.timespan;
 
  var stats = getLatencyStatsByClient(settings.apiKey, settings.netId, t0, t1);
 
  stats.forEach(function(s) {   
      // flatten object and add network info
      var flat = flattenObject(s);
      flat['networkId'] = settings.netId;
      data.push(flat);
  
      // set keys
      Object.keys(flat).forEach(function(value){
        if (keys.indexOf(value)==-1) keys.push(value);
      });
   });
  
  displayJSON(data,keys);
}

function callFailedConnections(){
  var data = [];
  var keys = []; 
  
  var t1 = Math.floor((new Date).getTime() / 1000);
  var t0 = t1 - settings.timespan;
 
  var stats = getFailedConnections(settings.apiKey, settings.netId, t0, t1);
 
  stats.forEach(function(s) {   
      // flatten object and add network info
      var flat = flattenObject(s);
      flat['networkId'] = settings.netId;
      data.push(flat);
  
      // set keys
      keys = ['clientMac'];
      Object.keys(flat).forEach(function(value){
        if (keys.indexOf(value)==-1) keys.push(value);
      });
   });
  
  displayJSON(data,keys);
}

function callUplinkInfos(){
  var data = [];
  var keys = [];
  var devices = getDevicesStatuses(settings.apiKey, settings.orgId);

  // Get uplink info for each device in the organization
  for (var i = 0; i < devices.length; i++){
    try{
      var uplink = getUplinkInfo(settings.apiKey, devices[i].networkId, devices[i].serial);

      // flatten object and add network info
      var flat = flattenObject(uplink);
      flat.networkId = devices[i].networkId;
      flat.serial = devices[i].serial;
      flat.mac = devices[i].mac;
      flat.deviceName = devices[i].name;
      flat.status = devices[i].status;
      
      // set data
      data.push(flat);
      
      // set keys
      keys = ['networkId','serial','mac','deviceName','status']
      Object.keys(flat).forEach(function(value){
        if (keys.indexOf(value)==-1) keys.push(value);
      });
      
    }catch(e){
      Logger.log('error'+e);
      continue;
    }     
  }
  displayJSON(data, keys);
}