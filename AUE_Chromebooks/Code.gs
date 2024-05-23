/**
 * 
 *  Het veranderen van de OU waarin een Chromebook zit.
 * 
 *  
 * Code and design by Jeff Schilders <jeff@schilders.com>
 * 
 * 
 */

var OrgPathLock = "/Devices/Chromebooks/Lockdown";

const CheckPaths = ["/Devices/Chromebooks/Lockdown", 
                    "/Devices/DigitalSignage/GoogleSign",
                    "/Devices/Afgeschreven",
                    "/Devices/Afgeschreven/Defect"
                   ];
var dateNow = new Date();

function checkExpiration() {

  var sheetID="16rhkBZY9R3NN5YS8iFQd704Ri2PEMXtZi9Kffk4dMYU";
  var sheetName="aue";
  var pageToken,page
  let counter=0;
  let oldchromebooks=0
  
  var lastuser;
  var header=[];

  
  var target_sheet=SpreadsheetApp.openById(sheetID).getSheetByName(sheetName)
  target_sheet.clearContents();
  header.push('ItemID','Serialnumber','Laatste gebruiker','Laatst online','AUE Datum','orgUnit');
  target_sheet.appendRow(header);
   
  range = target_sheet.getRange(1, 1, 1, 6);
  range.setBackground("#E6E6E6");
  range.setFontSize(12).setFontWeight("bold")
 
  do {
    page = AdminDirectory.Chromeosdevices.list("my_customer",
       {maxResults: 500, pageToken: pageToken,"query": "status: ACTIVE"}
    )
    var devices = page.chromeosdevices;

    for (device in devices) {
      counter=counter+1
      var expirationDate=devices[device].autoUpdateExpiration
      
      var date = new Date(expirationDate*1);
      
      if(differenceInMonths(dateNow,date) > 6 ){
        try{
          lastuser=devices[device].recentUsers[0].email;
        } catch (e) {
          lastuser="";
        }
        
        try {
           lastused=Utilities.formatDate(new Date(devices[device].activeTimeRanges[0].date),'CET+1','dd-MM-YYYY');
        } catch (e) {
          lastused="";
        }

        auedate=Utilities.formatDate(new Date(date),'CET+1','MM-YYYY');
      
        var row = [];
        row.push(devices[device].annotatedAssetId, devices[device].serialNumber,lastuser,lastused,auedate,devices[device].orgUnitPath)
        target_sheet.appendRow(row);

        console.log(devices[device].serialNumber)
        oldchromebooks=oldchromebooks+1;
       
        updateChromebookState_(devices[device].deviceId,devices[device].orgUnitPath);
       
      }

    }  
    pageToken = page.nextPageToken;
  } while (pageToken)
    console.log("Checked "+counter+" Active Chromebooks");
    console.log("Disabled "+oldchromebooks+" Chromebooks");
}


function differenceInMonths(date1, date2) {
  const monthDiff = date1.getMonth() - date2.getMonth();
  const yearDiff = date1.getYear() - date2.getYear();
  return monthDiff + yearDiff * 12;
}


function test(){
  updateChromebookState_("40f35fee-227f-41ef-8ccf-c3c4467c96e7","");
}

function updateChromebookState_(did,path){
   
  if (!CheckPaths.includes(path)){

    let item= AdminDirectory.Chromeosdevices.get("my_customer",did)
    
    var expirationDate=item.autoUpdateExpiration
    var date = new Date(expirationDate*1);   
    
    if(differenceInMonths(dateNow,date) > 6 ){
      var DeviceData = {orgUnitPath: OrgPathLock};
      try {
        var device = AdminDirectory.Chromeosdevices.update(DeviceData,"my_customer",did);     
        console.log("Move: "+item.serialNumber+" To "+OrgPathLock);
      } catch (e) {
        Logger.log('Error:'+e.message);
      } 
   }  
  }
}