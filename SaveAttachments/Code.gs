const ValidUsers = ["j.schilders@domain.lt","j.dow@domain.tl"]; // Who can send files 
const ValidBrin = ["19VD","19QK"]; // SO 19VD  VSO 19QK

const JsonFolderId="xxxxxxxxxxxxxxxxxxxxxxxxxx"; // Folder where to save the json file
const JsonBackupFolderId="xxxxxxxxxxxxxxxxxx"; // Where to move the json file when ready
const SavefolderId = "xxxxxxxxxxx"; // Where to save the attachment
function saveNewAttachmentsToDrive() {

 
  var searchQuery = "from:@ooz.nl has:attachment filename: Vrije export csv"; // Replace with the search query to find emails with attachments

  var lastExecutionTime = getLastExecutionDate();
  var threads = GmailApp.search(searchQuery + " after:" + lastExecutionTime);
  var driveFolder = DriveApp.getFolderById(SavefolderId);
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      
      var message = messages[j];
      
      if(message.isUnread()){
        console.log(message.getId())
        var attachments = message.getAttachments();
        
        var sender=message.getFrom().replace(/^.+<([^>]+)>$/, "$1");
        console.log(sender);

        if ( ValidUsers.includes(sender)) {
          var Brin=message.getSubject().toUpperCase().substring(5,9);
          if (ValidBrin.includes(Brin)){
            for (var k = 0; k < attachments.length; k++) {
              var attachment = attachments[k];
              var attachmentBlob = attachment.copyBlob();
              var fileName = attachment.getName();
              if (fileName.indexOf('Vrije_export') > -1) {
                var newfile=driveFolder.createFile(attachmentBlob).setName(fileName);
                 console.log(newfile.getId())  
                 var status=createJsonFile(newfile.getId(),Brin);
              } else {
               console.log("Not a valid file name:"+fileName);
               status="wrong";
              }  
           }
          } else {
            console.log('Not a valid Brin');
            status="wrong";
          }
        } else {
          console.log('Not a valid user');
          status="wrong";
        }
        
        switch(status) {
          case "ok":
            message.moveToTrash();
            mailBody="<p>Beste, </p>"+
                     "<p>Het bestand is opgeslagen en de leerlingen gegevens voor <b>"+Brin+"</b> zullen in de loop van de dag verwerkt worden door het systeem.</p>"+
                     "<p>Dit is een automatisch gegenereerd bericht.<br>Onderwijscentrum de Twijn</p>";
          break;
          case "wrong":
            message.moveToTrash();
            mailBody="<p>Beste, </p>"+
                     "<p>Er is een fout opgetreden.</p>"+
                     "<p>Het systeem heeft ("+Brin+") herkent in het onderwerp, deze Brincode of de bestand indeling zijn mischien verkeerd aangeleverd!</p>"+
                     "<p>Dit is een automatisch gegenereerd bericht.<br>Onderwijscentrum de Twijn</p>";
          break;
          case "wait":
            newfile.setTrashed(true);
            mailBody="<p>Beste, </p>"+
                     "<p>Het bestand is ontvangen voor <b>"+Brin+"</b>  maar er loopt nog een eerder aanvraag.</p>"+
                     "<p>Het systeem zal het op een later moment nog eens proberen, u hoeft verder geen actie te ondernemen.</p>"+
                     "<p>Je ontvangt op een later moment een mailtje al het geleukt is.</p>"+
                     "<p>Dit is een automatisch gegenereerd bericht.<br>Onderwijscentrum de Twijn</p>";
          break;
        }

        MailApp.sendEmail({
          to: sender,
          subject: 'GSuite_Sync Status - ' + new Date(),
          htmlBody: mailBody,   
        });

      }
      
    }
  }


  updateLastExecutionDate();
}



function createJsonFile(InPutFileId,brin) {
 
 var fileName = "students.json";
  var leerlingnummer = 0;
  var json=[];
  var controle=0

  var Folder = DriveApp.getFolderById(JsonFolderId);
  var files = Folder.getFilesByName(fileName);  
  if (files.hasNext()) {
    console.log("Er staat al een bestand klaar !!");
    return "wait";
  }
  var inFile = DriveApp.getFileById(InPutFileId);
  var filedata = inFile.getBlob().getDataAsString();  
  var csvdata = Utilities.parseCsv(filedata, ';');
  
  for (var i = 0; i < csvdata[0].length; i++) {
    if (csvdata[0][i] == "Leerlingnummer"){leerlingnummer = i;controle=controle+1;}
    if (csvdata[0][i] == "Roepnaam"){roepnaam = i;controle=controle+1;}
    if (csvdata[0][i] == "Tussenvoegsel"){voegsel = i;controle=controle+1;}
    if (csvdata[0][i] == "Achternaam"){achternaam = i;controle=controle+1;}
    if (csvdata[0][i] == "Groepsnaam huidige groepsindeling"){stamgroep = i;controle=controle+1;}
    if (csvdata[0][i] == "BSN"){bsn = i;controle=controle+1;}
  }  
  if(leerlingnummer < 0  || roepnaam < 0 || voegsel < 0 || achternaam < 0 || stamgroep < 0 || bsn < 0 ) {
   console.log("error:'Verkeerde bestandsindeling'");
   return "wrong";
  }  
  
  for (var i = 1; i < csvdata.length; i++) {            
     var groepOU=csvdata[i][stamgroep]; 
     if(groepOU == ''){groepOU="NAN"};
     
     var voornaam = csvdata[i][roepnaam]; 
     var bsnnumber = csvdata[i][bsn];

     if( csvdata[i][voegsel] != '' ){
      var famnaam = csvdata[i][voegsel]+' '+csvdata[i][achternaam];
     } else {
      var famnaam = csvdata[i][achternaam];
     }  
  
    json.push({nummer:csvdata[i][leerlingnummer],bsn:bsnnumber,voornaam:voornaam,achternaam:famnaam,groep:groepOU});
  }
  
  var myJSON = JSON.stringify({school:brin,leerlingen:json});
  var Folder = DriveApp.getFolderById(JsonFolderId);
  Folder.createFile('students.json',myJSON,MimeType.PLAIN_TEXT);
  inFile.setTrashed(true);
  return "ok";    
}



function getLastExecutionDate() {
  var properties = PropertiesService.getUserProperties();
  return properties.getProperty("lastExecutionDate") || "2023-09-01";
}

function resetLastExecutionDate() {
  PropertiesService.getUserProperties().deleteProperty("lastExecutionDate");
}

function updateLastExecutionDate() {
  var now = new Date();
  var dateString = now.toISOString().split("T")[0];
  var properties = PropertiesService.getUserProperties();
  properties.setProperty("lastExecutionDate", dateString);
}
