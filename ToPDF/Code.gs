
/**
 * MergeData
 * 
 * Google Script to merge data from a Google Sheet (dataID) into a Google Doc template (TemplateID) convert it to a PDF document 
 * andn put in into a directory.(outputPdfFolderId)
 * 
 * Change dataID, TemplateID and outputPdfFolderId to match yours and then run createPdf()
 * 
 * Update:
 *   - Create a JSON array cols to match fieldname and pos in data sheet. 
 *   - Add a .pdf extention to output pdf file
 * 
 * Version 1.2 
 * Date: 7 Juli 2022
 * 
 * By: Jeff Schilders (j.schilders@zeven-linden.nl | jeff@schilders.com)
 */



/** Data Sheet doc ID */
var dataID = '1c_XcP4wRNlgdisQgdVhJqNhhmB2_cfBVXWYKHS1MN-k';
/** Template doc ID */
var TemplateID = '1chrkOy3Xq5W0HD0ZbuvWmd2x7amYkeUpABfMQ9Bv750'; 
/** Output Directory id */
var outputPdfFolderId = '1CanQMujwmaByRlGCw9vwqo5094miqThD';

/** %fieldname% in Google Template doc : col pos in Google data sheet. */
var cols = [{"school":0,
             "leerlingnummer":4,
             "leerlingnaam":5,
             "advies":6,
             "onderwijsniveau":7,
             "stamgroep":2,
             "niveau":1,
             "eindtoets":8,
             "IQ":9,
             "spelling":10,
             "technische":11,
             "begrijpend":12,
             "rekenen":13,
             "dyslexie":14,
             "dyscalculie":15,
             "bevorderendefactoren":16,
             "belemmerendefactoren":17,
             "toelichting":18,
             "ondersteuning":19,
             "aandachtspunten":20,
             "warmeoverdracht":21,
             "informatieoverdracht":22
            }];


  function testme(){
    Logger.log(cols[0]['leerlingnummer']);
  }

  function createPdf() {
   var folder = DriveApp.getFolderById(outputPdfFolderId);

   var ss = SpreadsheetApp.openById(dataID);
   var data = ss.getDataRange().getValues();

  /** Walk data rows form data sheet, skip first row */
  for (var i = 1; i < data.length; i++) {

   /** Create a copy of the Template file */
   var copyFile = DriveApp.getFileById(TemplateID).makeCopy(),
      copyId = copyFile.getId(),
      copyDoc = DocumentApp.openById(copyId),
      copyBody = copyDoc.getActiveSection(); 

    /** Define new filename for output pdf file based on leerlingnummer en leerlingnaam */
    var leerlingnummer=data[i][cols[0]['leerlingnummer']];
    var leerlingnaam=data[i][cols[0]['leerlingnaam']];
    var PDF_FILE_NAME = leerlingnummer+"_"+leerlingnaam+".pdf";
    
    /** Search and replace fields based on cols array */
    for(var key in cols[0]){
        var sub_key = key;
        var sub_val = cols[0][key];
        copyBody.replaceText('%' + sub_key + '%', data[i][sub_val]);
    }
    /** Save and close the copy of template */
    copyDoc.saveAndClose();

    /** Create a PDF document from Google Docs and move to output directory */
    if (PDF_FILE_NAME !== '') {
      var newFile = DriveApp.createFile(copyFile.getAs('application/pdf'));
      newFile.setName(PDF_FILE_NAME);
      newFile.moveTo(folder);
      console.log(i+" save "+PDF_FILE_NAME);
    } 

    /** Remove the template copy */
    copyFile.setTrashed(true)   

  }

}