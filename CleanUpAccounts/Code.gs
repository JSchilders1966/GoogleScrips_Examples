/**
 * CleanUpAccounts.
 * 
 * Code and design by Jeff Schilders <jeff@schilders.com>
 * 
 * Birthday of this script: 24 Maart 2011
 * 
 * YOU NEED TO CHANGE THIS SCRIPT BEFORE YOU CAN USE IT!!!!
 * 
 * Accounts that are left in last year students OU wil be moved to quarantaine.
 * When moved a messages will be send with instukties to move data from Drive to personal drive   
 * 
 * 
 * Contact me <jeff@schilders.com>, if you need help to change this script to the needs of your school,
 * Last edit: April 2022
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 *
 **/
var emailDomain="zeven-linden.nl";
var baseOrgUnit="zeven-linden.nl/leerlingen";
var mailonderwerp="Je schoolwerk veilig stellen";

var schoolyear;
var searchQuery;
var classname;
var DEBUG = false;


function cleanup() { 
  schoolyear=getLastSchoolYear_();
  var ParentOrg=""+baseOrgUnit+"/"+schoolyear; 
  var page = AdminDirectory.Orgunits.list('my_customer', {
     orgUnitPath: ParentOrg,
     viewType:'domain_public',
     type: 'all'
  });
  var units = page.organizationUnits;
  if (units) {
    for (var i = 0; i < units.length; i++) {
      var unit = units[i];
      Logger.log(unit.name);
      //getStudents_(unit.name); 
      liststudents(unit.name);
    }
  } else {
    Logger.log('No units found.');
  }
  
}


function liststudents(classname){

  let page = '';
  let pageToken = '';
  let leerlingen = [];
  
  schoolyear=getLastSchoolYear_();
  var CheckOrgPath=baseOrgUnit+"/"+schoolyear;
  var QuarantainePath=baseOrgUnit+"/quarantaine";
  var searchQuery = 'orgUnitPath=/'+CheckOrgPath+"/"+classname;

  do{
    page = AdminDirectory.Users.list({        
        customer: 'my_customer',
        query:searchQuery,
        pageToken: pageToken
      })
      
    if (page.users) {
      // Store only selected info for each user.
      let leerling = page.users.map(user => {
       return {
        'firstName': user.name.givenName,
        'lastName' : user.name.familyName,
        'email'    : user.primaryEmail
        }
      })
      leerlingen = leerlingen.concat(leerling)
    }
    pageToken = page.nextPageToken;
  
  } while(pageToken)
 
    for (var i = 0; i < leerlingen.length; i++) {
      console.log(leerlingen[i].email)
      var UserParam = {
          primaryEmail:leerlingen[i].email,
          orgUnitPath:'/'+QuarantainePath,  
      }
      updateUser_(UserParam,leerlingen[i].email);
      /** Stuur mail dat account wordt verwijderd */
      var leerling =
      {
        firstname:leerlingen[i].firstName,
        lastname:leerlingen[i].lastName,
        email:leerlingen[i].email,        
      }
      informUser_(leerling);
    }
    console.log('stop');

    
  
}

function informUser_(data){
  var templ = HtmlService
      .createTemplateFromFile('mailtemplate');
  templ.data = data;
  var message = templ.evaluate().getContent(); 
  if(!DEBUG){
    MailApp.sendEmail({
      to: data.email,
      subject: mailonderwerp,
      htmlBody: message
    });
  } else {
    console.log(message);
  }  
  Logger.log('Send to:'+data.email)
}

function updateUser_(UserParam,email){
     Logger.log('Update '+ email);
     if(!DEBUG){
       try {
          var user = AdminDirectory.Users.update(UserParam,email);
        } catch (e) {
          Logger.log(UserParam);
          Logger.log('Error:'+e.message);
          
       }
     }  
    return;
} 

function getStudents_(classname) {
  schoolyear=getLastSchoolYear_();

  var CheckOrgPath=baseOrgUnit+"/Leerlingen/"+schoolyear;
  var searchQuery = 'orgUnitPath=/'+CheckOrgPath+"/"+classname;

  console.log(searchQuery);

  const  students = AdminDirectory.Users.list({
      customer: 'my_customer',
      query:searchQuery,
  })
  console.log(students.lenght);
  if (students.length > 0 ) {
    for (var i = 0; i < students.length; i++) {
      console.log(students.users[i].primaryEmail);
      
    }
  }  


}


function getLastSchoolYear_(){
  var date = new Date();
  var mt = date.getMonth();
  var ye = date.getFullYear();

  if (mt >= 7) { // Aug
    var y = ye - 1;
    var schoolyear = y+"-"+ye;
  } else {
    var y = ye - 2;
    var ye = ye -1;
    var schoolyear = y+"-"+ye;

  }  
  return schoolyear;
} 
