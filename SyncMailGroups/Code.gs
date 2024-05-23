/**
 * Update Maillist from classes.
 * 
 * Code and design by Jeff Schilders <jeff@schilders.com>
 * 
 * Contact me <jeff@schilders.com>, if you need help to change this script to the needs of your school,
 * 
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
var DEBUG = false;
var emailDomain="zeven-linden.nl";
var baseOrgUnit="zeven-linden.nl";
var schooljaar=getCurrentSchoolYear_();
var accounts = [];
var groupmembers = [];
var groupmail;
var NOMEMBER=true;
var mailBody="<h2>G-Suite Mailgroups Status Raport:</h2><br/>";
var mailAdd="";
var mailRemove="";
var mailError="";


function test(){
 //createMailGroup_('1hl@klassen.detwijn.nl',"1H");
 editGroup_('',"");
 
}

function main(){
  var ParentOrg="/"+baseOrgUnit+"/leerlingen/"+schooljaar; 
  Logger.log(ParentOrg);
  var page = AdminDirectory.Orgunits.list('my_customer', {
     orgUnitPath: ParentOrg,
     type: 'all'
  });
  var units = page.organizationUnits;
  if (units) {
    for (var i = 0; i < units.length; i++) {
      var unit = units[i];
      
      var groupmail='l'+unit.name.toLowerCase().slice(1)+'@'+emailDomain;

      if (checkGroup_(groupmail,unit.name.toUpperCase()))
      { 
         editGroup_(groupmail,unit.name.toUpperCase()); // Update Settings for Group.
         accounts.length =0;
         groupmembers.length =0;
         accounts = getUsers_(ParentOrg+"/"+unit.name.toLowerCase());
        
         // Check OU users is a member of the mailgroup OU
        for (var s = 0; s < accounts.length; s++) 
        {
            if(DEBUG) {
              // console.log("Check if "+accounts[s].email.toString()+" in group "+groupmail);
            }
            if(!AdminDirectory.Members.hasMember(groupmail,accounts[s].email.toString()).isMember)
            {
              addUsertoGroup_(accounts[s].email.toString(),groupmail);
            }
        }
         
        groupmembers=getMembers_(groupmail);
         // Check OU users is a member of the mailgroup OU
        for (var m = 0; m < groupmembers.length; m++) 
        {
          ISMEMBER=false;
          for (var s = 0; s < accounts.length; s++) 
          {
            if (accounts[s].email.toString() == groupmembers[m][1].toString())
            {
              ISMEMBER=true;
            } 
          }
          if(!ISMEMBER)
          {
            // Remove groep user to this group
            removeUserfromGroup_(groupmembers[m][1].toString(),groupmail)
          } 
        }                   
      } 
      else 
      {
         Logger.log(unit.name.toLowerCase()+" Is not a valid mailgroup!!!!");
         mailError=mailError+unit.name.toLowerCase()+" is not a valid mailgroup!!!<br/>"
      }  
    }
  }
  mailBody="<b>Add</b><br/>"+mailAdd+"<br/><b>Remove</b><br/>"+mailRemove+"<br/><b>Error</b><br/>"+mailError+"<br/>";
  MailApp.sendEmail({
         to: Session.getActiveUser().getEmail(),
         subject: 'GSuite_Sync Mailgroups - ' + new Date(),
         htmlBody: mailBody,   
  });
  Logger.log('End Script'); 
}

// Function //

function checkGroup_(groupname,name){
  try {
     var page = AdminDirectory.Groups.get(groupname);
     return true;
  } catch (e) {
     return createMailGroup_(groupname,name)
  } 
}

function addUsertoGroup_(usermail,groupmail){
  var member = {
    email: usermail,
    role: "MEMBER"
  };
  if(!DEBUG){
      try {
          var page = AdminDirectory.Members.insert(member,groupmail);
          Logger.log("ADD "+usermail+" to "+groupmail);
          mailAdd=mailAdd+"Add "+usermail+" to "+groupmail+"<br/>";
      } catch (e) {
         Logger.log(e);
         mailError=mailError+"Add user error: "+member+ " to "+ groupmail+" {"+e+"}<br/>";
      }
  } else {
    console.log("Add "+usermail+" from "+groupmail);
  }     
}

function removeUserfromGroup_(usermail,groupmail){
  if(!DEBUG){
      try {
         var page = AdminDirectory.Members.remove(groupmail,usermail);
         Logger.log("REMOVE "+usermail+" from "+groupmail);
         mailRemove=mailRemove+"Remove "+usermail+" from "+groupmail+"<br/>";
      } catch (e) {
         Logger.log(e);
         mailError=mailError+e+"<br/>"
         mailError=mailError+"Remove user error: "+member+ " to "+ groupmail+" {"+e+"}<br/>";
      }
  } else {
    console.log("Remove "+usermail+" from "+groupmail);
  }    
 
}
function getMembers_(groupKey){
    var pageToken, page;
    do {
        page = AdminDirectory.Members.list(groupKey,
        {
            customer: 'my_customer',
            maxResults: 500,
            pageToken: pageToken,
        });
      var members = page.members
        if (members) {
            for (var i = 0; i < members.length; i++)
            {
                var member = members[i];
                var row = [groupKey, member.email, member.role, member.status];
                groupmembers.push(row);
            }        

        } 
      pageToken = page.nextPageToken;

    } while (pageToken);
    return groupmembers;

}

function getUsers_(path){
  let page = '';
  let pageToken = '';
  var searchQuery='orgUnitPath:'+path;
    do{
    page = AdminDirectory.Users.list({        
        customer: 'my_customer',
        query:searchQuery,
        pageToken: pageToken
      })
      
    if (page.users) {
      // Store only selected info for each user.
      let account = page.users.map(user => {
       return {
        'firstName': user.name.givenName,
        'lastName' : user.name.familyName,
        'email'    : user.primaryEmail
        }
      })
      accounts = accounts.concat(account)
    }
    pageToken = page.nextPageToken;
  
  } while(pageToken)
  
  return accounts;
}

function getCurrentSchoolYear_(){
  var date = new Date();
  var mt = date.getMonth();
  var ye = date.getFullYear();
  if (mt >= 7) { // Aug
    var y = ye + 1;
    var schoolyear = ye+"-"+y;
  } else {
    var y = ye - 1;
    var schoolyear = y+"-"+ye;
  }  
  return schoolyear;
}

function createMailGroup_(email,name)
{
    groupname="L"+name.slice(1).toUpperCase();
    console.log("Create "+groupname);
    try {
       AdminDirectory.Groups.insert({        
         "email": email,
         "name": groupname
       }); 

    } catch (e) {
       Logger.log("Error");
       return false;
    }
    return true;
}

function editGroup_(email,name){
    console.log("Update "+email)
    var groupId = email;
    if(!DEBUG){
        var group = AdminGroupsSettings.newGroups();
        group.name = name;
        group.whoCanAdd = 'NONE_CAN_ADD';
        group.whoCanJoin = 'CAN_REQUEST_TO_JOIN';
        group.whoCanPostMessage = 'ALL_IN_DOMAIN_CAN_POST';
        group.whoCanViewMembership = 'ALL_IN_DOMAIN_CAN_VIEW';
        group.whoCanViewGroup = 'ALL_IN_DOMAIN_CAN_VIEW';
        group.whoCanInvite = 'ALL_MANAGERS_CAN_INVITE';
        group.whoCanContactOwner = 'ALL_IN_DOMAIN_CAN_CONTACT';
        group.whoCanViewGroup = 'ALL_MEMBERS_CAN_VIEW';
        group.allowExternalMembers = false;
        group.allowWebPosting = false;
        group.allowGoogleCommunication = false;
        group.membersCanPostAsTheGroup = false;
        group.includeInGlobalAddressList = false;
        group.whoCanLeaveGroup = 'NONE_CAN_LEAVE';
        AdminGroupsSettings.Groups.patch(group, groupId);
    }
 }

