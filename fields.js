
var fields = {
  'fn': {
    dispName: 'First Name',
    directory: function(userDirObj) {return (userDirObj.name ? userDirObj.name.givenName : "");},
    contacts: function(userContObj) {return (userContObj.names && userContObj.names.length > 0 ? userContObj.names[0].givenName : "");}
  },
  'ln': {
    dispName: 'Last Name',
    directory: function(userDirObj) {return (userDirObj.name ? userDirObj.name.familyName : "");},
    contacts: function(userContObj) {return (userContObj.names && userContObj.names.length > 0 ? userContObj.names[0].familyName : "");}
  },
  'nm': {
    dispName: 'Full Name',
    directory: function(userDirObj) {return (userDirObj.name ? userDirObj.name.fullName : "");},
    contacts: function(userContObj) {return (userContObj.names && userContObj.names.length > 0 ? userContObj.names[0].displayName : "");}
  },
  'ti': {
    dispName: 'Title',
    directory: function(userDirObj) {return (userDirObj.organizations && userDirObj.organizations.length > 0 ? userDirObj.organizations[0].title : "");},
    contacts: function(userContObj) {return (userContObj.organizations && userContObj.organizations.length > 0 ? userContObj.organizations[0].title : "");}
  },
  'de': {
    dispName: 'Department',
    directory: function(userDirObj) {return (userDirObj.organizations && userDirObj.organizations.length > 0 ? userDirObj.organizations[0].department : "");},
    contacts: function(userContObj) {return (userContObj.organizations && userContObj.organizations.length > 0 ? userContObj.organizations[0].department : "");}
  },
}

