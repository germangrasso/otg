/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      $('#run').click(run);
    });
  };

  function run() {
    
    
    /**
     * Insert your Outlook code here
     */
    console.log(Office.context.mailbox.item);
    loadProps(Office.context.mailbox.item);
    
  }

  // Load properties from the Item base object, then load the
  // type-specific properties.
  function loadProps(item) {
    
    $('#itemType').text(item.itemType);
    
    item.body.getAsync('html', function(result){
      if (result.status === 'succeeded') {
        $('#bodyHtml').text(result.value);
      }
    });
    
    item.body.getAsync('text', function(result){
      if (result.status === 'succeeded') {
        $('#bodyText').text(result.value);
      }
    });
    
    item.subject.getAsync({}, function(result){
        if (result.status === 'succeeded') {
          $('#subject').text(result.value);
        }
      });

    item.to.getAsync({}, function(result){
        if (result.status === 'succeeded') {
            $('#to').html(buildEmailAddressesString(result.value));
        }
    });
    
    item.cc.getAsync({}, function(result){
        if (result.status === 'succeeded') {
            $('#cc').html(buildEmailAddressesString(result.value));
        }
    });
    
    item.bcc.getAsync({}, function(result){
        if (result.status === 'succeeded') {
            $('#bcc').html(buildEmailAddressesString(result.value));
        }
    });
    $('#conversationId').text(item.conversationId);

  }

   
  function loadNewItem(eventArgs) {
    loadProps(Office.context.mailbox.item);
  };
  
  // Take an array of AttachmentDetails objects and
  // build a list of attachment names, separated by a line-break
  function buildAttachmentsString(attachments) {
    if (attachments && attachments.length > 0) {
      var returnString = "";
      
      for (var i = 0; i < attachments.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + attachments[i].name;
      }
      
      return returnString;
    }
    
    return "None";
  }
  
  // Format an EmailAddressDetails object as
  // GivenName Surname <emailaddress>
  function buildEmailAddressString(address) {
    return address.displayName + " &lt;" + address.emailAddress + "&gt;";
  }
  
  // Take an array of EmailAddressDetails objects and
  // build a list of formatted strings, separated by a line-break
  function buildEmailAddressesString(addresses) {
    if (addresses && addresses.length > 0) {
      var returnString = "";
      
      for (var i = 0; i < addresses.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + buildEmailAddressString(addresses[i]);
      }
      
      return returnString;
    }
    
    return "None";
  }

})();