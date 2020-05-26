//
//  Title: CreateFilters.gs
//  Author: Gary Holeman
//  Date: May 25, 2020
//
//  This is a Google App Script that will test the first 100 messages in the Inbox that do not have the label Contact on them.
//  It tests to see if the message is from someone that is in you Contacts.  If they are, it marks the thread with the Contact label so
//  it won't test them again the next time.  I have this set to run every 5 minutes.  If the thread is not from a Contact, it labels the thread
//  with the Filtered label and then creates a new filter for this address that will remove from the Inbox and label with Filtered automatically
//  in the future.  It then moves this thread to the archive.
//
//
function testMessages() { 
  
  var threads, message, thread = {};
  var from = "";
  var t=0;
  var matchingContacts = [];
  var label = {};
  
  // Wrap the entire function in a try / catch, in case there is an error, log it.
  try {
    
    // Get the most recent 100 threads in your inbox that are not already marked as from a contact
    threads = GmailApp.search("in:inbox -label:Contact", 0, 100);
    
    // If there are threads
    if (threads.length > 0) {
      
      // For each thread
      for (t=threads.length-1; t>=0; t--) {
        
        
        // Get the current thread we are iterating over
        thread = threads[t];
            
        // Get the first message in the thread
        message = thread.getMessages()[0];
            
        // Get the sender's email address.  The regular expression will extract the email address out of "First Last <email@address.com>".
        from = message.getFrom().replace(/^.+<([^>]+)>$/, "$1");


// Find all of my Contacts that have this email address
        matchingContacts = ContactsApp.getContactsByEmailAddress(from);
        if (matchingContacts.length > 0) {
//  If this address is in my Contacts, mark this thread with the label Contact so that I won't test this message again
          label = GmailApp.createLabel("Contact");
          if (label !== null) {
            thread.addLabel(label);
          }
        } else {
// If I have no Contacts with this address, then create a filter for this address and label this thread Filtered, moving the thread to the Archive			
		createToFilter(from,'Filtered');
        label = GmailApp.getUserLabelByName("Filtered");
        thread.addLabel(label);
        Logger.log('Filtered ' + from);
        thread.moveToArchive();
        }
        
      }
    }
  } catch (e) {
    Logger.log(e.toString());
  }
}


// Creates a filter to put all email from ${toAddress} into
// Gmail label ${labelName}
function createToFilter(toAddress, labelName) {

try 
{
// Lists all the filters for the user running the script, 'me'
var labels = Gmail.Users.Settings.Filters.list('me')

// Search through the existing filters for ${toAddress}
var label = true
labels.filter.forEach(function(l) {
    if (l.criteria.to === toAddress) {
        label = null
    }
})
}
  catch (e) {
    Logger.Log(e.ToString);
  }
  finally {
// If the filter does exist, return
if (label === null) return
else { 

// Create the new label 
GmailApp.createLabel(labelName)

// Lists all the labels for the user running the script, 'me'
var labelids = Gmail.Users.Labels.list('me')

// Search through the existing labels for ${labelName}
// this operation is still needed to get the label ID 
var labelid = false
labelids.labels.forEach(function(a) {
    if (a.name === labelName) {
        labelid = a
    }
})

// Create a new filter object (really just POD)
var filter = Gmail.newFilter()

// Make the filter activate when the to address is ${toAddress}
filter.criteria = Gmail.newFilterCriteria()
filter.criteria.to = toAddress

// Make the filter remove the label id of ${"INBOX"}
filter.action = Gmail.newFilterAction()
filter.action.removeLabelIds = ["INBOX"];
// Make the filter apply the label id of ${labelName} 
filter.action.addLabelIds = [labelid.id];

// Add the filter to the user's ('me') settings
Gmail.Users.Settings.Filters.create(filter, 'me')
}
}
}
