// This is the support Javascript for a simple Google Sheets CRM implementation
// This was created by Dan Ledger (@dledge)

// The duration of an entire day in milliseconds
var oneDay = 24*60*60*1000;

// Today
var today = new Date();

// Parameters

// This script will write a block of 4 columns (date last contacted, days since contact, flagged?, did I reply?)
// COLUMN_WHERE_MAGIC_BEGINS is the column where these four will be inserted.  Note that the first column has a value
// of 1.

var COLUMN_WHERE_MAGIC_BEGINS = 7;

// And this is the column where your email addresses are.  Again, a value of 1 means the first column.
var COLUMN_WITH_EMAIL_ADDRESSES = 2;

function catParty() {
  
  // Use the cache to determine how far we got last time before the inevitable 5 minute Google timeout
  var cache = CacheService.getScriptCache(); 
  var lastRowProcessed = cache.get("lastRow")*1.0;
  
  // If this cache doesn't yet exist, create it and set last row to 1
  if (lastRowProcessed == null) {
    lastRowProcessed = 1;
    cache.put("lastRow", lastRowProcessed, 60*60*24); // cache for 25 minutes
  }  
  
  // Connect to our active sheet and collect all of our email addresses in the second column
  var sheet = SpreadsheetApp.getActiveSheet();
  var totalRows = sheet.getLastRow();
  var range = sheet.getRange(2, COLUMN_WITH_EMAIL_ADDRESSES, totalRows, 1);
  var emails = range.getValues();  
  
  // Attempt to iterate through 100 times (although we'll timeout before this)
  for (var cntr = 0; cntr<100; cntr++ ) {
    
    // If we've reached the end of our last, wrap to the front
    if (lastRowProcessed >= totalRows) lastRowProcessed = 1;
    
    // Increment the row we're processing
    var currentRow = lastRowProcessed+1;
    
    // Get the email address from the current row
    var email = emails[currentRow-2][0];
    
    // If the email address field is empty, skip to the next row
    if (!email) {
      lastRowProcessed = currentRow;
      cache.put("lastRow", currentRow, 60*60*24); 
      continue;
    } 
    
    // Look for all threads from me to this person
    var threads = GmailApp.search('from:me to:'+email);
    
    // If there are no threads, I haven't emailed them before
    if (threads.length == 0) {
      
      // Update the spreadsheet row to show we've never emailed
      var range = sheet.getRange(currentRow, COLUMN_WHERE_MAGIC_BEGINS,1, 4 ).setValues([["NEVER", "", "", ""]] );
      
      //  And cary one
      lastRowProcessed = currentRow;
      cache.put("lastRow", currentRow, 60*60*24); // cache for 25 minutes    
      continue;
    }
    
    // Beyond a reasonable doubt
    var latestDate = new Date(1970, 1, 1);
    
    var starredMsg = "";
    var iReplied = ""
    
    // Iterate through each of the message threads returned from our search
    for (var thread in threads) {
      
      // Grab the last message date for this thread
      var threadDate = threads[thread].getLastMessageDate();
                        
      // If this is the latest thread we've seen so far, make note!
      if (threadDate > latestDate) {
        
        latestDate = threadDate;
        
        // Check to see if we starred the message (we may be back to overwrite this)
        if (threads[thread].hasStarredMessages()) {
          starredMsg = "Y";
        } else {
          starredMsg = "";
        }           
        
        // Open the thread to see who was the last to speak
        var messages = threads[thread].getMessages();
        var lastMsg = messages[messages.length-1];
        var lastMsgFrom = lastMsg.getFrom();
        
        // Use regex so we can make our search case insensitive
        var re = new RegExp(email,"i");
        
        // If we can find their email address in the email address from the last message, they spoke last
        // (we may be back to overwrite this)
        if (lastMsgFrom.search(re) >= 0) {
          iReplied = "N";
        } else {
          iReplied = "";
        }       
      }
    }
    
    // Determine how many days have passed since our last correspondence 
    var daysSinceContact = Math.round(Math.abs((today.getTime() - latestDate.getTime())/(oneDay)));
    
    // Format the date so it plays nicely with Google Sheets
    sheet.get
    var latestDate = Utilities.formatDate(latestDate, SpreadsheetApp.getActive().getSpreadsheetTimeZone(),  "MMM d yyyy");
    
    // Write the row!
    var range = sheet.getRange(currentRow, COLUMN_WHERE_MAGIC_BEGINS, 1, 4 ).setValues([[latestDate, daysSinceContact, starredMsg, iReplied]] );
    
    // update cache
    cache.put("lastRow", currentRow, 60*60*24); 
    
    // update lastRowProcessed
    lastRowProcessed = currentRow;
    
    // Log it (mostly to see how many of these we're making it through per run
    Logger.log("processed "+currentRow);   
  }
}


/*
function getEmailAddress(name,domain) {
  
  var url = "54.152.93.219/ssearch/search.php?q="+name+"%20"+"@"+domain;
  var result = JSON.parse(UrlFetchApp.fetch(url).getContentText());
  
  var results = result['bossresponse']['web']['results'];
  
  var found = false;
  for (var i in results) {
   var abstract=results[i]['abstract'];
  
    if (abstract.search('@'+domain)>= 0) {
      found = true;
    } else {
      found = false;
    }
  }
}

function testDomain() {
 var res = getEmailAddress("Tom essery","timex.com");
  res = res;
  
}


function GetLastEmail(email) {
  var threads = GmailApp.search('from:me to:'+email);
  
  if (threads.length == 0) return ["Never","",""]
  
  var latestDate = new Date(2010, 1, 1);
  
  var starredMsg = "";
  
  for (var thread in threads) {
    var threadDate = threads[thread].getLastMessageDate();
        
    if (threadDate > latestDate) {
      latestDate = threadDate;

      if (threads[thread].hasStarredMessages()) {
        starredMsg = "Y";
      } else {
        starredMsg = "";
      }      
    }    
  }
  
  
  var daysSinceContact = Math.round(Math.abs((today.getTime() - latestDate.getTime())/(oneDay)));
  var latestDate = Utilities.formatDate(latestDate, Session.getScriptTimeZone(),  "MMM d yyyy");
 
//  return Utilities.formatDate(latestDate, Session.getScriptTimeZone(),  "MMM d yyyy HH:mm:ss");
  return [latestDate, daysSinceContact, starredMsg];
}



function run() {
  
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var ROWS = sheet.getLastRow();
  var range = sheet.getRange(2, 2, ROWS, 1);
  var values = range.getValues();  
  
  
  var latestDates = []
  
  for (var v in values) {
    var email = values[v][0];
    if (!email) {
      latestDates.push(["","",""]);
    } else {
      var lastContact = GetLastEmail(email);    
      latestDates.push(lastContact)
    }
  }
  
  var range = sheet.getRange(2, 7, ROWS, 3);
  var values = range.setValues(latestDates);  
  
  
 // res = GetLastEmail('cbonavito@nyas.org');
}
*/
