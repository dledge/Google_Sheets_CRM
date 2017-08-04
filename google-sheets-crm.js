// This is the support Javascript for a simple Google Sheets CRM implementation
// This was created by Dan Ledger (@dledge)

// The duration of an entire day in milliseconds
var oneDay = 24*60*60*1000;

// Today
var today = new Date();

// My email address
var myEmailAddress = Session.getActiveUser().getEmail();

// Parameters

// This script will write a block of 4 columns (date last contacted, days since contact, flagged?, did I reply?)
// COLUMN_WHERE_MAGIC_BEGINS is the column where these four will be inserted.  Note that the first column has a value
// of 1.
var COLUMN_WHERE_MAGIC_BEGINS = 10;

// And this is the column where your email addresses are.  Again, a value of 1 means the first column.
var COLUMN_WITH_EMAIL_ADDRESSES = 2;

function runSimpleCRM() {
  
  // Use the cache to determine how far we got last time before the inevitable 5 minute Google timeout
  var cache = CacheService.getScriptCache(); 
  var lastRowProcessed = cache.get("lastRow")*1.0;
  
  // If this cache doesn't yet exist, create it and set last row to 1
  if (lastRowProcessed == null || lastRowProcessed == 0) {
    lastRowProcessed = 1;
    cache.put("lastRow", lastRowProcessed, 60*60*24); // cache for 25 minutes
  }  
  
  // Keep track of where we started 
  var startingRow = lastRowProcessed;
  
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
    
    // If we've been through all records, bail
    if (currentRow == startingRow) {
      break; 
    }
    
    // Get the email address from the current row
    var theirEmail = emails[currentRow-2][0].trim();
    
    // If the email address field is empty, skip to the next row
    if (!theirEmail) {
      lastRowProcessed = currentRow;
      cache.put("lastRow", currentRow, 60*60*24); 
      continue;
    } 
    
    // Look for all threads from me to this person
    var threads = GmailApp.search('from:me to:'+theirEmail);
    
    // Add a quick pause so we don't run into rate limiting issues with gmail
    Utilities.sleep(20);
    
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
    var threadStatus = ""

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
        var totalMessages = messages.length;
        var lastMsg = messages[messages.length-1];
        var lastMsgFrom = lastMsg.getFrom();
        
        
        // Use regex so we can make our search case insensitive
        var reTheirEmail = new RegExp(theirEmail,"i");
        var reMyEmail = new RegExp(myEmailAddress,"i");

        // If we can find their email address in the email address from the last message, they spoke last
        // (we may be back to overwrite this)
        /*
        var rawEmail = lastMsg.getRawContent();
        var reCal1 = new RegExp('text\/calendar',"i");
        var reCal2 = new RegExp('calendar-notification',"i");

        if (theirEmail == "dledger@gmail.com") {
          var hereWeStop = true; 
          Logger.clear();
          Logger.log(rawEmail);
        }

        
        if (rawEmail.search('text/calendar') >= 0 || rawEmail.search('calendar-notification') >= 0) {
          // Ignore it unless this is the only correspondence so far
          threadStatus = "We were on a calendar invite together"
        }         
        else */
        if (lastMsgFrom.search(reTheirEmail) >= 0) {
          if (totalMessages == 1) {
            threadStatus = "They emailed me, I haven't replied";
          } else {
            threadStatus = "We had an email exchange, they replied last";
          }
        } else if (lastMsgFrom.search(reMyEmail) >= 0) {
          if (totalMessages == 1) {
            threadStatus = "I emailed them, they haven't replied";
          } else {
            threadStatus = "We had an email exchange, I replied last";            
          }          
        } else {
          threadStatus = "We were on a thread together, neither of us replied last";            
        }
      }
    }
  
    // Determine how many days have passed since our last correspondence 
    var daysSinceContact = Math.round(Math.abs((today.getTime() - latestDate.getTime())/(oneDay)));
    
    // Format the date so it plays nicely with Google Sheets
    sheet.get
    var latestDate = Utilities.formatDate(latestDate, SpreadsheetApp.getActive().getSpreadsheetTimeZone(),  "MMM d yyyy");
    
    // Write the row!
    var range = sheet.getRange(currentRow, COLUMN_WHERE_MAGIC_BEGINS, 1, 4 ).setValues([[latestDate, daysSinceContact, starredMsg, threadStatus]] );
    
    // update cache
    cache.put("lastRow", currentRow, 60*60*24); 
    
    // update lastRowProcessed
    lastRowProcessed = currentRow;
    
    // Log it (mostly to see how many of these we're making it through per run
    Logger.log("processed "+currentRow);   
  }
}

