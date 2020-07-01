var fileName = 'nuxcNewsletter';

// DO NOT CHANGE ANYTHING BELOW THIS, PLEASE :)

var sheet = SpreadsheetApp.getActiveSheet();
var data = sheet.getDataRange().getValues();
var doc = DocumentApp.create(fileName);
var body = doc.getBody();


////////////////// helper functions /////////////////////////////

// adds a custom option to the menu called 'Generate Newsletter' so the script can be run w/o accessing script editor
function onOpen() {  
  var ui = SpreadsheetApp.getUi();
  // var menuEntries = [ {name: "well what are ya waiting for?", functionName: "automaticNewsletter"} ];
  ui.createMenu("Generate Newsletter")
  .addItem("well what are ya waiting for?", "automaticNewsletter")
  .addToUi();
  
}

function getIdFromUrl(url) { return url.match(/[-\w]{25,}/); }


////////////////// style /////////////////////////////

// consistent sizing of images
function resizeImg(img, targetHeight) {
    var height = img.getHeight();
    var width = img.getWidth();
    var factor = height / targetHeight;
    img.setHeight(height / factor);
    img.setWidth(width / factor);
};

function bodyStyle(content) {
  content.setFontSize(11);
  content.setFontFamily('Source Sans Pro');
  content.setBold(false);  
  content.setBackgroundColor(null);
}

function bodyStyleBold(content) {
  bodyStyle(content);
  content.setBold(true);
}

function headingStyle(content) {
  content.setFontSize(18);
  content.setFontFamily('Source Sans Pro');
  content.setBold(true);
  content.setBackgroundColor(null);
}


////////////////// main formats /////////////////////////////

function listResponses(title, column, startingRow, endingRow) {
  var heading = body.appendParagraph(title);
  headingStyle(heading);
  for (var i = startingRow - 1; i < endingRow; i++) {
    if (data[i][column]) {
      var bullet = body.appendParagraph(data[i][column] + ' ['+ data[i][1] + ']');
      bodyStyle(bullet);
    }
  }
  body.appendPageBreak();
}

function spotlightText(boldText, nameIndex, column) {
      var bullet = body.appendParagraph(boldText);
      bodyStyleBold(bullet);
      var response = bullet.appendText(data[nameIndex][column]);
      response.setBold(false);
}


///////////////  main function: automatic newsletter generator ///////////////////////

function automaticNewsletter() {
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Starting row # for this week\'s responses:').getResponseText();
  var response2 = ui.prompt('Ending row # for this week\'s responses:').getResponseText();
  var email = ui.prompt('Email address to send the automatic newsletter to: ').getResponseText();
  var startingRow = Number(response);
  var endingRow = Number(response2);
  ui.alert('little elves in the clouds are putting your newsletter together now. They will send it to your email ASAP!');
  
  // Days until Big Tens!
  var todaysDate = new Date();
  var bigTensDate = new Date(2020, 10, 1);
  var millisDiff = bigTensDate.getTime() - todaysDate.getTime();
  var daysDiff = millisDiff/1000/60/60/24;
  var roundedDays = Math.ceil(daysDiff);
  var bullet = body.appendParagraph('Days until Big Tens: ' + roundedDays);
  headingStyle(bullet);
  body.appendPageBreak();
  
  
  // set up for email containing missing files since Google Docs doesn't support .heic or .mov
  var subject = 'files that could not be added to newsletter automatically: ';
         var imagesNotIncluded = DocumentApp.create('imagesNotIncluded');
         var body2 = imagesNotIncluded.getBody();
  
  // NUXC Squad Spotlights
  for (var i = startingRow-1; i < endingRow; i++) {
    
    // write name
    var name = body.appendParagraph(data[i][1]);
    name.setFontSize(14);
    name.setFontFamily('Yellowtail');
    name.setBackgroundColor('#ffffff');
    name.setBold(true); 
    
    // upload image(s) if provided
     if (data[i][8]) {
       var pictureHolder = body.appendParagraph('');
       var photoLink = data[i][8];
       var multipleLinks = photoLink.split(', ');
       
       for (var link in multipleLinks) {
         var currentLink = multipleLinks[link];
         var imageID = getIdFromUrl(currentLink);      
         var img = DriveApp.getFileById(imageID).getBlob();
      
       // adds photo(s), includes error handling for invalid photo file types   
         try {
           var image = pictureHolder.appendInlineImage(img);
           resizeImg(image, 300);
         }
         
         catch(err) {
           // add to list of invalid photo file types to be emailed
           body2.appendParagraph(data[i][1] + ': ' + currentLink);
         }
       
       }         
       
    }
        
    // add photo caption if available
    if (data[i][14]) {
      var bullet = body.appendParagraph('📸: ' + data[i][14]);
      bodyStyle(bullet);
      body.appendParagraph(' ');
    }
    
    // write Highlight of Training Week if not empty (if true then it's not blank)
    if (data[i][3]) {
      spotlightText('Highlight of the training week: ', i, 3);
    }
    
    // Write Biggest Struggle or Lesson if not empty
    if (data[i][4]) {
      spotlightText('Biggest struggle or lesson: ', i, 4);
    }
    
    // Write Goal for Next Week if not empty
    if (data[i][10]) {
      spotlightText('Goal for next week: ', i, 10);
    } 
    body.appendParagraph('');
   
  }
  
  // List out people's responses (compatible with wording changes week-to-week)
  
  // List out people's responses!  (previous)
  listResponses('What do you spend the majority of your waking hours doing?', 6, startingRow, endingRow);
  
  listResponses('Music for this week!', 7, startingRow, endingRow);
  
  listResponses('Best media you have consumed this week?', 9, startingRow, endingRow);
  
  listResponses('Fun family/quarantine-buddy stories?', 12, startingRow, endingRow);
  
  listResponses('What is one thing you are grateful for?', 13, startingRow, endingRow);
  
  listResponses('Who are the 3-5 people in the Big Ten you are working to beat this season racing', 15, startingRow, endingRow);
  
  // send email of missing images/videos
  GmailApp.sendEmail(email, subject, imagesNotIncluded.getUrl());

  // send email of automated newsletter :) 
  GmailApp.sendEmail(email, 'automatic newsletter!', 'here she is: ' + doc.getUrl());
 
}