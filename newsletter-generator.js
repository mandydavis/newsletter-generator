// whatever you wish the name of the new Google Doc to be 
var fileName = 'nuxcNewsletter';

// set up
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

// need for retrieving images later
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

// for listing responses one-by-one underneath the particular question. results in a bulleted list, lists the submittee's name in brackets after their response
function listResponses(title, column, startingRow, endingRow) {
  var heading = body.appendParagraph(title);
  headingStyle(heading);
  for (var i = startingRow - 1; i < endingRow; i++) {
    if (data[i][column]) {
      // append an individual response with a square-shaped bullet point
      var bullet = body.appendListItem(data[i][column] + ' ['+ data[i][20] + ']').setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);;
      bodyStyle(bullet);
    }
  }
  body.appendPageBreak();
}

// to generate the 'squad spotlights' sections (name, images, image caption, spotlight questions)
function spotlightText(boldText, nameIndex, column) {
      var bullet = body.appendParagraph(boldText);
      bodyStyleBold(bullet);
      var response = bullet.appendText(data[nameIndex][column]);
      response.setBold(false);
}


///////////////  main function: automatic newsletter generator ///////////////////////

function automaticNewsletter() {
  
  var ui = SpreadsheetApp.getUi();
  
  // first input
  var response = ui.prompt('Starting row # for this week\'s responses:');

  // process the user's response to first input
  var button = response.getSelectedButton();
  var text = response.getResponseText();
  
  if (button == ui.Button.CLOSE) {
    // if the user hits the cancel button, terminate the script
  }
  
 
  // all's good (user did not click 'X' on first input) so move on to second input
  else if (button == ui.Button.OK) {
      var response2 = ui.prompt('Email address to send the automatic newsletter to: ');
      // process the user's response to the second input
      var button2 = response2.getSelectedButton();
      var text2 = response2.getResponseText();
    
      if (button2 == ui.Button.CLOSE) {
         // if the user hits the cancel button, terminate the script
      }
    
      // the user has not inputted everything we need, so it's time to generate the newsletter! 
      else if (button2 == ui.Button.OK) {
        var startingRow = Number(response.getResponseText());
        var endingRow = sheet.getLastRow();
        var email = response2.getResponseText();
        ui.alert('little elves in the clouds are putting your newsletter together now. They will send it to your email ASAP!');
      
        // days until Big Tens! (the calculates the number of days between today and Nov. 1, 2020)
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
        
        // generate NUXC Squad Spotlights
        for (var i = startingRow-1; i < endingRow; i++) {
          
          // write name
          var name = body.appendParagraph(data[i][20]);
          
          // style
          name.setFontSize(14);
          name.setFontFamily('Yellowtail');
          name.setBackgroundColor('#ffffff');
          name.setBold(true); 
          
          // upload image(s) if provided
          if (data[i][8]) {
            var pictureHolder = body.appendParagraph('');
            var photoLink = data[i][8];
            var multipleLinks = photoLink.split(', ');
            
            // handle multiple photo uploads
            for (var link in multipleLinks) {
              var currentLink = multipleLinks[link];
              var imageID = getIdFromUrl(currentLink);      
              var img = DriveApp.getFileById(imageID).getBlob();
              
              // adds photo(s), includes error handling for invalid photo file types   
              try {
                var image = pictureHolder.appendInlineImage(img);
                resizeImg(image, 300);
                pictureHolder.setBackgroundColor(null);
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
        body.appendPageBreak();
        
        // list out people's responses to this week's questions
        // an 'X' in the cell below a question in the sheet means it will be included in the newsletter
        for (col = 5; col < sheet.getLastColumn(); col++) {
          if (data[2][col] == 'X') {
            listResponses(data[1][col], col, startingRow, endingRow);
          } 
        }
        
        // attempt with actual checkboxes, doesn't currently work:
        //  for (col = 5; col < sheet.getLastColumn(); col++) {
        //    var range = sheet.getRange(2,col);
        //    if (range.isChecked()) {
        //      listResponses(data[1][col], col, startingRow, endingRow);
        //    }
        //  }
        
        // send email of missing images/videos
        GmailApp.sendEmail(email, subject, imagesNotIncluded.getUrl());
        
        // send email of automated newsletter :) 
        GmailApp.sendEmail(email, 'automatic newsletter!', 'here she is: ' + doc.getUrl());
        
      }
  }
      
 
}