/*----------------------------------------------------------------------------------------------------------------

GLOBAL PROPERTIES/VARIABLES (accessible via File > Project Properties > Script Properties)

docNamePrefix - this will prefix the filename for each document generated
targetFolderID - this is the ID of the Drive folder into which all generated docs will be saved
targetEmailLabel - this is the Gmail label that will be checked for new messages to harvest

----------------------------------------------------------------------------------------------------------------*/


/*----------------------------------------------------------------------------------------------------------------

harvestCitationEmails()

PURPOSE: Primary function, will harvests emails from [targetEmailLabel] & copy attached jpgs to a Google Doc

ARGUMENTS: n/a

RETURN: n/a, all work happens within

----------------------------------------------------------------------------------------------------------------*/
function harvestCitationEmails() {

  var imageArray = [];
  var citationNumber = '';

  var query = 'label:' + PropertiesService.getScriptProperties().getProperty('targetEmailLabel') + ' filename:jpg OR filename:jpeg';
  var threads = GmailApp.search(query);

  for (var i in threads) { // for each email thread in the citations label

    var messages = threads[i].getMessages();

    for (j in messages) { // for each message in the thread (should be one)

      var attachments = messages[j].getAttachments();

      var subject = messages[j].getSubject();
      var subjectparts = subject.split(" "); // email subject is expected to have a citation number at the END, after space
      var citationNumber = subjectparts[subjectparts.length-1].trim();
      Logger.log(subjectparts.length);

      var imagesFound = 0;

      for (k in attachments) { // for each file attached to the message (should be three)
        var attachment = attachments[k];
        if ( checkFileExtension(attachment,'jpg,jpeg') ) { // helper function, defined below
          imageArray[imagesFound] = attachment.copyBlob();
          imagesFound++;
        }
      }

      if (imagesFound) {
        imagesToDoc(imageArray, citationNumber); // function to add the images to a Google Doc, defined below
      }
    }

    var thisLabel = GmailApp.getUserLabelByName(PropertiesService.getScriptProperties().getProperty('targetEmailLabel'));

    threads[i].removeLabel(thisLabel); //remove label so thread won't be processed again
  }
}


/*----------------------------------------------------------------------------------------------------------------

Function to take an array of image objects as an argument and append them to a new Google Doc

ARGUMENTS: (1) an array of image objects; (2) a string for citation number

return: n/a, all work happens within
----------------------------------------------------------------------------------------------------------------*/
function imagesToDoc(imgArray, cNumber) {

  var currDT = Utilities.formatDate(new Date(), "ET", "' - 'yyyy-MM-dd");
  var newDocName = PropertiesService.getScriptProperties().getProperty('docNamePrefix') + cNumber + currDT;
  var newDocFolderID = PropertiesService.getScriptProperties().getProperty('targetFolderID');

  // first, create the target Google Doc
  var newPTDoc = createPTdocument(newDocName, newDocFolderID); // helper function, defined below

  for (var i in imgArray) {

    var thisBlob = imgArray[i];

    var inlineImg = newPTDoc.getBody().insertImage((i),thisBlob);

    var width = inlineImg.getWidth();
    var newW = width;
    var height = inlineImg.getHeight();
    var newH = height;
    var ratio = width/height;

    if(width>640){
      newW = 640;
      newH = parseInt(newW/ratio);
    }

    inlineImg.setWidth(newW).setHeight(newH) // downsize the image to fit standard Google Doc width

    if (i < (imgArray.length - 1) ) {
      // I decided page breaks aren't necessary, as the page will typically break naturally between images
      // newPTDoc.getBody().insertPageBreak(((2*i)+1));
    }

  }
}


/*----------------------------------------------------------------------------------------------------------------

Function to create a new Google Doc with a name provided as an argument

ARGUMENTS: (1) a name for the new document; (2) the parent folder ID

return: a Google Doc object
----------------------------------------------------------------------------------------------------------------*/
function createPTdocument(docName, folderID) {

  var newDoc = DocumentApp.create(docName); // create the file
  docFile = DriveApp.getFileById( newDoc.getId() );
  DriveApp.getFolderById(folderID).addFile( docFile ); // add the file to the parent folder
  DriveApp.getRootFolder().removeFile(docFile); // remove the file from the root

  return newDoc;
}


/*----------------------------------------------------------------------------------------------------------------

Function to check attachment for particular file extension

ARGUMENTS: (1) file attachment object; (2) a string containing all acceptable file extensions

return: Boolean
----------------------------------------------------------------------------------------------------------------*/
function checkFileExtension(attachment, acceptableExtensions){

  var fileName = attachment.getName();
  var temp = fileName.split('.');
  var fileExtension = temp[temp.length-1].toLowerCase();

  if ( acceptableExtensions.search(fileExtension) < 0 ) {
    return false;
  }
  else {
    return true;
  }
}
