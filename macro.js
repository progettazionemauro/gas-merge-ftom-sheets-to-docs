//
// Create a PDF by merging values from a Google spreadsheet into a Google Doc
// ==========================================================================
//
// Demo GSheet & script - http://bit.ly/createPDF
// Demo GDoc template - 1QnWfeGrZ-86zY_Z7gPwbLoEx-m9YreFb7fc9XPWkwDw

//

// Config
// ======

// GDoc Template
// -------------
//
// Replace this with ID of your template document: "https://docs.google.com/document/d/YOUR_GDOC_TEMPLATE_ID_HERE/edit"

// var GDOC_TEMPLATE_ID = ''
var GDOC_TEMPLATE_ID = '12hWNJ10pYYl5HG7CU8dsX7oBd8isjHeYE53TJUXJ3nI' // Template per perizie su II e III Gruppo

// PDF File or GDoc
// ----------------
//
// If set to false the merged file is left as a GDoc

// true or false
var PDF_FILE_CREATE = false

// New Merged File Name
// --------------------
//
// You can specify a name for the new PDF file here, or leave empty to use the default name
// of the form "Merge - [YYYYMMdd_hhmmss].pdf", e.g. "Merge - 20200431_112132.pdf".
// Alternatively one of the columns can be used to name the new file

// This has priority over NEW_FILE_NAME, set to '' to ignore
var HEADER_TO_USE_FOR_FILE_NAME = ''

// Set to '' to ignore, '.pdf' will be added on
var NEW_FILE_NAME = ''

var NEW_FILE_NAME_DEFAULT = 'SGTG_PER_' // + timestamp NEL CASO DI PERIZIE DEL II E III GRUPPO

// Email
// -----
//
// Specify the column header to use for email and whether or not to send emails

// true or false
var EMAIL_SEND = false // true or false

var EMAIL_FIELD_NAME = 'Email'

var EMAIL_SUBJECT = 'The email subject ---- UPDATE ME -----'
var EMAIL_BODY = 'The email body ------ UPDATE ME ---------'

// PDF Folder
// ----------
//
// Specify the ID of the folder that the PDF will be put into

var NEW_FOLDER_ID = ''

// Code
// ====

function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('Mail Merge')
    .addItem('Merge data from active row to create new file', 'mailMerge')
    .addToUi()
} 

/**  
 * Take the fields from the active row in the active sheet
 * and, using a Google Doc template, create a PDF doc with these
 * fields replacing the keys in the template. The keys are identified
 * by being wrapped in curly brackets, e.g. {{Name}}.
 *
 * @return {Object} the completed PDF file
 */

function mailMerge() {

  var copyFile        = null 
  var activeRowIndex  = null 
  var activeRowValues = null
  var headerRow       = null
  var copyBody        = null
  var copyDoc         = null
  var newFile         = null
  var recipient       = null
  var newFileName     = null
  var newFileFolder   = null
  
  var ui              = getUi()
  
  if (!gotGDocTemplate()) {return}
  getSheetData()
  if (isHeaderRow()) {return}  
  replacePlaceholders()  
  locateNewFile()
  setFileName()
  sendEmail()
  displayFinalDialog()
  return
  
  // Private Functions
  // -----------------

  function getUi() {
    var ui = SpreadsheetApp.getUi()
    return {
      UI: ui,
      TITLE: 'Create PDF',
      BUTTONS: ui.ButtonSet.OK
    }
  }

  function getSheetData() {
    
    copyFile = DriveApp.getFileById(GDOC_TEMPLATE_ID).makeCopy()
    var copyId = copyFile.getId()
    copyDoc = DocumentApp.openById(copyId)
    copyBody = copyDoc.getActiveSection()
    var activeSheet = SpreadsheetApp.getActiveSheet()
    var numberOfColumns = activeSheet.getLastColumn()
    activeRowIndex = activeSheet.getActiveRange().getRowIndex()
    var activeRowRange = activeSheet.getRange(activeRowIndex, 1, 1, numberOfColumns)
    activeRowValues = activeRowRange.getDisplayValues()
    headerRow = activeSheet.getRange(1, 1, 1, numberOfColumns).getValues()
  }

  function gotGDocTemplate() {
  
    if (GDOC_TEMPLATE_ID !== '') {return true}    
    ui.UI.alert(ui.TITLE, 'GDOC_TEMPLATE_ID needs to be defined in Code.gs.', ui.BUTTONS)
    return false
  }

  function isHeaderRow() {   
  
    if (activeRowIndex <= 1) {
      ui.UI.alert(ui.TITLE, 'Select a row below the header row.', ui.BUTTONS)
      return true
    } else {
      return false
    }
  }

  function replacePlaceholders() {
  
    for (var columnIndex = 0; columnIndex < headerRow[0].length; columnIndex++) {    
    
      var nextHeader = headerRow[0][columnIndex]
      
      // Replace any non-alphanumeric values in the header and make it case-insenstive
      var nextPlaceholder = '(?i){{' + nextHeader.replace(/[^a-z0-9\s]/gi, ".") + '}}'
      
      var nextValue = activeRowValues[0][columnIndex]        
      
      if (EMAIL_SEND && nextHeader.toLowerCase() === EMAIL_FIELD_NAME.toLowerCase()) {
        recipient = nextValue
      }  
      
      if (HEADER_TO_USE_FOR_FILE_NAME !== '' && 
          nextHeader.toLowerCase() === HEADER_TO_USE_FOR_FILE_NAME.toLowerCase()) {
        newFileName = nextValue
      }
      
      copyBody.replaceText(nextPlaceholder, nextValue)                         
    }  
    
    copyDoc.saveAndClose()
  }

  function locateNewFile() {
  
    var copyFileParentFolder = copyFile.getParents().next() // Assume just one parent  
    
    if (NEW_FOLDER_ID !== '') {
      newFileFolder = DriveApp.getFolderById(NEW_FOLDER_ID)
    } else {
      newFileFolder = copyFileParentFolder
    }
          
    if (PDF_FILE_CREATE) {
    
      var blob = copyFile.getAs('application/pdf')
      newFile = newFileFolder.createFile(blob)         
      copyFile.setTrashed(true)
      
    } else {
    
      newFile = copyFile        
    }
    
    if (NEW_FOLDER_ID !== '') {
      
      // make an orphan
      copyFileParentFolder.removeFile(newFile) 
      
      // then add to new folder, so never in two places
      newFileFolder.addFile(newFile)
      
    } else {
      // The file is already in the right place
    }
    
  } // mailMerge.locateNewFile()

  function setFileName() {  
  
    if (HEADER_TO_USE_FOR_FILE_NAME) {
      if (newFileName === null) {
        throw new Error('Could not find header "' + HEADER_TO_USE_FOR_FILE_NAME + '" for file name')
      }
    } else {
      if (NEW_FILE_NAME !== '') {
        newFileName = NEW_FILE_NAME
      } else {
        var timeZone = Session.getScriptTimeZone()    
        newFileName = NEW_FILE_NAME_DEFAULT + ' - ' + Utilities.formatDate(new Date(), timeZone, 'YYYYMMdd_hhmmss')
      }
    }
    
    newFileName = newFileName + (PDF_FILE_CREATE ? '.pdf' : '')
    newFile.setName(newFileName)
  }

  function sendEmail() {
  
    if (recipient === null) {return}
    
    MailApp.sendEmail(
      recipient, 
      EMAIL_SUBJECT, 
      EMAIL_BODY,
      {attachments: [newFile]})
      
    ui.UI.alert(ui.TITLE, 'New file emailed to ' + recipient + '.', ui.BUTTONS)
  }
  
  function displayFinalDialog() {  
  
    var message = 'the same folder as this GSheet'
    
    if (NEW_FOLDER_ID !== '') {
      message = '"' + newFileFolder.getName() + '" folder'
    }
    
    var userInterface = HtmlService
      .createHtmlOutput(
        '<a href="' + newFile.getUrl() + '" target="_blank">' + 
          'New file "' + newFileName + '" created in ' + message + '.</a>')
      .setWidth(300)
      .setHeight(80)      
      
    ui.UI.showModalDialog(userInterface, ui.TITLE)
  }
} 
