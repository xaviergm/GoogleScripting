/**
 * Google Apps Script - List all files and folders within the selected folder (see hint) that have been shared with others & write into a speadsheet.
 *    - Main function: List all files & folders.
 * 
 * Hint: Set your folder ID first! You may copy the folder ID from the browser's address field. 
 *       The folder ID is everything after the 'folders/' portion of the URL.
 *       The root folder in your drive can be replaced by the alias 'root'.
 * 
 *       Based on the version 1.0 of mesgarpour's script for listing files and folders. See https://github.com/mesgarpour
 *
 *
 * @Version 1.0
 * @Creator: xavier(dot)gutierrez(at)roche(dot)com
 */
 
// TODO: Set folder ID - The folder that is entered here will be the starting point. Use 'root' for the whole "My Drive".
var initFolderID = '0B8YQte3jdzH1RzVHOF9yV1JtOVE';

// Other Variables
var initOwnerEmail = Session.getActiveUser().getEmail();  // Obtaining the Owner Email Address in order to not include files/folders not owned by them.

// Main function: List all files & folders & write into the current sheet.
function listAllShared(){
  getFolderTree(initFolderID, initOwnerEmail, true, 'test'); // Sometimes you need to pass some useless variable due to Google Scripting Cache issues (well known issue).
};

// =================
// Get Folder Tree
function getFolderTree(folderId, ownerEmail, listAll) {
  try {
    
    // Get folder by id
    var parentFolder = DriveApp.getFolderById(folderId);
    
    // Initialise the sheet
    var file, data, sheet = SpreadsheetApp.getActiveSheet();
    sheet.clear();
    sheet.setName("Folder: " + parentFolder); // Sets the name of the sheet to that of the parent folder where the search is being conducted.
    sheet.appendRow(["Identification of files"]);
    sheet.appendRow(["In folder:       " + parentFolder, "", "Number of Folders:"]);
    sheet.getRange(2,4).setFormula('=COUNTIF(B:B, "Folder")'); // This will count and display the number of folders owned by the user and shared with others.
    sheet.appendRow(["Owned by:    " + ownerEmail, "", "Number of Files:"]);
    sheet.getRange(3,4).setFormula('=COUNTIF(B:B, "File")'); // This will count and display the number of files shared by the user and shared with others.
    sheet.appendRow(["That are shared with others."]);
    sheet.appendRow([" "]);
    sheet.appendRow(["Name", "Type", "Owner", "Sharing", "Editors", "Viewers / Commenters", "Full Path", "URL", "Date", "Last Updated", "Size", "Description"]); // Column Labels
    sheet.setFrozenRows(6); // Freezes just below the labels for scrolling later on.
    
   
    // Get files and folders by calling the necessary function.
    getChildFolders(parentFolder.getName(), parentFolder, data, sheet, ownerEmail, listAll); 
    
  } catch (e) {
    Logger.log(e.toString()); // Logs errors to "Logs". Use CTRL + ENTER to view Logs.
  }
};

// Get the list of files and folders and their metadata in recursive mode.
function getChildFolders(parentName, parent, data, sheet, ownerEmail, listAll) {
  var childFolders = parent.getFolders();
  
  // List folders inside the folder
  while (childFolders.hasNext()) {
    var childFolder = childFolders.next();
    var childFolderOwnerEmail = childFolder.getOwner().getEmail();
    var childFolderEditors = childFolder.getEditors();
    var childFolderEditorsCons = ""
    
    // Consolidate all folder editors in a single variable for writing to a single cell
    for (var i = 0; i < childFolderEditors.length; i++) {
      childFolderEditorsCons = childFolderEditorsCons + childFolderEditors[i].getEmail() + "\n";
    }
    childFolderEditorsCons = childFolderEditorsCons.replace(/\r?\n?[^\r\n]*$/, "");
    var childFolderViewers = childFolder.getViewers();
    var childFolderViewersCons = ""
    
    // Consolidate all folder viewers and commenters in a single variable for writing to a single cell
    for (var i = 0; i < childFolderViewers.length; i++) {
      childFolderViewersCons = childFolderViewersCons + childFolderViewers[i].getEmail() + "\n";
    }
    childFolderViewersCons = childFolderViewersCons.replace(/\r?\n?[^\r\n]*$/, "");
    data = [ 
      childFolder.getName(),
      "Folder",
      childFolder.getOwner().getEmail(),
      childFolder.getSharingAccess(),
      childFolderEditorsCons,
      childFolderViewersCons,
      parentName + "/" + childFolder.getName(),
      childFolder.getUrl(),
      childFolder.getDateCreated(),
      childFolder.getLastUpdated(),
      childFolder.getSize(),
      childFolder.getDescription()
    ];
    
    // Write to the sheet
    if (ownerEmail == childFolderOwnerEmail && (childFolderEditorsCons != "" || childFolderViewersCons != "") ) {sheet.appendRow(data)}; // Only write if the person executing the script owns the folder
    
    // List files inside the folder
    var files = childFolder.getFiles();
    while (listAll & files.hasNext()) {
      var childFile = files.next();
      var childFileOwnerEmail = childFile.getOwner().getEmail();
      var childFileEditors = childFile.getEditors();
      var childFileEditorsCons = ""
      
      // Consolidate all file editors in a single variable for writing to a single cell
      for (var i = 0; i < childFileEditors.length; i++) {
        childFileEditorsCons = childFileEditorsCons + childFileEditors[i].getEmail() + "\n";
      }
      childFileEditorsCons = childFileEditorsCons.replace(/\r?\n?[^\r\n]*$/, "");
      var childFileViewers = childFile.getViewers();
      var childFileViewersCons = ""
      
      // Consolidate all file viewers and commenters in a single variable for writing to a single cell
      for (var i = 0; i < childFileViewers.length; i++) {
        childFileViewersCons = childFileViewersCons + childFileViewers[i].getEmail() + "\n";
      }
      childFileViewersCons = childFileViewersCons.replace(/\r?\n?[^\r\n]*$/, "");
      data = [ 
          childFile.getName(),
          "File",
          childFile.getOwner().getEmail(),
          childFile.getSharingAccess(),
          childFileEditorsCons,
          childFileViewersCons,
          parentName + "/" + childFolder.getName() + "/" + childFile.getName(),
          childFile.getUrl(),
          childFile.getDateCreated(),
          childFile.getLastUpdated(),
          childFile.getSize(),
          childFile.getDescription()
      ];
      
      // Write to the sheet
      if (ownerEmail == childFileOwnerEmail && (childFileEditorsCons != "" || childFileViewersCons != "") ) {sheet.appendRow(data)} // Only write if the person executing the script owns the file
    }
    
    // Recursive call of the subfolder
    getChildFolders(parentName + "/" + childFolder.getName(), childFolder, data, sheet, ownerEmail, listAll);  
  }
};
