
function OnSubmit(e) {
  var destinationFolders = DriveApp.getFoldersByName('weLinh');
  if (destinationFolders.hasNext()) {
    var folder = destinationFolders.next();
    
    var destinationFolderId = folder.getId();
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3);
    var rows = range.getValues();
    
    var lastAppendedRow = rows[rows.length - 1];
    var fileInfo = {
      date: lastAppendedRow[0],
      mssv: lastAppendedRow[1],
      fileId: lastAppendedRow[2].substring(lastAppendedRow[2].indexOf("=") + 1, lastAppendedRow[2].length)
    }
    
    moveFileTo(fileInfo, destinationFolderId);

  }
  
  function moveFileTo(fileInfo, destinationFolderId) {
    var targetFile = DriveApp.getFileById(fileInfo.fileId);
    var fullName = targetFile.getName();
    var fileName = fullName.substr(0, fullName.indexOf(' '));
    var randomNumber = new Date(fileInfo.date).getTime();
    
    var extension = fullName.substr(fullName.indexOf('.'), fullName.length);
    
    if (targetFile) {
      var newFileName = randomNumber + '[' + fileInfo.mssv + '][' + fileName + ']' + extension;
      
      var parentFolders = targetFile.getParents();
      while(parentFolders.hasNext()) {
        var parent = parentFolders.next();
        parent.removeFile(targetFile);
      }
      
      var destinationFolder = DriveApp.getFolderById(destinationFolderId);
      destinationFolder.addFile(targetFile);
      
      targetFile.setName(newFileName);  
    }
  }
  
}
