/* ****************************************************************************
 * This Google Apps Script is used to copy a directory, recursively, in Drive.
 * It should be uploaded & run from script.google.com. getContinuationToken is
 * complicated, at best, for a recursive directory copy, so it skips 'large'
 * files (1+ MB) in order to help stay under the GAS execution time limit.
 * Multiple runs may be necessary (adjusting sourceId and targetId to
 * subfolders).
 * Author: Cody Voight
 * Version: 1.0
 * ***************************************************************************/

function main() {
  const sourceId = "source_folder_ID";
  const targetId = "target_folder_ID";
  const source = DriveApp.getFolderById(sourceId);
  const target = DriveApp.getFolderById(targetId);
  copyDirectory(source, target);
}

function copyDirectory(source, target) {
  const sourceFiles = source.getFiles();
  while (sourceFiles.hasNext()) {
    sourceFile = sourceFiles.next();
    //skip files >1 MB
    if (sourceFile.getSize() > 1000000) {
      continue;
    }
    sourceFile.makeCopy(sourceFile.getName(), target);
  }
  const sourceFolders = source.getFolders();
  while (sourceFolders.hasNext()) {
    const sourceFolder = sourceFolders.next();
    const targetFolder = target.createFolder(sourceFolder.getName());
    copyDirectory(sourceFolder, targetFolder);
  }
}
