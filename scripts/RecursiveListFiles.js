/**
* Recursive function that lists all files within the specified folder, as well as all subfolders.
* @param {string} id - the ID of the parent folder to perform the recursive search on. 
*/
function recursiveListFiles(id) {
  const _FOLDERS = DriveApp.getFolderById(id).getFolders();
  const _FILES = DriveApp.getFolderById(id).getFiles();
  let allFiles = [];
  
  // pushes immediate files in parent folder into allFiles list.
  while (_FILES.hasNext()) {
    let file = _FILES.next();
    let name = file.getName();
    allFiles.push(name);
  }
  
  // iterate through every folder in the parent folder. 
  while (_FOLDERS.hasNext()) {
    let foldersNext = _FOLDERS.next();
    // iterate through every subfolder and get all files and subfolders.
    listFilesAndSubfolders(foldersNext);
  }
  
  function listFilesAndSubfolders(folder) {
    // get all files in current folder.
    listFiles(folder);
    // go through each subfolder in current folder.
    let subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      let subfolder = subfolders.next();
      // run this function again on all subfolders inside the given folder.
      listFilesAndSubfolders(subfolder);
    }
  }
  
  // get all files in the given folder.
  function listFiles(foldersNext) {
    let files = foldersNext.getFiles();
    while (files.hasNext()) {
      let file = files.next();
      let name = file.getName();
      allFiles.push(name);
    }
  }
  return allFiles;
}
