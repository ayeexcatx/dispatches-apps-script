const DISPATCH_ARCHIVES_FOLDER_ID = '1_idlRE8_vOcQSEX97k9iS1huoWf-ZbRb';

function doGet(e) {
  const companyRaw = e.parameter.company;
  if (!companyRaw) {
    return ContentService.createTextOutput('Error: Missing company parameter.');
  }

  const company = companyRaw;
  const fileName = `${company}_dispatch_list.html`;
  const archivesFolder = DriveApp.getFolderById(DISPATCH_ARCHIVES_FOLDER_ID);
  const file = findFileRecursively(archivesFolder, fileName);

  if (!file) {
    Logger.log(`❌ File ${fileName} not found`);
    return ContentService.createTextOutput(`No dispatch list found for company: ${company}.`);
  }

  const blob = file.getBlob();
  const content = blob.getDataAsString();

  Logger.log(`📦 File size: ${blob.getBytes().length} bytes`);
  Logger.log(`🧪 First 1000 characters: ${content.slice(0, 1000)}`);

  return HtmlService.createHtmlOutput(content)
    .setTitle(`${company} Dispatch List`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function findFileRecursively(folder, targetFileName) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName() === targetFileName) {
      return file;
    }
  }
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const result = findFileRecursively(subfolder, targetFileName);
    if (result) return result;
  }
  return null;
}
