// delete dispatches after 100 days
function cleanupOldDispatches() {
  const parentFolder = DriveApp.getFolderById(DISPATCH_ARCHIVES_FOLDER_ID);
  const now = new Date();
  const cutoff = new Date(now.getTime() - 100 * 24 * 60 * 60 * 1000); // 100 days ago

  const companyFolders = parentFolder.getFolders();
  while (companyFolders.hasNext()) {
    const companyFolder = companyFolders.next();
    const truckFolders = companyFolder.getFolders();
    while (truckFolders.hasNext()) {
      const truckFolder = truckFolders.next();
      const files = truckFolder.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        const name = file.getName();

        if (name.endsWith('.html')) continue;

        const match = name.match(/^(\d{4}-\d{2}-\d{2})_(\d{4})_/);
        if (match) {
          const fileDate = new Date(match[1]);
          if (fileDate < cutoff) {
            Logger.log(`🗑️ Deleting ${name} from ${truckFolder.getName()}`);
            file.setTrashed(true); // Safe: moves to trash
          }
        }
      }
    }
  }
}

  // Delete files from Replaced Dispatches folders automatically after 100 days
function deleteOldReplacedDispatches() {
  const DISPATCH_ARCHIVES_FOLDER_ID = "1_idlRE8_vOcQSEX97k9iS1huoWf-ZbRb"; // Main Dispatch Archives folder
  const DAYS_TO_KEEP = 100;
  const now = new Date();
  const cutoffDate = new Date(now.getTime() - DAYS_TO_KEEP * 24 * 60 * 60 * 1000);

  const mainFolder = DriveApp.getFolderById(DISPATCH_ARCHIVES_FOLDER_ID);
  const companyFolders = mainFolder.getFolders();

  while (companyFolders.hasNext()) {
    const companyFolder = companyFolders.next();
    const truckFolders = companyFolder.getFolders();

    while (truckFolders.hasNext()) {
      const truckFolder = truckFolders.next();
      const replacedFolderIterator = truckFolder.getFoldersByName("Replaced Dispatches");

      if (!replacedFolderIterator.hasNext()) continue;

      const replacedFolder = replacedFolderIterator.next();
      const files = replacedFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const created = file.getDateCreated();

        if (created < cutoffDate) {
          Logger.log(`Deleting old file: ${file.getName()} (created: ${created})`);
          file.setTrashed(true);
        }
      }
    }
  }
}
  // End of Delete files from Replaced Dispatches folders