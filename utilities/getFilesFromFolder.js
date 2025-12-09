const fs = require("fs");
const path = require("path");

// Helper to get all files in a folder
function getFilesFromFolder(folderPath) {
  const absFolderPath = path.resolve(__dirname, folderPath);
  return fs.readdirSync(absFolderPath).map(file => path.join(absFolderPath, file));
}

module.exports={getFilesFromFolder};