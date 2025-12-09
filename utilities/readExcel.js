const XLSX = require("xlsx");
const fs = require("fs");


function writeDataToExcel(filename, sheetName, jsonData) {
   let wb;
  // if (fs.existsSync(filename)) {
  //   wb = XLSX.readFile(filename);
  // } else {
  //   wb = XLSX.utils.book_new();
  // }
  // const ws = XLSX.utils.json_to_sheet(jsonData);
  // if (wb.SheetNames.includes(sheetName)) {
  //   delete wb.Sheets[sheetName];
  //   wb.SheetNames = wb.SheetNames.filter((name) => name !== sheetName);
  // }
  // XLSX.utils.book_append_sheet(wb, ws, sheetName);
  // XLSX.writeFile(wb, filename);
  // console.log(` Data written to ${filename} (${sheetName})`);
  if (fs.existsSync(filename)) {
  wb = XLSX.readFile(filename);

  // Delete existing sheet if present
  if (wb.SheetNames.includes(sheetName)) {
    delete wb.Sheets[sheetName];
    wb.SheetNames = wb.SheetNames.filter((name) => name !== sheetName);
    console.log(`Old sheet "${sheetName}" deleted from ${filename}`);
  }
} else {
  wb = XLSX.utils.book_new();
}

const ws = XLSX.utils.json_to_sheet(jsonData);
XLSX.utils.book_append_sheet(wb, ws, sheetName);
XLSX.writeFile(wb, filename);
}



function readDataFromExcel(filename, sheetName) {
  if (!fs.existsSync(filename)) throw new Error("Excel file not found!");
  
  const wb = XLSX.readFile(filename);
  const ws = wb.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(ws, { defval: "" }); // defval prevents undefined
  
  return jsonData.map(row => {
    const cleaned = {};
    for (let key in row) {
      const trimmedKey = key.trim();              // remove extra spaces in column names
      const trimmedValue = String(row[key]).trim(); // normalize values as string
      cleaned[trimmedKey] = trimmedValue;
    }
    return cleaned;
  });
}


module.exports = { writeDataToExcel, readDataFromExcel };
