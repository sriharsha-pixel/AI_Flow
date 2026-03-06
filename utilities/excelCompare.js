const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const { getMismatches } = require("./getMismatches");
const { writeDataToExcel } = require("./readExcel");

// Read all sheets and return cleaned JSON data
function getAllSheetsData(filePath) {
    const wb = XLSX.readFile(filePath);
    const allData = {};

    wb.SheetNames.forEach(sheet => {
        const ws = wb.Sheets[sheet];
        allData[sheet] = XLSX.utils.sheet_to_json(ws, { defval: "" }).map(row => {
            const cleaned = {};
            for (let key in row) {
                cleaned[key.trim()] = String(row[key]).trim();
            }
            return cleaned;
        });
    });

    return allData;
}

// Compare two Excel files sheet-wise and save mismatch report
async function compareExcelsSheetWise(previewFile, prodFile) {
    const previewData = getAllSheetsData(previewFile);
    const prodData = getAllSheetsData(prodFile);

    const mismatchFile = path.join(process.cwd(), "output", "Mismatch_Report_SheetWise.xlsx");

    if (fs.existsSync(mismatchFile)) fs.unlinkSync(mismatchFile);

    let hasAnyMismatch = false;

    for (const sheetName of Object.keys(prodData)) {
        const prodSheet = prodData[sheetName] || [];
        const previewSheet = previewData[sheetName] || [];

        const mismatches = getMismatches(prodSheet, previewSheet);

        if (mismatches.length > 0) {
            hasAnyMismatch = true;
            writeDataToExcel(mismatchFile, sheetName, mismatches);
            console.log(`Mismatches found in sheet "${sheetName}"`);
        } else {
            console.log(`No mismatches in sheet "${sheetName}"`);
        }
    }

    if (hasAnyMismatch) {
        console.log(`All mismatches saved in: ${mismatchFile}`);
    }

    return hasAnyMismatch;
}

module.exports = { getAllSheetsData, compareExcelsSheetWise };