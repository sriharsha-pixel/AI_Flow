const { excuteSteps } = require("../../utilities/actions");
const { test,expect } = require("@playwright/test");
const {extractTextFromImage} =require("../../utilities/extractTextFromImage");
const path = require("path");
const {writeDataToExcel,readDataFromExcel} = require("../../utilities/readExcel");
const {getMismatches} =require("../../utilities/getMismatches");
require("dotenv").config();
const XLSX = require("xlsx");
const fs = require("fs");
const {scrollToElement}=require("../../utilities/scrollInView");
const{getFilesFromFolder}=require("../../utilities/getFilesFromFolder");

exports.CashPosting = class CashPosting {
    constructor(test, page) {
        this.test=test;
        this.page=page;
        this.firstcard=page.locator("(//div[@data-radix-scroll-area-content]/div/div)[1]");
        this.cashPostinggetStartedBtn=page.locator("//h3[text()='Cash Posting Reconciliation']/following::button[1]");
        this.bankStatementFileUploadBtn=page.locator("//input[@id='bank-file-input']");
        this.rfmsFileUploadBtn=page.locator("//input[@id='rfms-file-input']");
        this.billingSystemfileUploadBtn=page.locator("//input[@id='billing-file-input']");
        this.runReconsillationBtn=page.locator("//button[text()='Run Reconciliation']");
        this.summaryCard=page.locator("(//div[@data-component-file='ReconSummary.tsx'])[1]");
        this.summaryCardProd=page.locator("(//div[@id='results-section']/div/div)[2]");
        this.matchedTransactionsHeader=page.locator("//h3[text()='Matched Transactions']");
        this.totalTransaction=page.locator("((//div[@id='results-section']/div/div)[2]/div)[1]//span[text()='Total Transactions:']/following-sibling::div");
        this.matched=page.locator("((//div[@id='results-section']/div/div)[2]/div)[1]//span[text()='Matched:']/following-sibling::div");
        this.transfers=page.locator("((//div[@id='results-section']/div/div)[2]/div)[1]//span[text()='Transfers:']/following-sibling::div");
        this.ndc=page.locator("((//div[@id='results-section']/div/div)[2]/div)[1]//span[text()='NDC Sweeps:']/following-sibling::div");
        this.unmatched=page.locator("((//div[@id='results-section']/div/div)[2]/div)[1]//span[text()='Unmatched:']/following-sibling::div/div");
        this.matchrate=page.locator("((//div[@id='results-section']/div/div)[2]/div)[1]//span[text()='Match Rate:']/following-sibling::div/span");
        this.processing=page.locator("//span[contains(text(),'Processing Reconciliation')]");
        // NewChanges
        this.header_cashReceiptsInBIllingSystem =   page.locator("//h3[contains(text(),'Cash Receipts in Billing System not Found in Bank')]") 
        this.tbl_cashReceiptsInBIllingSystem=page.locator("//h3[contains(text(),'Cash Receipts in Billing System not Found in Bank')]//following::table[1]")
        this.tbl_cashReceiptsInBIllingSystemAll =page.locator("//h3[contains(text(),'Cash Receipts in Billing System not Found in Bank')]//following::table[1]//tr")
        
        this.header_cashReceiptsInBIllingSystemDropDown= page.locator("//h3[contains(text(),'Cash Receipts in Billing System not Found in Bank')]//following::*[@data-lov-name='ChevronDown']")

        this.headerNSFTransactions  =   page.locator("//h3[contains(text(),'NSF Transactions')]") 
        this.header_NSFTransactionsDropDownIon =   page.locator("//h3[contains(text(),'NSF Transactions')]//following::*[@data-lov-name='ChevronDown']") 
        this.tbl_NSFTransactions    =   page.locator("//h3[contains(text(),'NSF Transactions')]//following::table[1]")
        this.tbl_NSFTransactionsAll    =   page.locator("//h3[contains(text(),'NSF Transactions')]//following::table[1]//tr")
        
        this.headerDepositsInBankRFMS                   =    page.locator("//h3[text()='Deposits in Bank/RFMS Not Found in Billing System']")
        this.headerDepositsInBankRFMSDropDownIcon       =     page.locator("//h3[text()='Deposits in Bank/RFMS Not Found in Billing System']//following::*[@data-lov-name='ChevronDown']")
        this.tblDepositsInBankRFMS                      =    page.locator("//h3[text()='Deposits in Bank/RFMS Not Found in Billing System']//following::table[1]")
        this.tblDepositsInBankRFMSAll                      =    page.locator("//h3[text()='Deposits in Bank/RFMS Not Found in Billing System']//following::table[1]//tr")

        this.headerReversalTransactions                   =    page.locator("//h3[text()='Reversal Transactions']")
        this.headerReversalTransactionsDropDownIcon       =     page.locator("//h3[text()='Reversal Transactions']//following::*[@data-lov-name='ChevronDown']")
        this.tbl_ReversalTransactions                      =    page.locator("//h3[text()='Reversal Transactions']//following::table[1]")
        this.tbl_ReversalTransactionsAll                      =    page.locator("//h3[text()='Reversal Transactions']//following::table[1]//tr")


        this.headerInternalBankTransfers                   =    page.locator("//h3[text()='Internal Bank Transfers']")
        this.headerInternalBankTransfersDropDownIcon       =     page.locator("//h3[text()='Internal Bank Transfers']//following::*[@data-lov-name='ChevronDown']")
        this.tbl_InternalBankTransfers                      =    page.locator("//h3[text()='Internal Bank Transfers']//following::table[1]")
        this.tbl_InternalBankTransfersAll                      =    page.locator("//h3[text()='Internal Bank Transfers']//following::table[1]//tr")


        this.headerNDCSweeps                   =    page.locator("//h3[text()='NDC Sweeps']")
        this.headerNDCSweepsDropDownIcon       =     page.locator("//h3[text()='NDC Sweeps']//following::*[@data-lov-name='ChevronDown']")
        this.tbl_NDCSweeps                      =    page.locator("//h3[text()='NDC Sweeps']//following::table[1]")
        this.tbl_NDCSweepsAll                      =    page.locator("//h3[text()='NDC Sweeps']//following::table[1]//tr");        

        this.lbl_BSA_TotalTransactions = page.locator("//h3[text()='Billing System Analysis']//following::span[text()='Total Transactions:']//following::div[1]")
        this.lbl_BSA_MatchedTransactions= page.locator("//h3[text()='Billing System Analysis']//following::span[text()='Matched:']//following::div[1]//div[1]")
        this.lbl_BSA_NSF= page.locator("//h3[text()='Billing System Analysis']//following::span[text()='NSFs:']//following::div[1]")
        this.lbl_BSA_Reversals= page.locator("//h3[text()='Billing System Analysis']//following::span[text()='Reversals:']//following::div[1]")
        this.lbl_BSA_unMatched= page.locator("//h3[text()='Billing System Analysis']//following::span[text()='Unmatched:']//following::div[1]//div[1]")
        this.lbl_BSA_matchRate = page.locator("//h3[text()='Billing System Analysis']//following::span[text()='Match Rate:']//following::span[1]")
        
        this.lbl_TotalMatchesFound_total= page.locator("(//h3[text()='Total Matches Found']//following::div//div//div)[1]")
        
        
        this.lbl_TotalMatchesFound_ExactMatches = page.locator("(//h3[text()='Total Matches Found']//following::div//div//div//div)[1]")
        
        this.lbl_TotalMatchesFound_oneToMany = page.locator("(//h3[text()='Total Matches Found']//following::div//div//div//div)[2]")
        
        this.lbl_TotalMatchesFound_manyToMany =page.locator("(//h3[text()='Total Matches Found']//following::div//div//div//div)[3]") 
        this.lbl_ReconStatus_totalDeposits = page.locator("//h3[text()='Reconciliation Status']//following::span[contains(text(), 'Total Deposits:')][1]//following::span[1]")
        this.lbl_ReconStatus_totalCashReceipts = page.locator("//h3[text()='Reconciliation Status']//following::span[contains(text(), 'Total Cash Receipts:')][1]//following::span[1]")
        this.lbl_ReconStatus_differences = page.locator("//h3[text()='Reconciliation Status']//following::span[contains(text(), 'Difference:')][1]//following::span[1]")
    
        this.lbl_cashReconciliationReport   =   page.locator("//h3[text()='Cash Reconciliation Report']")
        this.tbl_cashReconciliationReport_TotalBillingTransactions =  page.locator("//h3[text()='Cash Reconciliation Report']/following::span[count(.|//h3[text()='Bank Deposits by Account:']/preceding::span)=count(//h3[text()='Bank Deposits by Account:']/preceding::span)]")
        this.header_cashReconciliationReport_BankDepositsByAccount = page.locator("//h3[text()='Bank Deposits by Account:']")
        this.header_cashReconciliationReport_BillingSystemAdjustment =page.locator("//h3[text()='Billing System Adjustments:']")
        this.tbl_cashReconciliationReport_BankDepositsByAccount = page.locator("//h3[text()='Bank Deposits by Account:']//following::div[1]//div//span")
        this.header_ReconciliationReport_Deductions = page.locator("//h3[text()='Deductions:']")
        this.tbl_cashReconciliationReport_Deductions = page.locator("//h3[text()='Deductions:']//following::div[1]//div//span")
        this.tbl_cashReconciliationReport_BillingSystemAdjustments =page.locator("//h3[text()='Billing System Adjustments:']//following::div[1]//div//span")
        this.header_UnMatchedBillingSystem = page.locator("//div[text()='Unmatched Billing System Transactions:']")
        this.tbl_cashReconciliationReport_unMatchedBillingSystemTransactions = page.locator("//div[text()='Unmatched Billing System Transactions:']//following::div[1]//span")
        this.header_cashReconciliationReport_Reconciliation=page.locator("//h3[text()='Reconciliation:']")
        this.tbl_cashReconciliationReport_Reconciliation= page.locator("//h3[text()='Reconciliation:']//following::div[1]//div//span")
    
    
    
    
    }

    clickOnFirstCard=async()=>{
        await excuteSteps(this.test,this.firstcard,"click",`Clicking on first card`);
    }

    scrollTillRunReconsillationBtn=async()=>{
        await excuteSteps(this.test,this.runReconsillationBtn,"scroll",`Scrolling into view`);
    }
    scrollTillReversalTransactions=async()=>{
        await excuteSteps(this.test,this.headerReversalTransactions,"scroll",`Scrolling into view`);
    }

    

    clickOnreconsillationBtn=async()=>{
        await excuteSteps(this.test,this.runReconsillationBtn,"click",`Clicking on reconsilation button after file upload`);
    };

    clickOnMatchedTransactions=async()=>{
        await excuteSteps(this.test,this.matchedTransactionsHeader,"click",`Clicking on matched transactions`)

    };
    scrollToHeaderNSFTransactions=async()=>{
        await excuteSteps(this.test,this.headerNSFTransactions ,"scroll",`Scrolling into view`);
    }
    
    clickOnCashPostingBtn=async()=>{
        await excuteSteps(this.test,this.cashPostinggetStartedBtn,"click",`Clicking on cash posting button`);
    };

    clickOnInternalBankTransferHeader=async()=>{
        await excuteSteps(this.test,this.headerInternalBankTransfers,"click",`Clicking on internal bank transfers`);
    };

     
    uploadingFilesInTest=async(bankStatementfilefolder,rfmsFileFolder,billingSystemFileFolder)=>{
        await this.clickOnCashPostingBtn();
        await this.page.waitForTimeout(parseInt(process.env.smallWait));
        const bankStatementFiles = getFilesFromFolder(bankStatementfilefolder);
        await this.bankStatementFileUploadBtn.setInputFiles(bankStatementFiles);
        const rfmsFiles = getFilesFromFolder(rfmsFileFolder);
        await this.rfmsFileUploadBtn.setInputFiles(rfmsFiles);
        const billingFiles = getFilesFromFolder(billingSystemFileFolder);
        await this.billingSystemfileUploadBtn.setInputFiles(billingFiles);
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        await this.scrollTillRunReconsillationBtn();
        await this.clickOnreconsillationBtn();
        await this.processing.waitFor({ state: "hidden" });
        await this.page.waitForTimeout(parseInt(process.env.largeWait));
        await this.summaryCard.screenshot({path:'screenshots/test.png'});
    }

    uploadingFilesInProd=async(bankStatementfilefolder,rfmsFileFolder,billingSystemFileFolder)=>{
        await this.clickOnCashPostingBtn();
        await this.page.waitForTimeout(parseInt(process.env.smallWait));
        const bankStatementFiles = getFilesFromFolder(bankStatementfilefolder);
        await this.bankStatementFileUploadBtn.setInputFiles(bankStatementFiles);
        const rfmsFiles = getFilesFromFolder(rfmsFileFolder);
        await this.rfmsFileUploadBtn.setInputFiles(rfmsFiles);
        const billingFiles = getFilesFromFolder(billingSystemFileFolder);
        await this.billingSystemfileUploadBtn.setInputFiles(billingFiles);

        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        await this.scrollTillRunReconsillationBtn();
        await this.clickOnreconsillationBtn();
        await this.page.waitForTimeout(parseInt(process.env.largeWait));
    }

    matchedTransactionsToExcelTest=async(path)=>{
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        await this.clickOnMatchedTransactions();
        const rows = await this.page.locator("((//h3[text()='Matched Transactions']/following::table)[1]//tr)").all();
        const tableData = [];
        console.log(rows);

        for (let i = 1; i < rows.length; i++) {
            const cells = await rows[i].locator("td").allInnerTexts();
            tableData.push({
            BillingSysDate: cells[0].trim(),
            BankDate: cells[1].trim(),
            BillingSysAmt: cells[2].trim(),
            BankAmt: cells[3].trim(),
            BillingSysDesc: cells[4].trim(),
            BankRFMSDesc: cells[5].trim(),
            MatchType: cells[6].trim(),
            });
        }

        writeDataToExcel(path, "MatchedTransactions", tableData);
    }

    
    compareAndExportMismatch=async(testEnvData,mismatchSheetName,mismatchFile)=>{
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        await this.clickOnMatchedTransactions();
        await this.page.waitForTimeout(parseInt(process.env.largeWait));
        const rows = await this.page.locator("((//h3[text()='Matched Transactions']/following::table)[1]//tr)").all();
        console.log("now ",rows.length);
        const prodData = [];

        for (let i = 1; i < rows.length; i++) {
            const cells = await rows[i].locator("td").allInnerTexts();
            prodData.push({
            BillingSysDate: cells[0].trim(),
            BankDate: cells[1].trim(),
            BillingSysAmt: cells[2].trim(),
            BankAmt: cells[3].trim(),
            BillingSysDesc: cells[4].trim(),
            BankRFMSDesc: cells[5].trim(),
            MatchType: cells[6].trim(),
            });
        }

        const testData = readDataFromExcel(testEnvData, "MatchedTransactions");
        let mismatches = [];

        for (let i = 0; i < Math.max(prodData.length, testData.length); i++) {
            const prodRow = prodData[i];
            const testRow = testData[i];
            if (!prodRow || !testRow) {
            mismatches.push({
                Row: i + 1,
                Column: "N/A",
                TestEnvValue: testRow ? JSON.stringify(testRow) : "Missing Row",
                ProdEnvValue: prodRow ? JSON.stringify(prodRow) : "Missing Row",
                TestEnvRow: testRow ? JSON.stringify(testRow) : "N/A",
                ProdEnvRow: prodRow ? JSON.stringify(prodRow) : "N/A",
            });
            continue;
            }

            Object.keys(prodRow).forEach((key) => {
            if (prodRow[key] !== testRow[key]) {
                mismatches.push({
                Row: i + 1,
                Column: key,
                TestEnvValue: testRow[key],
                ProdEnvValue: prodRow[key],
                TestEnvRow: JSON.stringify(testRow),
                ProdEnvRow: JSON.stringify(prodRow),
                });
            }
            });
        }
        if (mismatches.length > 0) {
            console.log("Mismatches found. Exporting detailed report to Excel...");

            writeDataToExcel(mismatchFile, mismatchSheetName, mismatches);
            expect.soft(mismatches.length, "Matched Transactions mismatch count").toBe(0);
        } else {
            console.log("All data matched!");

    //  If the file exists, check if the mismatch sheet exists
    if (fs.existsSync(mismatchFile)) {
      const wb = XLSX.readFile(mismatchFile);

      if (wb.SheetNames.includes(mismatchSheetName)) {
        delete wb.Sheets[mismatchSheetName];
        wb.SheetNames = wb.SheetNames.filter(
          (name) => name !== mismatchSheetName
        );
        if (wb.SheetNames.length === 0) {
                    const emptySheet = XLSX.utils.aoa_to_sheet([[""]]);
                    wb.SheetNames.push("Sheet1");
                    wb.Sheets["Sheet1"] = emptySheet;
                }
        XLSX.writeFile(wb, mismatchFile);
        console.log(`Old mismatch sheet "${mismatchSheetName}" deleted.`);
      }
    }

    await test.step("Verified, All test data matched", async () => {
      console.log("All data matched, old mismatch sheet removed if present!");
    });

        }
    }

    writingBankFeedSummaryTest=async(path)=>{
        const totalTransactions = await this.totalTransaction.innerText();
        const matched = await this.matched.innerText();
        const transfers = await this.transfers.innerText();
        const ndcSweeps = await this.ndc.innerText();
        const unmatched = await this.unmatched.innerText();
        const matchRate = await this.matchrate.innerText();
        const testData = [{
        "TotalTransactions": totalTransactions,
        "Matched": matched,
        "Transfers": transfers,
        "NDCSweeps": ndcSweeps,
        "Unmatched": unmatched,
        "MatchRate": matchRate
        }];

    writeDataToExcel(path, "Bank Feed Analysis", testData);

    console.log("Data written to TestEnvData.xlsx");
    }
    
    gettingBandFeedDetailsAndCompareProd =async(testEnvData,mismatchSheetName,mismatchFile)=>{
        const totalTransactions = await this.totalTransaction.innerText();
        const matched = await this.matched.innerText();
        const transfers = await this.transfers.innerText();
        const ndcSweeps = await this.ndc.innerText();
        const unmatched = await this.unmatched.innerText();
        const matchRate = await this.matchrate.innerText();

        const prodData = [{
        "TotalTransactions": totalTransactions,
        "Matched": matched,
        "Transfers": transfers,
        "NDCSweeps": ndcSweeps,
        "Unmatched": unmatched,
        "MatchRate": matchRate
        }];

        console.log("first card,prod data",prodData);

        const testData = readDataFromExcel(testEnvData, "Bank Feed Analysis");
        console.log("test data",testData);
        const mismatches = getMismatches(prodData, testData);
        //const mismatchFile = path.resolve("./excel/Mismatches/BankFeedMissMatch.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found:", mismatches);
            writeDataToExcel(mismatchFile, mismatchSheetName, mismatches);
            expect.soft(mismatches.length, "Mismatches found in data").toBe(0);
        } 
        else {
            console.log("All data matched!");
        if (fs.existsSync(mismatchFile)) {
        const wb = XLSX.readFile(mismatchFile);

        if (wb.SheetNames.includes(mismatchSheetName)) {
            delete wb.Sheets[mismatchSheetName];
            wb.SheetNames = wb.SheetNames.filter(
            (name) => name !== mismatchSheetName
            );
            XLSX.writeFile(wb, mismatchFile);
            console.log(`Old mismatch sheet "${mismatchSheetName}" deleted.`);
        }
    }

    await test.step("Verified, All test data matched", async () => {
      console.log("All data matched, old mismatch sheet removed if present!");
    });     
        }
    }
    scrollToHeaderNSFTransactions=async()=>{
        await excuteSteps(this.test,this.headerNSFTransactions ,"scroll",`Scrolling into view`);
    }

    scrollToHeaderCashReceiptBilling=async()=>{
        await excuteSteps(this.test,this.header_cashReceiptsInBIllingSystem,"scroll",`Scrolling into view`);
    }
    CashReceiptsInBillingSystemNotFoundInBankExcelTest=async(path)=>{
        await this.scrollToHeaderCashReceiptBilling()
        if  (await this.tbl_cashReceiptsInBIllingSystem.isVisible()){
            await expect.soft(this.tbl_cashReceiptsInBIllingSystem).toBeVisible()
        }else{
            await this.header_cashReceiptsInBIllingSystem.click()
            await expect.soft(this.tbl_cashReceiptsInBIllingSystem).toBeVisible()
        }
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        await this.page.locator
        const rows = await this.tbl_cashReceiptsInBIllingSystemAll.all();
        const tableData = [];
        console.log(rows);

        for (let i = 1; i < rows.length; i++) {
            const cells = await rows[i].locator("td").allInnerTexts();
            tableData.push({
            BillingSysDate: cells[0].trim(),
            Amount: cells[1].trim(),
            BillingSysDesc: cells[2].trim(),
            Status: cells[3].trim(),
            });
        }

        writeDataToExcel(path, "UnmatchedBilling", tableData);
    }

    
    compareAndExportCashReceiptsInBillingSystem = async (testEnvData,mismatchSheetName,mismatchFile) => {
    await scrollToElement(this.header_cashReceiptsInBIllingSystem);
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    const rows = await this.tbl_cashReceiptsInBIllingSystemAll.all();
    const prodData = [];
    for (let i = 1; i < rows.length; i++) {
        const cells = await rows[i].locator("td").allInnerTexts();
        prodData.push({
            BillingSysDate: cells[0]?.trim() || "",
            Amount: cells[1]?.trim() || "",
            BillingSysDesc: cells[2]?.trim() || "",
            Status: cells[3]?.trim() || "",
        });
    }

    const testData = readDataFromExcel(testEnvData, "UnmatchedBilling");
    console.log(testData);
    let mismatches = [];

    for (let i = 0; i < Math.max(prodData.length, testData.length); i++) {
        const prodRow = prodData[i];
        const testRow = testData[i];
        if (!prodRow || !testRow) {
            mismatches.push({
                Row: i + 1,
                Column: "N/A",
                Test_BillingSysDate: testRow?.BillingSysDate || "Missing",
                Test_Amount: testRow?.Amount || "Missing",
                Test_BillingSysDesc: testRow?.BillingSysDesc || "Missing",
                Test_Status: testRow?.Status || "Missing",
                Prod_BillingSysDate: prodRow?.BillingSysDate || "Missing",
                Prod_Amount: prodRow?.Amount || "Missing",
                Prod_BillingSysDesc: prodRow?.BillingSysDesc || "Missing",
                Prod_Status: prodRow?.Status || "Missing",
            });
            continue;
        }

        Object.keys(prodRow).forEach((key) => {
            if (prodRow[key] !== testRow[key]) {
                mismatches.push({
                    Row: i + 1,
                    Column: key,
                    Test_BillingSysDate: testRow?.BillingSysDate || "N/A",
                    Test_Amount: testRow?.Amount || "N/A",
                    Test_BillingSysDesc: testRow?.BillingSysDesc || "N/A",
                    Test_Status: testRow?.Status || "N/A",
                    Prod_BillingSysDate: prodRow?.BillingSysDate || "N/A",
                    Prod_Amount: prodRow?.Amount || "N/A",
                    Prod_BillingSysDesc: prodRow?.BillingSysDesc || "N/A",
                    Prod_Status: prodRow?.Status || "N/A",
                });
            }
        });
    }

    //const mismatchFile = path.resolve("./excel/Mismatches/UnmatchedBilling_Mismatch.xlsx");

    if (mismatches.length > 0) {
        console.log("Mismatches found. Exporting detailed report to Excel...");
        writeDataToExcel(mismatchFile, mismatchSheetName, mismatches);
        expect.soft(mismatches.length, "UnmatchedBilling mismatch count").toBe(0);
    } else {

                    console.log("All data matched!");
                if (fs.existsSync(mismatchFile)) {
                const wb = XLSX.readFile(mismatchFile);

                if (wb.SheetNames.includes(mismatchSheetName)) {
                    delete wb.Sheets[mismatchSheetName];
                    wb.SheetNames = wb.SheetNames.filter(
                    (name) => name !== mismatchSheetName
                    );
                    XLSX.writeFile(wb, mismatchFile);
                    console.log(`Old mismatch sheet "${mismatchSheetName}" deleted.`);
                }
                }

                await test.step("Verified, All test data matched", async () => {
                console.log("All data matched, old mismatch sheet removed if present!");
                });

    }
};

    scrollToDepositInBankTransaction=async()=>{
        await excuteSteps(this.test,this.headerDepositsInBankRFMS ,"scroll",`Scrolling into view`);
    }
    DepositsInBankRFMSNotFoundTransactionsToExcelTest=async(path)=>{
        await this.scrollToDepositInBankTransaction()
        if  (await this.tblDepositsInBankRFMS.isVisible())
            {
                await expect.soft(this.tblDepositsInBankRFMS).toBeVisible()
            }else
            {
                await this.headerDepositsInBankRFMS.click()
                await expect.soft(this.tblDepositsInBankRFMS).toBeVisible()
            }
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        const rows = await this.tblDepositsInBankRFMSAll.all();
        const tableData = [];
        console.log(rows);

        for (let i = 1; i < rows.length; i++) {
            const cells = await rows[i].locator("td").allInnerTexts();
            tableData.push({
            BankDate: cells[0].trim(),
            Amount: cells[1].trim(),
            BankRFMSDesc: cells[2].trim(),
            Status: cells[3].trim(),
            });
        }

        writeDataToExcel(path, "UnmatchedBankDeposits", tableData);
    }

    compareAndExportDepositsInBankRFMSNotFoundTransactions=async(testEnvData,mismatchSheetName,mismatchFile)=>{
        await this.scrollToDepositInBankTransaction()
        if  (await this.tblDepositsInBankRFMS.isVisible()){
            await expect.soft(this.tblDepositsInBankRFMS).toBeVisible()
        }else{
            await this.headerDepositsInBankRFMS.click()
            await expect.soft(this.tblDepositsInBankRFMS).toBeVisible()
        }
        
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        const rows = await this.tblDepositsInBankRFMSAll.all();
        const prodData = [];

        for (let i = 1; i < rows.length; i++) {
            const cells = await rows[i].locator("td").allInnerTexts();
            prodData.push({
            BankDate: cells[0].trim(),
            Amount: cells[1].trim(),
            BankRFMSDesc: cells[2].trim(),
            Status: cells[3].trim(),
            
            });
        }

        const testData = readDataFromExcel(testEnvData, "UnmatchedBankDeposits");
        let mismatches = [];

        for (let i = 0; i < Math.max(prodData.length, testData.length); i++) {
            const prodRow = prodData[i];
            const testRow = testData[i];

            if (!prodRow || !testRow) {
                mismatches.push({
                    Row: i + 1,
                    Column: "N/A",
                    Test_BankDate: testRow?.BankDate || "Missing",
                    Test_Amount: testRow?.Amount || "Missing",
                    Test_BankRFMSDesc: testRow?.BankRFMSDesc || "Missing",
                    Test_Status: testRow?.Status || "Missing",
                    Prod_BankDate: prodRow?.BankDate || "Missing",
                    Prod_Amount: prodRow?.Amount || "Missing",
                    Prod_BankRFMSDesc: prodRow?.BankRFMSDesc || "Missing",
                    Prod_Status: prodRow?.Status || "Missing",
                });
                continue;
            }

            Object.keys(prodRow).forEach((key) => {
                if (prodRow[key] !== testRow[key]) {
                    mismatches.push({
                        Row: i + 1,
                        Column: key,
                        Test_BankDate: testRow?.BankDate || "N/A",
                        Test_Amount: testRow?.Amount || "N/A",
                        Test_BankRFMSDesc: testRow?.BankRFMSDesc || "N/A",
                        Test_Status: testRow?.Status || "N/A",
                        Prod_BankDate: prodRow?.BankDate || "N/A",
                        Prod_Amount: prodRow?.Amount || "N/A",
                        Prod_BankRFMSDesc: prodRow?.BankRFMSDesc || "N/A",
                        Prod_Status: prodRow?.Status || "N/A",
                    });
                }
            });
        }

        
        //const mismatchFile = path.resolve("./excel/Mismatches/UnmatchedBankDeposits_Mismatch.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found. Exporting detailed report to Excel...");
            writeDataToExcel(mismatchFile, mismatchSheetName, mismatches);
             expect.soft(mismatches.length, "UnmatchedBankDeposits mismatch count").toBe(0);
        } else {
                    console.log("All data matched!");

            // ✅ If the file exists, check if the mismatch sheet exists
            if (fs.existsSync(mismatchFile)) {
            const wb = XLSX.readFile(mismatchFile);

            if (wb.SheetNames.includes(mismatchSheetName)) {
                delete wb.Sheets[mismatchSheetName];
                wb.SheetNames = wb.SheetNames.filter(
                (name) => name !== mismatchSheetName
                );
                if (wb.SheetNames.length > 0) {
    XLSX.writeFile(wb, mismatchFile);
    console.log(`Old mismatch sheet "${mismatchSheetName}" deleted.`);
  } else {
    console.warn(`Workbook became empty after deleting "${mismatchSheetName}", skipping save.`);
  }
            }
            }

            await test.step("Verified, All test data matched", async () => {
            console.log("All data matched, old mismatch sheet removed if present!");
            });

        }
    }

    NSFTransactionsToExcelTest=async(path)=>{
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        await this.scrollToHeaderNSFTransactions()
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        if  (await this.tbl_NSFTransactions.isVisible()){
            await expect.soft(this.tbl_NSFTransactions).toBeVisible()
        }else{
            await this.headerNSFTransactions.click()
            await expect.soft(this.tbl_NSFTransactions).toBeVisible()
        }

        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        const rows = await this.tbl_NSFTransactionsAll.all();
        const tableData = [];
        console.log(rows);

        for (let i = 1; i < rows.length; i++) {
            const cells = await rows[i].locator("td").allInnerTexts();
            tableData.push({
            BillingSysDate: cells[0].trim(),
            Amount: cells[1].trim(),
            BillingSysDesc: cells[2].trim(),
            });
        }

        writeDataToExcel(path, "NSFTransactions", tableData);
    }
    compareAndExportNSFTransactionsMismatch=async(testEnvData,mismatchSheetName,mismatchFile)=>{
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        await this.scrollToHeaderNSFTransactions()
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        if  (await this.tbl_NSFTransactions.isVisible()){

            await expect(this.tbl_NSFTransactions).toBeVisible()
        }else{
            await this.headerNSFTransactions.click()
            await expect(this.tbl_NSFTransactions).toBeVisible()
        }

        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        const rows = await this.tbl_NSFTransactionsAll.all();
        const prodData = [];

        for (let i = 1; i < rows.length; i++) {
            const cells = await rows[i].locator("td").allInnerTexts();
            prodData.push({
            BillingSysDate: cells[0].trim(),
            Amount: cells[1].trim(),
            BillingSysDesc: cells[2].trim(),

            });
        }

        const testData = readDataFromExcel(testEnvData, "NSFTransactions");
        
        let mismatches = [];

    for (let i = 0; i < Math.max(prodData.length, testData.length); i++) {
        const prodRow = prodData[i];
        const testRow = testData[i];

        if (!prodRow || !testRow) {
            mismatches.push({
                Row: i + 1,
                Column: "N/A",
                Test_BillingSysDate: testRow?.BillingSysDate || "Missing",
                Test_Amount: testRow?.Amount || "Missing",
                Test_BillingSysDesc: testRow?.BillingSysDesc || "Missing",
                Prod_BillingSysDate: prodRow?.BillingSysDate || "Missing",
                Prod_Amount: prodRow?.Amount || "Missing",
                Prod_BillingSysDesc: prodRow?.BillingSysDesc || "Missing",
            });
            continue;
        }

        Object.keys(prodRow).forEach((key) => {
            if (prodRow[key] !== testRow[key]) {
                mismatches.push({
                    Row: i + 1,
                    Column: key,
                    Test_BillingSysDate: testRow?.BillingSysDate || "N/A",
                    Test_Amount: testRow?.Amount || "N/A",
                    Test_BillingSysDesc: testRow?.BillingSysDesc || "N/A",
                    Prod_BillingSysDate: prodRow?.BillingSysDate || "N/A",
                    Prod_Amount: prodRow?.Amount || "N/A",
                    Prod_BillingSysDesc: prodRow?.BillingSysDesc || "N/A",
                });
            }
        });
    }

        //const mismatchFile = path.resolve("./excel/Mismatches/NSFTransactions_Mismatch.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found. Exporting detailed report to Excel...");

            writeDataToExcel(mismatchFile, mismatchSheetName, mismatches);
            expect.soft(mismatches.length, "mismatch_NSFTransactions mismatch count").toBe(0);
        } else {
                    console.log("All data matched!");

            if (fs.existsSync(mismatchFile)) {
            const wb = XLSX.readFile(mismatchFile);

            if (wb.SheetNames.includes(mismatchSheetName)) {
                delete wb.Sheets[mismatchSheetName];
                wb.SheetNames = wb.SheetNames.filter(
                (name) => name !== mismatchSheetName
                );
                if (wb.SheetNames.length > 0) {
    XLSX.writeFile(wb, mismatchFile);
    console.log(`Old mismatch sheet "${mismatchSheetName}" deleted.`);
  } else {
    console.warn(`Workbook became empty after deleting "${mismatchSheetName}", skipping save.`);
  }
            }
            }

            await test.step("Verified, All test data matched", async () => {
            console.log("All data matched, old mismatch sheet removed if present!");
            });

        }
    }

    scrollToReversalTransaction=async()=>{
        await excuteSteps(this.test,this.headerReversalTransactions ,"scroll",`Scrolling into view`);
    }
    
    ReversalTransactionsTransactionsToExcelTest=async(path)=>{
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        if  (await this.tbl_ReversalTransactions.isVisible()){
            await expect.soft(this.tbl_ReversalTransactions).toBeVisible()
        }else{
            await this.headerReversalTransactions.click()
            await expect.soft(this.tbl_ReversalTransactions).toBeVisible()
        }

        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        const rows = await this.tbl_ReversalTransactionsAll.all();
        const tableData = [];
        console.log(rows);

        for (let i = 1; i < rows.length; i++) {
            const cells = await rows[i].locator("td").allInnerTexts();
            tableData.push({
            BillingSysDate: cells[0].trim(),
            Amount: cells[1].trim(),
            BillingSysDesc: cells[2].trim(),
            
            });
        }

        writeDataToExcel(path, "ReversalTransactions", tableData);
    }
    compareAndExportReversalTransactionMismatch=async(testEnvData,mismatchSheetName,mismatchFile)=>{
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        if  (await this.tbl_ReversalTransactions.isVisible()){
            await expect.soft(this.tbl_ReversalTransactions).toBeVisible()
        }else{
            await this.headerReversalTransactions.click()
            await expect.soft(this.tbl_ReversalTransactions).toBeVisible()
        }

        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        
        const rows = await this.tbl_ReversalTransactionsAll.all();
        const prodData = [];
        console.log("reversal transactions count",rows.length);
        for (let i = 1; i < rows.length; i++) {
            const cells = await rows[i].locator("td").allInnerTexts();
            prodData.push({
            BillingSysDate: cells[0].trim(),
            Amount: cells[1].trim(),
            BillingSysDesc: cells[2].trim(),
            });
        }
        console.log("Reversal transactions prod data",prodData);
        const testData = readDataFromExcel(testEnvData, "ReversalTransactions");
        let mismatches = [];

    for (let i = 0; i < Math.max(prodData.length, testData.length); i++) {
        const prodRow = prodData[i];
        const testRow = testData[i];

        if (!prodRow || !testRow) {
            mismatches.push({
                Row: i + 1,
                Column: "N/A",
                Test_BillingSysDate: testRow?.BillingSysDate || "Missing",
                Test_Amount: testRow?.Amount || "Missing",
                Test_BillingSysDesc: testRow?.BillingSysDesc || "Missing",
                Prod_BillingSysDate: prodRow?.BillingSysDate || "Missing",
                Prod_Amount: prodRow?.Amount || "Missing",
                Prod_BillingSysDesc: prodRow?.BillingSysDesc || "Missing",
            });
            continue;
        }

        Object.keys(prodRow).forEach((key) => {
            if (prodRow[key] !== testRow[key]) {
                mismatches.push({
                    Row: i + 1,
                    Column: key,
                    Test_BillingSysDate: testRow?.BillingSysDate || "N/A",
                    Test_Amount: testRow?.Amount || "N/A",
                    Test_BillingSysDesc: testRow?.BillingSysDesc || "N/A",
                    Prod_BillingSysDate: prodRow?.BillingSysDate || "N/A",
                    Prod_Amount: prodRow?.Amount || "N/A",
                    Prod_BillingSysDesc: prodRow?.BillingSysDesc || "N/A",
                });
            }
        });
    }
        //const mismatchFile = path.resolve("./excel/Mismatches/Mismatch_ReversalTransactions.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found. Exporting detailed report to Excel...");
            writeDataToExcel(mismatchFile,mismatchSheetName, mismatches);
            expect.soft(mismatches.length, "Mismatch_ReversalTransactions mismatch count").toBe(0);
        } else {
                    console.log("All data matched!");

            // ✅ If the file exists, check if the mismatch sheet exists
            if (fs.existsSync(mismatchFile)) {
            const wb = XLSX.readFile(mismatchFile);

            if (wb.SheetNames.includes(mismatchSheetName)) {
                delete wb.Sheets[mismatchSheetName];
                wb.SheetNames = wb.SheetNames.filter(
                (name) => name !== mismatchSheetName
                );
                if (wb.SheetNames.length > 0) {
    XLSX.writeFile(wb, mismatchFile);
    console.log(`Old mismatch sheet "${mismatchSheetName}" deleted.`);
  } else {
    console.warn(`Workbook became empty after deleting "${mismatchSheetName}", skipping save.`);
  }
            }
            }

            await test.step("Verified, All test data matched", async () => {
            console.log("All data matched, old mismatch sheet removed if present!");
            });
        }
    }

    scrollToInternalBankTransfers=async()=>{
        await excuteSteps(this.test,this.headerInternalBankTransfers ,"scroll",`Scrolling into view`);
    }

    InternalBankTransfersTransactionsToExcelTest=async(path)=>{
        await scrollToElement(this.headerInternalBankTransfers);
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        if  (await this.tbl_InternalBankTransfers.isVisible()){
            
            await expect.soft(this.tbl_InternalBankTransfers).toBeVisible()
        }else{
            await this.headerInternalBankTransfers.click()
            await expect.soft(this.tbl_InternalBankTransfers).toBeVisible()
        }

        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        const rows = await this.tbl_InternalBankTransfersAll.all();
        const tableData = [];
        console.log("No of Rows in Internal Bank Transfers are:", rows.length)
        console.log(rows);
        for (let i = 1; i < rows.length-1; i++) {
            const cells = await rows[i].locator("td").allInnerTexts();
            tableData.push({
            Date: cells[0].trim(),
            Amount: cells[1].trim(),
            TypeCode: cells[2].trim(),
            Description: cells[3].trim(),
            Status: cells[4].trim(),
            });        
        }
        const rowsCount= rows.length-1
        const cellsData= await rows[rowsCount].locator("td").allInnerTexts();
        console.log("Internal Transfer LastRow Data", rowsCount, cellsData);
        let data1= cellsData[0].trim()
        let data2 = cellsData[1].trim()
        console.log("The data to write excel", data1, ":", data2);
        tableData.push({
            Date: data1,
            Amount: data2,
            TypeCode: "",
            Description: "",
            Status: "",
            
        })

        writeDataToExcel(path, "InternalBankTransfers", tableData);
    }
    compareAndExportInternalBankTransfersMismatch=async(testEnvData,mismatchSheetName,mismatchFile)=>{
        await scrollToElement(this.headerInternalBankTransfers);
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        if  (await this.tbl_InternalBankTransfers.isVisible()){
            
            await expect.soft(this.tbl_InternalBankTransfers).toBeVisible()
        }else{
            await this.headerInternalBankTransfers.click()
            await expect.soft(this.tbl_InternalBankTransfers).toBeVisible()
        }

        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        const rows = await this.tbl_InternalBankTransfersAll.all();
        const prodData = [];
        for (let i = 1; i < rows.length-1; i++) {
            const cells = await rows[i].locator("td").allInnerTexts();
            prodData.push({
            Date: cells[0].trim(),
            Amount: cells[1].trim(),
            TypeCode: cells[2].trim(),
            Description: cells[3].trim(),
            Status: cells[4].trim(),
            });
        }
        const  rowsCount= rows.length-1;
        console.log(rowsCount);
        const cellsData= await rows[rowsCount].locator("td").allInnerTexts();
        console.log("Internal Transfer prod LastRow Data", rowsCount, cellsData);
        let data1= cellsData[0].trim();
        let data2 = cellsData[1].trim();
        console.log("The data to write excel Internal Transfers", data1, ":", data2);
        prodData.push({
            Date: data1,
            Amount: data2,
            TypeCode: "",
            Description: "",
            Status: "",
            
        })
        const testData = readDataFromExcel(testEnvData, "InternalBankTransfers");
        let mismatches = [];

    for (let i = 0; i < Math.max(prodData.length, testData.length); i++) {
        const prodRow = prodData[i];
        const testRow = testData[i];

        if (!prodRow || !testRow) {
            mismatches.push({
                Row: i + 1,
                Column: "N/A",
                Test_Date: testRow?.Date || "Missing",
                Test_Amount: testRow?.Amount || "Missing",
                Test_TypeCode: testRow?.TypeCode || "Missing",
                Test_Description: testRow?.Description || "Missing",
                Test_Status: testRow?.Status || "Missing",
                Prod_Date: prodRow?.Date || "Missing",
                Prod_Amount: prodRow?.Amount || "Missing",
                Prod_TypeCode: prodRow?.TypeCode || "Missing",
                Prod_Description: prodRow?.Description || "Missing",
                Prod_Status: prodRow?.Status || "Missing",
            });
            continue;
        }
        Object.keys(prodRow).forEach((key) => {
            if (prodRow[key] !== testRow[key]) {
                mismatches.push({
                    Row: i + 1,
                    Column: key,
                    Test_Date: testRow?.Date || "Missing",
                    Test_Amount: testRow?.Amount || "Missing",
                    Test_TypeCode: testRow?.TypeCode || "Missing",
                    Test_Description: testRow?.Description || "Missing",
                    Test_Status: testRow?.Status || "Missing",
                    Prod_Date: prodRow?.Date || "Missing",
                    Prod_Amount: prodRow?.Amount || "Missing",
                    Prod_BillingSysDesc: prodRow?.TypeCode || "Missing",
                    Prod_BillingSysDesc: prodRow?.Description || "Missing",
                    Prod_Status: prodRow?.Status || "Missing",
                });
            }
        });
    }

        //const mismatchFile = path.resolve("./excel/Mismatches/InternalBankTransfers_Mismatch.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found. Exporting detailed report to Excel...");

            writeDataToExcel(mismatchFile, mismatchSheetName, mismatches);
            expect.soft(mismatches.length, "Mismatch_InternalBankTransfers mismatch count").toBe(0);
        } else {
                    console.log("All data matched!");

            if (fs.existsSync(mismatchFile)) {
            const wb = XLSX.readFile(mismatchFile);

            if (wb.SheetNames.includes(mismatchSheetName)) {
                delete wb.Sheets[mismatchSheetName];
                wb.SheetNames = wb.SheetNames.filter(
                (name) => name !== mismatchSheetName
                );
                if (wb.SheetNames.length > 0) {
    XLSX.writeFile(wb, mismatchFile);
    console.log(`Old mismatch sheet "${mismatchSheetName}" deleted.`);
  } else {
    console.warn(`Workbook became empty after deleting "${mismatchSheetName}", skipping save.`);
  }
            }
            }

            await test.step("Verified, All test data matched", async () => {
            console.log("All data matched, old mismatch sheet removed if present!");
            });

        }
    }
    

    scrollToHeaderNDCSweeps=async()=>{
        await excuteSteps(this.test,this.headerNDCSweeps ,"scroll",`Scrolling into view`);
    }
    NDCSweepsTransactionsToExcelTest=async(path)=>{
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        await this.scrollToHeaderNDCSweeps();
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        if  (await this.tbl_NDCSweeps.isVisible()){
            await expect.soft(this.tbl_NDCSweeps).toBeVisible()
        }else{
            await this.headerNDCSweeps.click()
            await expect.soft(this.tbl_NDCSweeps).toBeVisible()
        }
        const rows = await this.tbl_NDCSweepsAll.all();
        const tableData = [];
        console.log("NDC Sweeps Table RowCount: ", rows.length)
        console.log(rows);

        for (let i = 1; i < rows.length-1; i++) {
            const cells = await rows[i].locator("td").allInnerTexts();
            tableData.push({
            Date: cells[0].trim(),
            Amount: cells[1].trim(),
            TypeCode: cells[2].trim(),
            Description: cells[3].trim(),
            BankAccount: cells[4].trim(),
            });
        }

        const rowsCount= rows.length-1
        const cellsData= await rows[rowsCount].locator("td").allInnerTexts();
        let data1= cellsData[0].trim()
        let data2 = cellsData[1].trim()
        console.log("The data to write excel", data1, ":", data2);
        tableData.push({
            Date: data1,
            Amount: data2,
            TypeCode: "",
            Description: "",
            Status: "",
            
        })
        writeDataToExcel(path, "NDCSweeps", tableData);
    }
    compareAndExportNDCSweepsMismatch=async(testEnvData,mismatchSheetName,mismatchFile)=>{

        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        await this.scrollToHeaderNDCSweeps();
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        if  (await this.tbl_NDCSweeps.isVisible()){
            await expect.soft(this.tbl_NDCSweeps).toBeVisible()
        }else{
            await this.headerNDCSweeps.click()
            await expect.soft(this.tbl_NDCSweeps).toBeVisible()
        }
        const rows = await this.tbl_NDCSweepsAll.all();
        const prodData = [];

        for (let i = 1; i < rows.length-1; i++) {
            const cells = await rows[i].locator("td").allInnerTexts();
            prodData.push({
            Date: cells[0].trim(),
            Amount: cells[1].trim(),
            TypeCode: cells[2].trim(),
            Description: cells[3].trim(),
            BankAccount: cells[4].trim(),
            });
        }
        const rowsCount= rows.length
        const cellsData= await rows[rowsCount-1].locator("td").allInnerTexts();
        let data1= cellsData[0].trim()
        let data2 = cellsData[1].trim()
        console.log("The data to write NDC Sweeps excel", data1, ":", data2);
        prodData.push({
            Date: data1,
            Amount: data2,
            TypeCode: "",
            Description: "",
            BankAccount: "",
            
        })
        

        const testData = readDataFromExcel(testEnvData, "NDCSweeps");

        let mismatches = [];

    for (let i = 0; i < Math.max(prodData.length, testData.length); i++) {
        const prodRow = prodData[i];
        const testRow = testData[i];

        if (!prodRow || !testRow) {
            mismatches.push({
                Row: i + 1,
                Column: "N/A",
                Test_Date: testRow?.Date || "Missing",
                Test_Amount: testRow?.Amount || "Missing",
                Test_TypeCode: testRow?.TypeCode || "Missing",
                Test_Description: testRow?.Description || "Missing",
                Test_Status: testRow?.Status || "Missing",
                Prod_Date: prodRow?.Date || "Missing",
                Prod_Amount: prodRow?.Amount || "Missing",
                Prod_BillingSysDesc: prodRow?.TypeCode || "Missing",
                Prod_BillingSysDesc: prodRow?.Description || "Missing",
                Prod_Status: prodRow?.Status || "Missing",
            });
            continue;
        }
        Object.keys(prodRow).forEach((key) => {
            if (prodRow[key] !== testRow[key]) {
                mismatches.push({
                    Row: i + 1,
                    Column: key,
                    Test_Date: testRow?.Date || "Missing",
                    Test_Amount: testRow?.Amount || "Missing",
                    Test_TypeCode: testRow?.TypeCode || "Missing",
                    Test_Description: testRow?.Description || "Missing",
                    Test_Status: testRow?.Status || "Missing",
                    Prod_Date: prodRow?.Date || "Missing",
                    Prod_Amount: prodRow?.Amount || "Missing",
                    Prod_BillingSysDesc: prodRow?.TypeCode || "Missing",
                    Prod_BillingSysDesc: prodRow?.Description || "Missing",
                    Prod_Status: prodRow?.Status || "Missing",
                });
            }
        });
    }        
        //const mismatchFile = path.resolve("./excel/Mismatches/Mismatch_NDCSweeps.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found. Exporting detailed report to Excel...");

            writeDataToExcel(mismatchFile,mismatchSheetName, mismatches);
            expect.soft(mismatches.length, "Mismatch_NDCSweeps mismatch count").toBe(0);
            //throw new Error("Data mismatch detected! Check MismatchReport.xlsx for details.");
        } else {
                    console.log("All data matched!");

            // ✅ If the file exists, check if the mismatch sheet exists
            if (fs.existsSync(mismatchFile)) {
            const wb = XLSX.readFile(mismatchFile);

            if (wb.SheetNames.includes(mismatchSheetName)) {
                delete wb.Sheets[mismatchSheetName];
                wb.SheetNames = wb.SheetNames.filter(
                (name) => name !== mismatchSheetName
                );
                if (wb.SheetNames.length > 0) {
    XLSX.writeFile(wb, mismatchFile);
    console.log(`Old mismatch sheet "${mismatchSheetName}" deleted.`);
  } else {
    console.warn(`Workbook became empty after deleting "${mismatchSheetName}", skipping save.`);
  }
            }
            }

            await test.step("Verified, All test data matched", async () => {
            console.log("All data matched, old mismatch sheet removed if present!");
            });

        }
    }

    scrollToHeaderCashReconciliationReport=async()=>{
        await excuteSteps(this.test,this.lbl_cashReconciliationReport ,"scroll",`Scrolling into view`);
    }

    scrollToHeaderCashReconciliationDeductionReport=async()=>{
        await excuteSteps(this.test,this.header_ReconciliationReport_Deductions ,"scroll",`Scrolling into view`);
    }
    

    scrollToHeaderCashReconciliationBankDepositsByAccount=async()=>{
        await excuteSteps(this.test,this.header_cashReconciliationReport_BankDepositsByAccount ,"scroll",`Scrolling into view`);
    }
    scrollToHeaderCashReconciliationBillingSystemAdjustments=async()=>{
        await excuteSteps(this.test,this.header_cashReconciliationReport_BillingSystemAdjustment ,"scroll",`Scrolling into view`);
    }
    
    scrollToHeaderCashReconciliationReconciliation=async()=>{
        await excuteSteps(this.test,this.header_cashReconciliationReport_Reconciliation ,"scroll",`Scrolling into view`);
    }
    
    
    async Recon_Reconciliation() {
        const tableData = [];
        const data = [];
        await this.scrollToHeaderCashReconciliationReconciliation()
        const rows_cashReconciliationReport_Reconciliation = await this.tbl_cashReconciliationReport_Reconciliation.all();
        console.log("rows length is for Cash Reconciliation Report --> Reconciliation:", rows_cashReconciliationReport_Reconciliation.length);
        for (let i = 0; i < rows_cashReconciliationReport_Reconciliation.length; i++) {
            let dataText = "";
            dataText = await rows_cashReconciliationReport_Reconciliation[i].textContent();
            tableData.push(dataText);
        }
        data.push(tableData);
        writeDataToExcel("./excel/Cash_Reconciliation_Report.xlsx", "Reconciliation", tableData);
    }


    
    async Recon_unMatchedSystemTrans() {
        const tableData = [];
        await this.scrollToHeaderCashReconciliationBillingSystemAdjustments();
        const rows_unMatchedBillingSystemTransactions = await this.tbl_cashReconciliationReport_unMatchedBillingSystemTransactions.all();
        const rows_unMatchedBillingSystemTransactions_allTexts = await this.tbl_cashReconciliationReport_unMatchedBillingSystemTransactions.allTextContents();
        console.log(rows_unMatchedBillingSystemTransactions_allTexts);
        console.log("rows length is for Cash Reconciliation Report --> UnMatched Billing Transactions:", rows_unMatchedBillingSystemTransactions.length);
        let rowsCount = await rows_unMatchedBillingSystemTransactions.length;
        for (let i = 0; i < rowsCount; i++) {
            // tableData.push(rows_unMatchedBillingSystemTransactions[i].innerText.trim());
            let dataText = "";
            dataText = await rows_unMatchedBillingSystemTransactions[i].innerText();
            // dataText= await rows_unMatchedBillingSystemTransactions[i].innerText 
            console.log("The text in UnMatched Billing Transactions is", dataText);
            tableData.push(dataText);
        }
        console.log("Table Data", tableData);
        // writeDataToExcel("./excel/CashReconRep_TableData_TestEnv.xlsx", "unMatchedSystemTrans", tableData);
        writeDataToExcel("./excel/Cash_Reconciliation_Report.xlsx", "UnMatchedSystemTrans", tableData);
    }

    async Recon_BillingSystemAdjustments() {
        const tableData = [];
        const rows_BillingSystemAdjustments = await this.tbl_cashReconciliationReport_BillingSystemAdjustments.all();
        const rows_BillingSystemAdjustments_allTexts = await this.tbl_cashReconciliationReport_BillingSystemAdjustments.allTextContents();
        console.log(rows_BillingSystemAdjustments_allTexts);
        console.log("rows length is for Cash Reconciliation Report --> BillingSystemAdjustments:", rows_BillingSystemAdjustments.length);
        for (let i = 0; i < rows_BillingSystemAdjustments.length; i++) {
            let dataText = "";
            dataText = await rows_BillingSystemAdjustments[i].textContent();
            tableData.push(dataText);
        }
        console.log("Table Data", tableData);
        // writeDataToExcel("./excel/CashReconRep_tableData_TestEnv.xlsx", "Adjustments", tableData);
        writeDataToExcel("./excel/Cash_Reconciliation_Report.xlsx", "BillingSystemAdjustments", tableData);
    }

    async Recon_Deductions() {
        const tableData = [];
        const data = [];

        // this.scrollToHeaderCashReconciliationDeductionReport();
        const rows_Deductions = await this.tbl_cashReconciliationReport_Deductions.all();
        const rows_Deductions_allText = await this.tbl_cashReconciliationReport_Deductions.allTextContents();
        console.log("rows length is for Cash Reconciliation Report --> Deductions:", rows_Deductions.length);
        for (let i = 0; i < rows_Deductions.length; i++) {
            // tableData.push(rows_Deductions[i].innerText.trim());
            let dataText = "";
            dataText = await rows_Deductions[i].textContent();
            tableData.push(dataText);
        }
        data.push(tableData);
        // writeDataToExcel("./excel/CashReconRep_Deductions_tableData_TestEnv.xlsx", "Deductions", tableData);
        writeDataToExcel("./excel/Cash_Reconciliation_Report.xlsx", "Deductions", data);
    }

    async ReconReport_TotalBilling() {
        const tableData=[]
        const data=[]
        const rows = await this.tbl_cashReconciliationReport_TotalBillingTransactions.all();
        console.log("rows length is for Cash Reconciliation Report --> Total Billing System Transactions:", rows.length);
        console.log(rows.length);
        let rowsCount = await rows.length;
        for (let i = 0; i < rowsCount; i++) {
            let dataText = "";
            dataText = await rows[i].textContent();
            console.log("dataText1", dataText);
            tableData.push(dataText);
        }
        data.push(tableData);
        console.log("Table Data", data);
        writeDataToExcel("./excel/Cash_Reconciliation_Report.xlsx", "TotalBilling", data);
    }

    async BankDepositByAccount() {
        const tableData = [];
        const data = [];
        this.scrollToHeaderCashReconciliationBankDepositsByAccount();
        const rows_BankDeposits = await this.tbl_cashReconciliationReport_BankDepositsByAccount.all();
        const rows_BankDeposits_allText = await this.tbl_cashReconciliationReport_BankDepositsByAccount.allTextContents();

        console.log(rows_BankDeposits_allText);
        // tableData.push(rows_BankDeposits_allText)
        // this.excuteSteps(this.test, this.header_cashReconciliationReport_BankDepositsByAccount, "scroll", "scroll to view")
        console.log("rows length is for Cash Reconciliation Report --> Bank Deposits By Account:", rows_BankDeposits.length);
        for (let i = 0; i < rows_BankDeposits.length; i++) {
            // tableData.push(rows_BankDeposits[i].innerText.trim());
            let dataText = "";
            dataText = await rows_BankDeposits[i].textContent();
            // dataText=await rows_BankDeposits[i].innerText()
            tableData.push(dataText);
        }
        data.push(tableData);
        console.log("Table Data", tableData);
        console.log("Table Data", data);
        
        writeDataToExcel("./excel/Cash_Reconciliation_Report.xlsx", "BankDepositByAccount", data);
    }

    compareAndExportReconReconciliationMismatch=async()=>{
        const prodData = [];
        const prodDataWrite =[];
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        const rows = await this.tbl_cashReconciliationReport_Reconciliation.all();         
        console.log("rows length is for Cash Reconciliation Report --> Reconciliation:", rows_cashReconciliationReport_Reconciliation.length);
        for (let i = 0; i < rows_cashReconciliationReport_Reconciliation.length; i++) {
            // tableData.push(rows_cashReconciliationReport_Reconciliation[i].innerText.trim());
            let dataText = "";
            dataText = await rows_cashReconciliationReport_Reconciliation[i].textContent();
            // dataText= rows_cashReconciliationReport_Reconciliation[i].innerText
            prodData.push(dataText);
        }
        prodDataWrite.push(prodData)

        const testData = readDataFromExcel("./excel/TestEnvData.xlsx", "Reconciliation_TestEnv");

        let mismatches = [];

        for (let i = 0; i < Math.max(prodDataWrite.length, data.length); i++) {
            const prodRow = prodDataWrite[i];
            const testRow = testData[i];

            if (!prodRow || !testRow) {
            mismatches.push({
                Row: i + 1,
                Column: "N/A",
                TestEnvValue: testRow ? JSON.stringify(testRow) : "Missing Row",
                ProdEnvValue: prodRow ? JSON.stringify(prodRow) : "Missing Row",
                TestEnvRow: testRow ? JSON.stringify(testRow) : "N/A",
                ProdEnvRow: prodRow ? JSON.stringify(prodRow) : "N/A",
            });
            continue;
            }

            Object.keys(prodRow).forEach((key) => {
            if (prodRow[key] !== testRow[key]) {
                mismatches.push({
                Row: i + 1,
                Column: key,
                TestEnvValue: testRow[key],
                ProdEnvValue: prodRow[key],
                TestEnvRow: JSON.stringify(testRow),
                ProdEnvRow: JSON.stringify(prodRow),
                });
            }
            });
        }
        const mismatchFile = path.resolve("./excel/Mismatch_CashReconRep_Data.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found. Exporting detailed report to Excel...");

            writeDataToExcel(mismatchFile, "CashReconRep_Data", mismatches);
            expect.soft(mismatches.length, "CashReconRep_Data mismatch count").toBe(0);
            //throw new Error(`Data mismatch detected! Check ${mismatchFile} for details.`);
        } else {
            if (fs.existsSync(mismatchFile)) {
                fs.unlinkSync(mismatchFile);
                console.log("Old Mismatch_CashReconRep_Data file deleted ✅");
                }

            await test.step("Verified, All rows in CashReconRep_Data matched successfully!",async()=>{
            console.log("All rows matched successfully!");
            })
        }
    }

    compareAndExportReconBillingSystemAdjustmentsMismatch=async()=>{
        const prodData = [];
        const prodDataWrite =[];
        await this.page.waitForTimeout(parseInt(process.env.mediumWait));
        const rows = await this.tbl_cashReconciliationReport_BillingSystemAdjustments.all();         
        for (let i = 0; i < rows.length; i++) {
            let dataText = "";
            dataText = await rows[i].textContent();
            // dataText= rows_cashReconciliationReport_Reconciliation[i].innerText
            prodData.push(dataText);
        }
        prodDataWrite.push(prodData)

        const testData = readDataFromExcel("./excel/TestEnvData.xlsx", "CashReconRep_Adjusts_TestEnv");

        let mismatches = [];

        for (let i = 0; i < Math.max(prodDataWrite.length, data.length); i++) {
            const prodRow = prodDataWrite[i];
            const testRow = testData[i];

            if (!prodRow || !testRow) {
            mismatches.push({
                Row: i + 1,
                Column: "N/A",
                TestEnvValue: testRow ? JSON.stringify(testRow) : "Missing Row",
                ProdEnvValue: prodRow ? JSON.stringify(prodRow) : "Missing Row",
                TestEnvRow: testRow ? JSON.stringify(testRow) : "N/A",
                ProdEnvRow: prodRow ? JSON.stringify(prodRow) : "N/A",
            });
            continue;
            }

            Object.keys(prodRow).forEach((key) => {
            if (prodRow[key] !== testRow[key]) {
                mismatches.push({
                Row: i + 1,
                Column: key,
                TestEnvValue: testRow[key],
                ProdEnvValue: prodRow[key],
                TestEnvRow: JSON.stringify(testRow),
                ProdEnvRow: JSON.stringify(prodRow),
                });
            }
            });
        }
        const mismatchFile = path.resolve("./excel/Mismatch_Adjustments.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found. Exporting detailed report to Excel...");

            writeDataToExcel(mismatchFile, "Mismatch_Recon_Adjustments", mismatches);
            expect.soft(mismatches.length, "Mismatch_Recon_Adjustments mismatch count").toBe(0);
            //throw new Error(`Data mismatch detected! Check ${mismatchFile} for details.`);
        } else {
            if (fs.existsSync(mismatchFile)) {
                fs.unlinkSync(mismatchFile);
                console.log("Old Mismatch_Adjustments file deleted ✅");
                }

            await test.step("Verified, All rows in CashReconRep_Adjustments matched successfully!",async()=>{
            console.log("All rows matched successfully!");
            })
        }
    }

    compareAndExportReconUnMatchedSystemTransMismatch=async()=>{
        const prodData = [];
        const prodDataWrite =[];
        const rows = await this.tbl_cashReconciliationReport_unMatchedBillingSystemTransactions.all();
        for (let i = 0; i < rows.length; i++) {
            let dataText = "";
            dataText = await rows[i].textContent();
            prodData.push(dataText);
        }
        prodDataWrite.push(prodData)

        const testData = readDataFromExcel("./excel/TestEnvData.xlsx", "unMatchedSystemTrans_TestEnv");

        let mismatches = [];

        for (let i = 0; i < Math.max(prodDataWrite.length, data.length); i++) {
            const prodRow = prodDataWrite[i];
            const testRow = testData[i];

            if (!prodRow || !testRow) {
            mismatches.push({
                Row: i + 1,
                Column: "N/A",
                TestEnvValue: testRow ? JSON.stringify(testRow) : "Missing Row",
                ProdEnvValue: prodRow ? JSON.stringify(prodRow) : "Missing Row",
                TestEnvRow: testRow ? JSON.stringify(testRow) : "N/A",
                ProdEnvRow: prodRow ? JSON.stringify(prodRow) : "N/A",
            });
            continue;
            }

            Object.keys(prodRow).forEach((key) => {
            if (prodRow[key] !== testRow[key]) {
                mismatches.push({
                Row: i + 1,
                Column: key,
                TestEnvValue: testRow[key],
                ProdEnvValue: prodRow[key],
                TestEnvRow: JSON.stringify(testRow),
                ProdEnvRow: JSON.stringify(prodRow),
                });
            }
            });
        }
        const mismatchFile = path.resolve("./excel/Mismatch_CashReconRep_unMatchedSystemTrans.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found. Exporting detailed report to Excel...");

            writeDataToExcel(mismatchFile, "Mismatch_CashReconRep_unMatchedSystemTrans", mismatches);
            expect.soft(mismatches.length, "Mismatch_CashReconRep_unMatchedSystemTrans mismatch count").toBe(0);
            //throw new Error(`Data mismatch detected! Check ${mismatchFile} for details.`);
        } else {
            if (fs.existsSync(mismatchFile)) {
                fs.unlinkSync(mismatchFile);
                console.log("Old Mismatch_CashReconRep_unMatchedSystemTrans file deleted ✅");
                }

            await test.step("Verified, All rows in CashReconRep_unMatchedSystemTrans matched successfully!",async()=>{
            console.log("All rows matched successfully!");
            })
        }
    }

    compareAndExportReconTotalBillingTransMismatch=async()=>{
        const prodData = [];
        const prodDataWrite =[];
        const rows = await this.tbl_cashReconciliationReport_TotalBillingTransactions.all();
        for (let i = 0; i < rows.length; i++) {
            let dataText = "";
            dataText = await rows[i].textContent();
            prodData.push(dataText);
        }
        prodDataWrite.push(prodData)

        const testData = readDataFromExcel("./excel/TestEnvData.xlsx", "CashReconRep_TotalBilling_TestEnv");

        let mismatches = [];

        for (let i = 0; i < Math.max(prodDataWrite.length, data.length); i++) {
            const prodRow = prodDataWrite[i];
            const testRow = testData[i];

            if (!prodRow || !testRow) {
            mismatches.push({
                Row: i + 1,
                Column: "N/A",
                TestEnvValue: testRow ? JSON.stringify(testRow) : "Missing Row",
                ProdEnvValue: prodRow ? JSON.stringify(prodRow) : "Missing Row",
                TestEnvRow: testRow ? JSON.stringify(testRow) : "N/A",
                ProdEnvRow: prodRow ? JSON.stringify(prodRow) : "N/A",
            });
            continue;
            }

            Object.keys(prodRow).forEach((key) => {
            if (prodRow[key] !== testRow[key]) {
                mismatches.push({
                Row: i + 1,
                Column: key,
                TestEnvValue: testRow[key],
                ProdEnvValue: prodRow[key],
                TestEnvRow: JSON.stringify(testRow),
                ProdEnvRow: JSON.stringify(prodRow),
                });
            }
            });
        }
        const mismatchFile = path.resolve("./excel/Mismatch_CashReconRep_TotalBilling.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found. Exporting detailed report to Excel...");

            writeDataToExcel(mismatchFile, "Mismatch_CashReconRep_TotalBilling", mismatches);
            expect.soft(mismatches.length, "Mismatch_CashReconRep_TotalBilling mismatch count").toBe(0);
            //throw new Error(`Data mismatch detected! Check ${mismatchFile} for details.`);
        } else {
            if (fs.existsSync(mismatchFile)) {
                fs.unlinkSync(mismatchFile);
                console.log("Old Mismatch_CashReconRep_TotalBilling file deleted ✅");
                }

            await test.step("Verified, All rows in CashReconRep_TotalBilling matched successfully!",async()=>{
            console.log("All rows matched successfully!");
            })
        }
    }

    compareAndExportBankDepositByAccountMismatch=async()=>{
        const prodData = [];
        const prodDataWrite =[];
        const rows = await this.tbl_cashReconciliationReport_BankDepositsByAccount.all();
        for (let i = 0; i < rows.length; i++) {
            let dataText = "";
            dataText = await rows[i].textContent();
            prodData.push(dataText);
        }
        prodDataWrite.push(prodData)

        const testData = readDataFromExcel("./excel/TestEnvData.xlsx", "BankDepositByAccount_TestEnv");

        let mismatches = [];

        for (let i = 0; i < Math.max(prodDataWrite.length, data.length); i++) {
            const prodRow = prodDataWrite[i];
            const testRow = testData[i];

            if (!prodRow || !testRow) {
            mismatches.push({
                Row: i + 1,
                Column: "N/A",
                TestEnvValue: testRow ? JSON.stringify(testRow) : "Missing Row",
                ProdEnvValue: prodRow ? JSON.stringify(prodRow) : "Missing Row",
                TestEnvRow: testRow ? JSON.stringify(testRow) : "N/A",
                ProdEnvRow: prodRow ? JSON.stringify(prodRow) : "N/A",
            });
            continue;
            }

            Object.keys(prodRow).forEach((key) => {
            if (prodRow[key] !== testRow[key]) {
                mismatches.push({
                Row: i + 1,
                Column: key,
                TestEnvValue: testRow[key],
                ProdEnvValue: prodRow[key],
                TestEnvRow: JSON.stringify(testRow),
                ProdEnvRow: JSON.stringify(prodRow),
                });
            }
            });
        }
        const mismatchFile = path.resolve("./excel/Mismatch_CashReconRep_BankDepositByAccount.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found. Exporting detailed report to Excel...");

            writeDataToExcel(mismatchFile, "Mismatch_CashReconRep_BankDepositByAccount", mismatches);
            expect.soft(mismatches.length, "Mismatch_CashReconRep_BankDepositByAccount mismatch count").toBe(0);
           // throw new Error(`Data mismatch detected! Check ${mismatchFile} for details.`);
        } else {
            if (fs.existsSync(mismatchFile)) {
                fs.unlinkSync(mismatchFile);
                console.log("Old CashReconRep_BankDepositByAccount file deleted ✅");
                }

            await test.step("Verified, All rows in CashReconRep_BankDepositByAccount matched successfully!",async()=>{
            console.log("All rows matched successfully!");
            })
        }
    }


    compareAndExportDeductionsMismatch=async()=>{
        const prodData = [];
        const prodDataWrite =[];
        const rows = await this.tbl_cashReconciliationReport_Deductions.all();
        for (let i = 0; i < rows.length; i++) {
            let dataText = "";
            dataText = await rows[i].textContent();
            prodData.push(dataText);
        }
        prodDataWrite.push(prodData)

        const testData = readDataFromExcel("./excel/TestEnvData.xlsx", "ReconRep_Deductions_TestEnv");

        let mismatches = [];

        for (let i = 0; i < Math.max(prodDataWrite.length, data.length); i++) {
            const prodRow = prodDataWrite[i];
            const testRow = testData[i];

            if (!prodRow || !testRow) {
            mismatches.push({
                Row: i + 1,
                Column: "N/A",
                TestEnvValue: testRow ? JSON.stringify(testRow) : "Missing Row",
                ProdEnvValue: prodRow ? JSON.stringify(prodRow) : "Missing Row",
                TestEnvRow: testRow ? JSON.stringify(testRow) : "N/A",
                ProdEnvRow: prodRow ? JSON.stringify(prodRow) : "N/A",
            });
            continue;
            }

            Object.keys(prodRow).forEach((key) => {
            if (prodRow[key] !== testRow[key]) {
                mismatches.push({
                Row: i + 1,
                Column: key,
                TestEnvValue: testRow[key],
                ProdEnvValue: prodRow[key],
                TestEnvRow: JSON.stringify(testRow),
                ProdEnvRow: JSON.stringify(prodRow),
                });
            }
            });
        }
        const mismatchFile = path.resolve("./excel/Mismatch_CashReconRep_BankDepositByAccount.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found. Exporting detailed report to Excel...");

            writeDataToExcel(mismatchFile, "Mismatch_CashReconRep_Deductions", mismatches);
            expect.soft(mismatches.length, "Mismatch_CashReconRep_Deductions mismatch count").toBe(0);
            //throw new Error(`Data mismatch detected! Check ${mismatchFile} for details.`);
        } else {
            if (fs.existsSync(mismatchFile)) {
                fs.unlinkSync(mismatchFile);
                console.log("Old Mismatch_CashReconRep_BankDepositByAccount file deleted ✅");
                }

            await test.step("Verified, All rows in CashReconRep_BankDepositByAccount matched successfully!",async()=>{
            console.log("All rows matched successfully!");
            })
        }
    }


     CashReconciliationReportToExcelTest=async(path)=>{
        await scrollToElement(this.lbl_cashReconciliationReport);
        const tableData=[];
        const totalBillingrows = await this.tbl_cashReconciliationReport_TotalBillingTransactions.all();
        let rowsCount = await totalBillingrows.length;
        
        for (let i = 0; i < rowsCount; i +=2) {
            let dataText1 = await totalBillingrows[i].textContent();
            let dataText2 = await totalBillingrows[i+1].textContent();
            
            tableData.push({
                "A":dataText1,
                "B":dataText2,
                "C":"",
            });
        }
        const rows_BankDeposits = await this.tbl_cashReconciliationReport_BankDepositsByAccount.all();
        
        for (let i = 0; i < rows_BankDeposits.length; i+=2) {
            let dataText1 = await rows_BankDeposits[i].textContent();
            let dataText2 = await rows_BankDeposits[i+1].textContent();
            tableData.push({
                "A":dataText1,
                "B":dataText2,
                "C":"",
            });
        }

        const rows_Deductions = await this.tbl_cashReconciliationReport_Deductions.all();
        
        for (let i = 0; i < rows_Deductions.length; i+=2) {
            let dataText1 = await rows_Deductions[i].textContent();
            let dataText2 = await rows_Deductions[i+1].textContent();
    
            tableData.push({
                "A":dataText1,
                "B":dataText2,
                "C":"",
            });
        }

        const rows_BillingSystemAdjustments = await this.tbl_cashReconciliationReport_BillingSystemAdjustments.all();
        
        for (let i = 0; i < rows_BillingSystemAdjustments.length; i+=2) {
            let dataText1 = await rows_BillingSystemAdjustments[i].textContent();
            let dataText2 = await rows_BillingSystemAdjustments[i+1].textContent();
            tableData.push({
                "A":dataText1,
                "B":dataText2,
                "C":"",
            });
        }

        const rows_unMatchedBillingSystemTransactions = await this.tbl_cashReconciliationReport_unMatchedBillingSystemTransactions.all();
        let rows_unmatchedBilling = await rows_unMatchedBillingSystemTransactions.length;
        
        for (let i = 0; i < rows_unmatchedBilling-3; i+=3) {
            let dataText1 = await rows_unMatchedBillingSystemTransactions[i].textContent();
            let dataText2 = await rows_unMatchedBillingSystemTransactions[i+1].textContent();
            let dataText3= await rows_unMatchedBillingSystemTransactions[i+2].textContent();

            tableData.push({
                "A":dataText1,
                "B":dataText2,
                "C":dataText3,
            });
        }
        let rowCount=rows_unmatchedBilling-2
        let dataText1 = await rows_unMatchedBillingSystemTransactions[rowCount].textContent();
        let dataText2 = await rows_unMatchedBillingSystemTransactions[rowCount+1].textContent();

        tableData.push({
                "A":dataText1,
                "B":dataText2,
                "C":"",
        });

        const rows_cashReconciliationReport_Reconciliation = await this.tbl_cashReconciliationReport_Reconciliation.all();
        for (let i = 0; i < rows_cashReconciliationReport_Reconciliation.length; i+=2) {
            let dataText1 = await rows_cashReconciliationReport_Reconciliation[i].textContent();
            let dataText2 = await rows_cashReconciliationReport_Reconciliation[i+1].textContent();

            tableData.push({
                "A":dataText1,
                "B":dataText2,
                "C":"",
            });
        }
        
        console.log(tableData);
        writeDataToExcel(path, "TotalBilling", tableData);
     
    }



    writingBankBillingSystemAnalysisTest=async(path)=>{
        const totalTransactions = await this.lbl_BSA_TotalTransactions.innerText();
        const matched = await this.lbl_BSA_MatchedTransactions.innerText();
        const nSFs= await this.lbl_BSA_NSF.innerText();
        const reversals= await this.lbl_BSA_Reversals.innerText();
        const unmatched = await this.lbl_BSA_unMatched.innerText();
        const matchRate = await this.lbl_BSA_matchRate.innerText();

        console.log({ totalTransactions, matched, nSFs, reversals, unmatched, matchRate });

        const testData = [{
        "TotalTransactions": totalTransactions,
        "Matched": matched,
        "NSFs": nSFs,
        "Reversals": reversals,
        "Unmatched": unmatched,
        "MatchRate": matchRate
        }];

    writeDataToExcel(path, "BankSystemAnalysis", testData);

    }



writingBankBillingSystemAnalysisTotalMatchesFoundTest=async(path)=>{
        const totalTransactions = await this.lbl_TotalMatchesFound_total.innerText();
        const exactMatches = await this.lbl_TotalMatchesFound_ExactMatches.innerText();
        const oneToMany= await this.lbl_TotalMatchesFound_oneToMany.innerText();
        const manyToMany= await this.lbl_TotalMatchesFound_manyToMany.innerText();

        const testData = [{
        "TotalTransactions": totalTransactions,
        "ExactMatches": exactMatches,
        "OneToMany": oneToMany,
        "ManyToMany": manyToMany
        }];

    writeDataToExcel(path, "TotalMatchesFound", testData);
    }


writingBankBillingSystemAnalysisReconciliationStatusTest=async(path)=>{
        const totalDeposits = await this.lbl_ReconStatus_totalDeposits.innerText();
        const totalCashReceipts = await this.lbl_ReconStatus_totalCashReceipts.innerText();
        const differences= await this.lbl_ReconStatus_differences.innerText();
        console.log({ totalDeposits, totalCashReceipts, differences });

        const testData = [{
        "TotalDeposits": totalDeposits,
        "TotalCashReceipts": totalCashReceipts,
        "Differences": differences
        }];

    writeDataToExcel(path, "ReconciliationStatus", testData);
    }


    gettingBandFeedDetailsBIllingSystemAnalysisAndCompareProd =async(testEnvData,mismatchSheetName,mismatchFile)=>{
            const totalTransactions = await this.lbl_BSA_TotalTransactions.innerText();
            const matched = await this.lbl_BSA_MatchedTransactions.innerText();
            const nSFs= await this.lbl_BSA_NSF.innerText();
            const reversals= await this.lbl_BSA_Reversals.innerText();
            const unmatched = await this.lbl_BSA_unMatched.innerText();
            const matchRate = await this.lbl_BSA_matchRate.innerText();
            const prodData = [{
            "TotalTransactions": totalTransactions,
            "Matched": matched,
            "NSFs": nSFs,
            "Reversals": reversals,
            "Unmatched": unmatched,
            "MatchRate": matchRate
            }];

            const testData = readDataFromExcel(testEnvData, "BankSystemAnalysis");
             console.log("prodData",prodData);
             console.log("testdata",testData);
            const mismatches = getMismatches(prodData, testData);
            //const mismatchFile = path.resolve("./excel/Mismatches/MisMatch_BankSystemAnalysis.xlsx");
            if (mismatches.length > 0) {
                console.log("Mismatches found:", mismatches);
                writeDataToExcel(mismatchFile, mismatchSheetName, mismatches);
                expect.soft(mismatches.length, "Mismatches found in data").toBe(0);
            } else {
                console.log("All data matched!");
            if (fs.existsSync(mismatchFile)) {
            const wb = XLSX.readFile(mismatchFile);

            if (wb.SheetNames.includes(mismatchSheetName)) {
                delete wb.Sheets[mismatchSheetName];
                wb.SheetNames = wb.SheetNames.filter(
                (name) => name !== mismatchSheetName
                );
                if (wb.SheetNames.length > 0) {
    XLSX.writeFile(wb, mismatchFile);
    console.log(`Old mismatch sheet "${mismatchSheetName}" deleted.`);
  } else {
    console.warn(`Workbook became empty after deleting "${mismatchSheetName}", skipping save.`);
  }
                }
                }
                await test.step("Verified, All test data matched", async () => {
                console.log("All data matched, old mismatch sheet removed if present!");
                });

                
            }
        }
    
gettingBandFeedDetailsTotalMatchesFoundAndCompareProd =async(testEnvData,mismatchSheetName,mismatchFile)=>{
        const totalTransactions = await this.lbl_TotalMatchesFound_total.innerText();
        const exactMatches = await this.lbl_TotalMatchesFound_ExactMatches.innerText();
        const oneToMany= await this.lbl_TotalMatchesFound_oneToMany.innerText();
        const manyToMany= await this.lbl_TotalMatchesFound_manyToMany.innerText();

        const prodData  = [{
            "TotalTransactions": totalTransactions,
            "ExactMatches": exactMatches,
            "OneToMany": oneToMany,
            "ManyToMany": manyToMany
            }];
    
        const testData = readDataFromExcel(testEnvData, "TotalMatchesFound");
        const mismatches = getMismatches(prodData, testData);
        //const mismatchFile = path.resolve("./excel/Mismatches/Mismatch_TotalMatchesFound_BSA.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found:", mismatches);
            writeDataToExcel(mismatchFile, "Mismatch_TotalMatchesFound_BSA", mismatches);
            expect.soft(mismatches.length, "Mismatches found in data").toBe(0);
        } else {
                    console.log("All data matched!");
            if (fs.existsSync(mismatchFile)) {
            const wb = XLSX.readFile(mismatchFile);

            if (wb.SheetNames.includes(mismatchSheetName)) {
                delete wb.Sheets[mismatchSheetName];
                wb.SheetNames = wb.SheetNames.filter(
                (name) => name !== mismatchSheetName
                );
                if (wb.SheetNames.length > 0) {
    XLSX.writeFile(wb, mismatchFile);
    console.log(`Old mismatch sheet "${mismatchSheetName}" deleted.`);
  } else {
    console.warn(`Workbook became empty after deleting "${mismatchSheetName}", skipping save.`);
  }
            }
            }
            await test.step("Verified, All test data matched", async () => {
            console.log("All data matched, old mismatch sheet removed if present!");
            });   
        }
    }
    
    gettingBandFeedDetailsReconciliationStatusAndCompareProd =async(testEnvData,mismatchSheetName,mismatchFile)=>{

        const totalDeposits = await this.lbl_ReconStatus_totalDeposits.innerText();
        const totalCashReceipts = await this.lbl_ReconStatus_totalCashReceipts.innerText();
        const differences= await this.lbl_ReconStatus_differences.innerText();
        const prodData = [{
            "TotalDeposits": totalDeposits,
            "TotalCashReceipts": totalCashReceipts,
            "Differences": differences
            }];

        const testData = readDataFromExcel(testEnvData, "ReconciliationStatus");
        const mismatches = getMismatches(prodData, testData);
        //const mismatchFile = path.resolve("./excel/Mismatches/Mismatch_ReconciliationStatus.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found:", {mismatches});
            writeDataToExcel(mismatchFile,mismatchSheetName, mismatches);
            expect.soft(mismatches.length, "Mismatches found in data").toBe(0);
        } else {
                        console.log("All data matched!");
                if (fs.existsSync(mismatchFile)) {
                const wb = XLSX.readFile(mismatchFile);

                if (wb.SheetNames.includes(mismatchSheetName)) {
                    delete wb.Sheets[mismatchSheetName];
                    wb.SheetNames = wb.SheetNames.filter(
                    (name) => name !== mismatchSheetName
                    );
                    XLSX.writeFile(wb, mismatchFile);
                    console.log(`Old mismatch sheet "${mismatchSheetName}" deleted.`);
                }
                }

                await test.step("Verified, All test data matched", async () => {
                console.log("All data matched, old mismatch sheet removed if present!");
                });

        }
    }


    compareAndExportCashReconciliationReportMismatch=async(testEnvData,mismatchSheetName,mismatchFile)=>{

        await scrollToElement(this.lbl_cashReconciliationReport);
        const prodData=[];
        const totalBillingrows = await this.tbl_cashReconciliationReport_TotalBillingTransactions.all();
        let rowsCount = await totalBillingrows.length;
        
        for (let i = 0; i < rowsCount; i +=2) {
            let dataText1 = await totalBillingrows[i].textContent();
            let dataText2 = await totalBillingrows[i+1].textContent();
            
            prodData.push({
                "A":dataText1,
                "B":dataText2,
                "C":"",
            });
        }
        const rows_BankDeposits = await this.tbl_cashReconciliationReport_BankDepositsByAccount.all();
        
        for (let i = 0; i < rows_BankDeposits.length; i+=2) {
            let dataText1 = await rows_BankDeposits[i].textContent();
            let dataText2 = await rows_BankDeposits[i+1].textContent();
            prodData.push({
                "A":dataText1,
                "B":dataText2,
                "C":"",
            });
        }

        const rows_Deductions = await this.tbl_cashReconciliationReport_Deductions.all();
        
        for (let i = 0; i < rows_Deductions.length; i+=2) {
            let dataText1 = await rows_Deductions[i].textContent();
            let dataText2 = await rows_Deductions[i+1].textContent();
    
            prodData.push({
                "A":dataText1,
                "B":dataText2,
                "C":"",
            });
        }

        const rows_BillingSystemAdjustments = await this.tbl_cashReconciliationReport_BillingSystemAdjustments.all();
        
        for (let i = 0; i < rows_BillingSystemAdjustments.length; i+=2) {
            let dataText1 = await rows_BillingSystemAdjustments[i].textContent();
            let dataText2 = await rows_BillingSystemAdjustments[i+1].textContent();
            prodData.push({
                "A":dataText1,
                "B":dataText2,
                "C":"",
            });
        }

        const rows_unMatchedBillingSystemTransactions = await this.tbl_cashReconciliationReport_unMatchedBillingSystemTransactions.all();
        let rows_unmatchedBilling = await rows_unMatchedBillingSystemTransactions.length;
        
        for (let i = 0; i < rows_unmatchedBilling-3; i+=3) {
            let dataText1 = await rows_unMatchedBillingSystemTransactions[i].textContent();
            let dataText2 = await rows_unMatchedBillingSystemTransactions[i+1].textContent();
            let dataText3= await rows_unMatchedBillingSystemTransactions[i+2].textContent();

            prodData.push({
                "A":dataText1,
                "B":dataText2,
                "C":dataText3,
            });
        }
        let rowCount=rows_unmatchedBilling-2
        let dataText1 = await rows_unMatchedBillingSystemTransactions[rowCount].textContent();
        let dataText2 = await rows_unMatchedBillingSystemTransactions[rowCount+1].textContent();

        prodData.push({
                "A":dataText1,
                "B":dataText2,
                "C":"",
        });

        const rows_cashReconciliationReport_Reconciliation = await this.tbl_cashReconciliationReport_Reconciliation.all();
        for (let i = 0; i < rows_cashReconciliationReport_Reconciliation.length; i+=2) {
            let dataText1 = await rows_cashReconciliationReport_Reconciliation[i].textContent();
            let dataText2 = await rows_cashReconciliationReport_Reconciliation[i+1].textContent();

            prodData.push({
                "A":dataText1,
                "B":dataText2,
                "C":"",
            });
        }
        console.log("Last card prod data",prodData);
        const testData = readDataFromExcel(testEnvData, "TotalBilling");
        console.log("------Last card test data-------- ",testData);
        let mismatches = [];

        for (let i = 0; i < Math.max(prodData.length, testData.length); i++) {
        const prodRow = prodData[i];
        const testRow = testData[i];

        if (!prodRow || !testRow) {
            mismatches.push({
                Row: i + 1,
                Column: "N/A",
                Test_A: testRow?.A || "Missing",
                Test_B: testRow?.B || "Missing",
                Test_C: testRow?.C || "Missing",
                Prod_A: testRow?.A || "Missing",
                Prod_B: testRow?.B || "Missing",
                Prod_C: testRow?.C || "Missing",
            });
            continue;
        }
        Object.keys(prodRow).forEach((key) => {
            if (prodRow[key] !== testRow[key]) {
                mismatches.push({
                    Row: i + 1,
                    Column: key,
                    Test_A: testRow?.A || "Missing",
                    Test_B: testRow?.B || "Missing",
                    Test_C: testRow?.C || "Missing",
                    Prod_A: prodRow?.A || "Missing",
                    Prod_B: prodRow?.B || "Missing",
                    Prod_C: prodRow?.C || "Missing",
                });
            }
        });
    }        
        //const mismatchFile = path.resolve("./excel/Mismatches/Cash_Reconciliation.xlsx");
        if (mismatches.length > 0) {
            console.log("Mismatches found. Exporting detailed report to Excel...");
            writeDataToExcel(mismatchFile, mismatchSheetName, mismatches);
            expect.soft(mismatches.length, "Cash_Reconciliation mismatch count").toBe(0);
        } else {
                    console.log("All data matched!");

            // ✅ If the file exists, check if the mismatch sheet exists
            if (fs.existsSync(mismatchFile)) {
            const wb = XLSX.readFile(mismatchFile);

            if (wb.SheetNames.includes(mismatchSheetName)) {
                delete wb.Sheets[mismatchSheetName];
                wb.SheetNames = wb.SheetNames.filter(
                (name) => name !== mismatchSheetName
                );
                if (wb.SheetNames.length > 0) {
    XLSX.writeFile(wb, mismatchFile);
    console.log(`Old mismatch sheet "${mismatchSheetName}" deleted.`);
  } else {
    console.warn(`Workbook became empty after deleting "${mismatchSheetName}", skipping save.`);
  }
            }
            }

            await test.step("Verified, All test data matched", async () => {
            console.log("All data matched, old mismatch sheet removed if present!");
            });

        }
        

        }



}



