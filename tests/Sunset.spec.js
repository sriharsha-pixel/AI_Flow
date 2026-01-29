const { test, expect } = require("@playwright/test");
const sections = require("../pageObjects/UI_Pages/pageIndex");
const path = require("path");
require("dotenv").config();

test("Getting summary and Matched Transaction details from Stage Env",async({page})=>{
    const loginPage = new sections.LoginPage(test, page);
      await loginPage.launchingApplication([process.env.base_url_env]);
      await loginPage.loginToLovable([process.env.lovableUsername],[process.env.lovablePassword]);
      await loginPage.loginWithValidCredentials(
        [process.env.user_name],
        [process.env.password]
      );
    const bankStatementfilefolder = path.resolve(__dirname, "../testData/Sunset/BankStatementFiles");
    const rfmsFileFolder=path.resolve(__dirname, "../testData/Sunset/RFMSFiles");
    const billingSystemFileFolder=path.resolve(__dirname, "../testData/Sunset/BillingSystemFiles");
    const testEnvData= path.resolve("./output/Sunset/TestEnvData.xlsx");
    const cashReconReport= path.resolve("./output/Sunset/Cash_Reconciliation_Report.xlsx");
    const testEnvDatareload= path.resolve("./output/Sunset/TestEnvData_Database.xlsx");
    const cashReconReportreload= path.resolve("./output/Sunset/Cash_Reconciliation_Report_Database.xlsx");
    const cashPosting=new sections.CashPosting(test,page);
    await cashPosting.uploadingFilesInTest(bankStatementfilefolder,rfmsFileFolder,billingSystemFileFolder);
    await cashPosting.writingBankFeedSummaryTest(testEnvData);
    await cashPosting.writingBankBillingSystemAnalysisTest(testEnvData);
    await cashPosting.writingBankBillingSystemAnalysisTotalMatchesFoundTest(testEnvData);
    await cashPosting.writingBankBillingSystemAnalysisReconciliationStatusTest(testEnvData);
    if (await page.locator("//h3[text()='Matched Transactions']").isVisible()) {
      await cashPosting.matchedTransactionsToExcelTest(testEnvData);
    }
    if(await page.locator("//h3[contains(text(),'Cash Receipts in Billing System not Found in Bank')]").isVisible()){
      await cashPosting.CashReceiptsInBillingSystemNotFoundInBankExcelTest(testEnvData);    
    }
    if(await page.locator("//h3[text()='Deposits in Bank/RFMS Not Found in Billing System']").isVisible()){
    await cashPosting.DepositsInBankRFMSNotFoundTransactionsToExcelTest(testEnvData);
    }
    if(await page.locator("//h3[contains(text(),'NSF Transactions')]").isVisible()){
    await cashPosting.NSFTransactionsToExcelTest(testEnvData);
    }
    if(await page.locator("//h3[text()='Reversal Transactions']").isVisible()){
    await cashPosting.ReversalTransactionsTransactionsToExcelTest(testEnvData);
    }
    if(await page.locator("//h3[text()='Internal Bank Transfers']").isVisible()){
    await cashPosting.InternalBankTransfersTransactionsToExcelTest(testEnvData);
    }
    if(await page.locator("//h3[text()='NDC Sweeps']").isVisible()){
    await cashPosting.NDCSweepsTransactionsToExcelTest(testEnvData);
    }
    if(await page.locator("//h3[text()='Cash Reconciliation Report']").isVisible()){
    await cashPosting.CashReconciliationReportToExcelTest(cashReconReport);
    }
    await page.reload();
    await page.waitForTimeout(parseInt(process.env.mediumWait));
    await cashPosting.clickOnFirstCard();
    await cashPosting.totalTransaction.waitFor({state:'visible'});
    await cashPosting.writingBankFeedSummaryTest(testEnvDatareload);
    await cashPosting.writingBankBillingSystemAnalysisTest(testEnvDatareload);
    await cashPosting.writingBankBillingSystemAnalysisTotalMatchesFoundTest(testEnvDatareload);
    await cashPosting.writingBankBillingSystemAnalysisReconciliationStatusTest(testEnvDatareload);
    if (await page.locator("//h3[text()='Matched Transactions']").isVisible()) {
      await cashPosting.matchedTransactionsToExcelTest(testEnvDatareload);
    }
    if(await page.locator("//h3[contains(text(),'Cash Receipts in Billing System not Found in Bank')]").isVisible()){
      await cashPosting.CashReceiptsInBillingSystemNotFoundInBankExcelTest(testEnvDatareload);    
    }
    if(await page.locator("//h3[text()='Deposits in Bank/RFMS Not Found in Billing System']").isVisible()){
    await cashPosting.DepositsInBankRFMSNotFoundTransactionsToExcelTest(testEnvDatareload);
    }
    if(await page.locator("//h3[contains(text(),'NSF Transactions')]").isVisible()){
    await cashPosting.NSFTransactionsToExcelTest(testEnvDatareload);
    }
    
    if(await page.locator("//h3[text()='Reversal Transactions']").isVisible()){
    await cashPosting.ReversalTransactionsTransactionsToExcelTest(testEnvDatareload);
    }
    if(await page.locator("//h3[text()='Internal Bank Transfers']").isVisible()){
    await cashPosting.InternalBankTransfersTransactionsToExcelTest(testEnvDatareload);
    }
    if(await page.locator("//h3[text()='NDC Sweeps']").isVisible()){
    await cashPosting.NDCSweepsTransactionsToExcelTest(testEnvDatareload);
    }
    if(await page.locator("//h3[text()='Cash Reconciliation Report']").isVisible()){
    await cashPosting.CashReconciliationReportToExcelTest(cashReconReportreload);
    }

  });

test("Comparing Stage Env Data with Prod Env Data", async({page})=>{
      const loginPage = new sections.LoginPage(test, page);
      await loginPage.launchingApplication([process.env.base_url_prod]);
      await loginPage.loginWithValidCredentials(
        [process.env.user_name],
        [process.env.password]
      );

    const cashPosting=new sections.CashPosting(test,page);

    const bankStatementfilefolder = path.resolve(__dirname, "../testData/Sunset/BankStatementFiles");
    const rfmsFileFolder=path.resolve(__dirname, "../testData/Sunset/RFMSFiles");
    const billingSystemFileFolder=path.resolve(__dirname, "../testData/Sunset/BillingSystemFiles");
    const testEnvData= path.resolve("./output/Sunset/TestEnvData.xlsx");
    const cashReconReport= path.resolve("./output/Sunset/Cash_Reconciliation_Report.xlsx");
    const testEnvDatareload= path.resolve("./output/Sunset/TestEnvData_Database.xlsx");
    const cashReconReportreload= path.resolve("./output/Sunset/Cash_Reconciliation_Report_Database.xlsx");
    const mismatchFile = path.resolve("./output/Mismatches/Sunset_MismatchResult.xlsx");
    const mismatchFileDatabase = path.resolve("./output/Mismatches/Sunset_MismatchResult_Database.xlsx");
    
    await cashPosting.uploadingFilesInProd(bankStatementfilefolder,rfmsFileFolder,billingSystemFileFolder);
    await cashPosting.gettingBandFeedDetailsAndCompareProd(testEnvData,"Bank Feed Analysis",mismatchFile);
    await cashPosting.gettingBandFeedDetailsBIllingSystemAnalysisAndCompareProd(testEnvData,"BankSystemAnalysis",mismatchFile);
    await cashPosting.gettingBandFeedDetailsTotalMatchesFoundAndCompareProd(testEnvData,"TotalMatchesFound_BSA",mismatchFile);
    await cashPosting.gettingBandFeedDetailsReconciliationStatusAndCompareProd(testEnvData,"ReconciliationStatus",mismatchFile);
    if (await page.locator("//h3[text()='Matched Transactions']").isVisible()) {
    await cashPosting.compareAndExportMismatch(testEnvData,"Matched Transactions",mismatchFile);
    }
    if(await page.locator("//h3[contains(text(),'Cash Receipts in Billing System not Found in Bank')]").isVisible()){
    await cashPosting.compareAndExportCashReceiptsInBillingSystem(testEnvData,"UnmatchedBilling",mismatchFile);
    }
    if(await page.locator("//h3[text()='Deposits in Bank/RFMS Not Found in Billing System']").isVisible()){
    await cashPosting.compareAndExportDepositsInBankRFMSNotFoundTransactions(testEnvData,"UnmatchedBankDeposits",mismatchFile);
    }
    if(await page.locator("//h3[contains(text(),'NSF Transactions')]").isVisible()){
    await cashPosting.compareAndExportNSFTransactionsMismatch(testEnvData,"NSFTransactions",mismatchFile);
    }
    if(await page.locator("//h3[text()='Reversal Transactions']").isVisible()){
    await cashPosting.compareAndExportReversalTransactionMismatch(testEnvData,"ReversalTransactions",mismatchFile);
    }
    if(await page.locator("//h3[text()='Internal Bank Transfers']").isVisible()){
    await cashPosting.compareAndExportInternalBankTransfersMismatch(testEnvData,"",mismatchFile);
    await cashPosting.compareAndExportInternalBankTransfersMismatch(testEnvData,"InternalBankTransfers",mismatchFile);
    }
    if(await page.locator("//h3[text()='NDC Sweeps']").isVisible()){
    await cashPosting.compareAndExportNDCSweepsMismatch(testEnvData,"NDCSweeps",mismatchFile);
    }
    if(await page.locator("//h3[text()='Cash Reconciliation Report']").isVisible()){
    await cashPosting.compareAndExportCashReconciliationReportMismatch(cashReconReport,"Cash Reconciliation",mismatchFile);
    }

    await page.reload();

    await page.waitForTimeout(parseInt(process.env.mediumWait));
    await cashPosting.clickOnFirstCard();
    await page.waitForTimeout(parseInt(process.env.mediumWait));
    await cashPosting.totalTransaction.waitFor({state:'visible'});
    await cashPosting.gettingBandFeedDetailsAndCompareProd(testEnvDatareload,"Bank Feed Analysis",mismatchFileDatabase);
    await cashPosting.gettingBandFeedDetailsBIllingSystemAnalysisAndCompareProd(testEnvDatareload,"BankSystemAnalysis",mismatchFileDatabase);
    await cashPosting.gettingBandFeedDetailsTotalMatchesFoundAndCompareProd(testEnvDatareload,"TotalMatchesFound_BSA",mismatchFileDatabase);
    await cashPosting.gettingBandFeedDetailsReconciliationStatusAndCompareProd(testEnvDatareload,"ReconciliationStatus",mismatchFileDatabase);
    if (await page.locator("//h3[text()='Matched Transactions']").isVisible()) {
    await cashPosting.compareAndExportMismatch(testEnvDatareload,"Matched Transactions",mismatchFileDatabase);
    }
    if(await page.locator("//h3[contains(text(),'Cash Receipts in Billing System not Found in Bank')]").isVisible()){
    await cashPosting.compareAndExportCashReceiptsInBillingSystem(testEnvDatareload,"UnmatchedBilling",mismatchFileDatabase);
    }
    if(await page.locator("//h3[text()='Deposits in Bank/RFMS Not Found in Billing System']").isVisible()){
    await cashPosting.compareAndExportDepositsInBankRFMSNotFoundTransactions(testEnvDatareload,"UnmatchedBankDeposits",mismatchFileDatabase);
    }
    if(await page.locator("//h3[contains(text(),'NSF Transactions')]").isVisible()){
    await cashPosting.compareAndExportNSFTransactionsMismatch(testEnvDatareload,"NSFTransactions",mismatchFileDatabase);
    }
    if(await page.locator("//h3[text()='Reversal Transactions']").isVisible()){
    await cashPosting.compareAndExportReversalTransactionMismatch(testEnvDatareload,"ReversalTransactions",mismatchFileDatabase);
    }
    if(await page.locator("//h3[text()='Internal Bank Transfers']").isVisible()){
    await cashPosting.compareAndExportInternalBankTransfersMismatch(testEnvDatareload,"InternalBankTransfers",mismatchFileDatabase);
    }
    if(await page.locator("//h3[text()='NDC Sweeps']").isVisible()){
    await cashPosting.compareAndExportNDCSweepsMismatch(testEnvDatareload,"NDCSweeps",mismatchFileDatabase);
    }
    if(await page.locator("//h3[text()='Cash Reconciliation Report']").isVisible()){
    await cashPosting.compareAndExportCashReconciliationReportMismatch(cashReconReportreload,"Cash Reconciliation",mismatchFileDatabase);
    }
    
});