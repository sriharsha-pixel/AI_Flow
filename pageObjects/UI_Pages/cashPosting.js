const { excuteSteps } = require("../../utilities/actions");
const { test, expect } = require("@playwright/test");
const { extractTextFromImage } = require("../../utilities/extractTextFromImage");
const path = require("path");
const { writeDataToExcel, readDataFromExcel } = require("../../utilities/readExcel");
const { getMismatches } = require("../../utilities/getMismatches");
require("dotenv").config();
const XLSX = require("xlsx");
const fs = require("fs");
const { scrollToElement } = require("../../utilities/scrollInView");
const { getFilesFromFolder } = require("../../utilities/getFilesFromFolder");
const { da } = require("@faker-js/faker");
const pdfParse = require("pdf-parse");

exports.CashPosting = class CashPosting {
  constructor(test, page) {
    this.test = test;
    this.page = page;
    this.lbl_NoTransactions = page.locator("//div[contains(text(),'No transactions to display')]");
    this.firstcard = page.locator("(//div[@data-radix-scroll-area-viewport]/div/div/div)[1]");
    this.cashPostinggetStartedBtn = page.locator("//h3[text()='Cash Posting Reconciliation']/following::button[1]");
    this.bankStatementFileUploadBtn = page.locator("//input[@id='bank-file-input']");
    this.rfmsFileUploadBtn = page.locator("//input[@id='rfms-file-input']");
    this.billingSystemfileUploadBtn = page.locator("//input[@id='billing-file-input']");
    this.runReconsillationBtn = page.locator("//button[text()='Run Reconciliation']");
    this.summaryCard = page.locator("(//div[@data-component-file='ReconSummary.tsx'])[1]");
    this.summaryCardProd = page.locator("(//div[@id='results-section']/div/div)[2]");
    this.matchedTransactionsHeader = page.locator("//h3[text()='Matched Transactions']");
    this.totalTransaction = page.locator("((//div[@id='results-section']/div/div)[2]/div)[1]//span[text()='Total Transactions:']/following-sibling::div");
    this.matched = page.locator("((//div[@id='results-section']/div/div)[2]/div)[1]//span[text()='Matched:']/following-sibling::div");
    this.transfers = page.locator("((//div[@id='results-section']/div/div)[2]/div)[1]//span[text()='Transfers:']/following-sibling::div");
    this.ndc = page.locator("((//div[@id='results-section']/div/div)[2]/div)[1]//span[text()='NDC Sweeps:']/following-sibling::div");
    this.unmatched = page.locator("((//div[@id='results-section']/div/div)[2]/div)[1]//span[text()='Unmatched:']/following-sibling::div/div");
    this.matchrate = page.locator("((//div[@id='results-section']/div/div)[2]/div)[1]//span[text()='Match Rate:']/following-sibling::div/span");
    this.processing = page.locator("//span[contains(text(),'Processing Reconciliation')]");
    // NewChanges
    this.header_cashReceiptsInBIllingSystem = page.locator("//h3[contains(text(),'Cash Receipts in Billing System not Found in Bank')]")
    this.tbl_cashReceiptsInBIllingSystem = page.locator("//h3[contains(text(),'Cash Receipts in Billing System not Found in Bank')]//following::table[1]")
    this.tbl_cashReceiptsInBIllingSystemAll = page.locator("//h3[contains(text(),'Cash Receipts in Billing System not Found in Bank')]//following::table[1]//tr")

    this.header_cashReceiptsInBIllingSystemDropDown = page.locator("//h3[contains(text(),'Cash Receipts in Billing System not Found in Bank')]//following::*[@data-lov-name='ChevronDown']")

    this.headerNSFTransactions = page.locator("//h3[contains(text(),'NSF Transactions')]")
    this.header_NSFTransactionsDropDownIon = page.locator("//h3[contains(text(),'NSF Transactions')]//following::*[@data-lov-name='ChevronDown']")
    this.tbl_NSFTransactions = page.locator("//h3[contains(text(),'NSF Transactions')]//following::table[1]")
    this.tbl_NSFTransactionsAll = page.locator("//h3[contains(text(),'NSF Transactions')]//following::table[1]//tr")

    this.headerDepositsInBankRFMS = page.locator("//h3[text()='Deposits in Bank/RFMS Not Found in Billing System']")
    this.headerDepositsInBankRFMSDropDownIcon = page.locator("//h3[text()='Deposits in Bank/RFMS Not Found in Billing System']//following::*[@data-lov-name='ChevronDown']")
    this.tblDepositsInBankRFMS = page.locator("//h3[text()='Deposits in Bank/RFMS Not Found in Billing System']//following::table[1]")
    this.tblDepositsInBankRFMSAll = page.locator("//h3[text()='Deposits in Bank/RFMS Not Found in Billing System']//following::table[1]//tr")

    this.headerReversalTransactions = page.locator("//h3[text()='Reversal Transactions']")
    this.headerReversalTransactionsDropDownIcon = page.locator("//h3[text()='Reversal Transactions']//following::*[@data-lov-name='ChevronDown']")
    this.tbl_ReversalTransactions = page.locator("//h3[text()='Reversal Transactions']//following::table[1]")
    this.tbl_ReversalTransactionsAll = page.locator("//h3[text()='Reversal Transactions']//following::table[1]//tr")


    this.headerInternalBankTransfers = page.locator("//h3[text()='Internal Bank Transfers']")
    this.headerInternalBankTransfersDropDownIcon = page.locator("//h3[text()='Internal Bank Transfers']//following::*[@data-lov-name='ChevronDown']")
    this.tbl_InternalBankTransfersAll = page.locator("//h3[text()='Internal Bank Transfers']//following::table[1]//tr")


    this.headerNDCSweeps = page.locator("//h3[text()='NDC Sweeps']")
    this.headerNDCSweepsDropDownIcon = page.locator("//h3[text()='NDC Sweeps']//following::*[@data-lov-name='ChevronDown']")
    this.tbl_NDCSweeps = page.locator("//h3[text()='NDC Sweeps']//following::table[1]")
    this.tbl_NDCSweepsAll = page.locator("//h3[text()='NDC Sweeps']//following::table[1]//tr");

    this.lbl_BSA_TotalTransactions = page.locator("//h3[text()='Billing System Analysis']//following::span[text()='Total Transactions:']//following::div[1]")
    this.lbl_BSA_MatchedTransactions = page.locator("//h3[text()='Billing System Analysis']//following::span[text()='Matched:']//following::div[1]//div[1]")
    this.lbl_BSA_NSF = page.locator("//h3[text()='Billing System Analysis']//following::span[text()='NSFs:']//following::div[1]")
    this.lbl_BSA_Reversals = page.locator("//h3[text()='Billing System Analysis']//following::span[text()='Reversals:']//following::div[1]")
    this.lbl_BSA_unMatched = page.locator("//h3[text()='Billing System Analysis']//following::span[text()='Unmatched:']//following::div[1]//div[1]")
    this.lbl_BSA_matchRate = page.locator("//h3[text()='Billing System Analysis']//following::span[text()='Match Rate:']//following::span[1]")

    this.lbl_TotalMatchesFound_total = page.locator("(//h3[text()='Total Matches Found']//following::div//div//div)[1]")


    this.lbl_TotalMatchesFound_ExactMatches = page.locator("(//h3[text()='Total Matches Found']//following::div//div//div//div)[1]")

    this.lbl_TotalMatchesFound_oneToMany = page.locator("(//h3[text()='Total Matches Found']//following::div//div//div//div)[2]")

    this.lbl_TotalMatchesFound_manyToMany = page.locator("(//h3[text()='Total Matches Found']//following::div//div//div//div)[3]")
    this.lbl_ReconStatus_totalDeposits = page.locator("//h3[text()='Reconciliation Status']//following::span[contains(text(), 'Total Deposits:')][1]//following::span[1]")
    this.lbl_ReconStatus_totalCashReceipts = page.locator("//h3[text()='Reconciliation Status']//following::span[contains(text(), 'Total Cash Receipts:')][1]//following::span[1]")
    this.lbl_ReconStatus_differences = page.locator("//h3[text()='Reconciliation Status']//following::span[contains(text(), 'Difference:')][1]//following::span[1]")

    this.lbl_cashReconciliationReport = page.locator("//h3[text()='Cash Reconciliation Report']")
    this.tbl_cashReconciliationReport_TotalBillingTransactions = page.locator("//h3[text()='Cash Reconciliation Report']/following::span[count(.|//h3[text()='Bank Deposits by Account:']/preceding::span)=count(//h3[text()='Bank Deposits by Account:']/preceding::span)]")
    this.header_cashReconciliationReport_BankDepositsByAccount = page.locator("//h3[text()='Bank Deposits by Account:']")
    this.header_cashReconciliationReport_BillingSystemAdjustment = page.locator("//h3[text()='Billing System Adjustments:']")
    this.tbl_cashReconciliationReport_BankDepositsByAccount = page.locator("//h3[text()='Bank Deposits by Account:']//following::div[1]//div//span")
    this.header_ReconciliationReport_Deductions = page.locator("//h3[text()='Deductions:']")
    this.tbl_cashReconciliationReport_Deductions = page.locator("//h3[text()='Deductions:']//following::div[1]//div//span")
    this.tbl_cashReconciliationReport_BillingSystemAdjustments = page.locator("//h3[text()='Billing System Adjustments:']//following::div[1]//div//span")
    this.header_UnMatchedBillingSystem = page.locator("//div[text()='Unmatched Billing System Transactions:']")
    this.tbl_cashReconciliationReport_unMatchedBillingSystemTransactions = page.locator("//div[text()='Unmatched Billing System Transactions:']//following::div[1]//span")
    this.header_cashReconciliationReport_Reconciliation = page.locator("//h3[text()='Reconciliation:']")
    this.tbl_cashReconciliationReport_Reconciliation = page.locator("//h3[text()='Reconciliation:']//following::div[1]//div//span")
    this.noReconciltext = page.locator("//div[contains(text(),'No reconciliation history yet')]");


    // New Changes (Harsha)
    this.loadingHistorySpinner = page.locator("//div[normalize-space(text())='Loading history...']")
    this.reconciliationCards = page.locator("//div[@data-radix-scroll-area-viewport]/div/div/div")

    this.matchManuallyBtn_CashReceiptsInBilling = page.locator("//h3[contains(text(),'Cash Receipts in Billing System not Found in Bank')]/following-sibling::div//button[text()='Match Manually']")
    this.matchManuallyBtn_DepositsInBankRFMS = page.locator("//h3[contains(text(),'Deposits in Bank/RFMS')]/following-sibling::div//button[text()='Match Manually']")

    this.tbl_CashReceiptsInBilling = page.locator("//h3[contains(text(),'Cash Receipts in Billing System')]/ancestor::div[2]/following-sibling::div//table")
    this.tbl_DepositsInBankRFMS = page.locator("//h3[contains(text(),'Deposits in Bank/RFMS')]/ancestor::div[2]/following-sibling::div//table")
    this.tbl_InternalBankTransfers = page.locator("(//h3[contains(text(),'Internal Bank Transfers')]/ancestor::div[2]/following-sibling::div//table)[1]")
    this.tbl_BillingSystemTransfers = page.locator("//h4[contains(text(),'Billing System Transfers')]/parent::div//table")
    this.tbl_NSFTransactions = page.locator("//h3[contains(text(),'NSF Transactions')]/ancestor::div[2]/following-sibling::div//table")

    this.headerBillingSystemTransfers = page.locator("//h4[contains(text(),'Billing System Transfers')]")

    this.rows_CashReceiptsInBilling = page.locator("//h3[contains(text(),'Cash Receipts in Billing System')]/ancestor::div[2]/following-sibling::div//table//tbody//tr")
    this.rows_DepositsInBankRFMS = page.locator("//h3[contains(text(),'Deposits in Bank/RFMS')]/ancestor::div[2]/following-sibling::div//table//tbody//tr")
    this.rows_InternalBankTransfers = page.locator("(//h3[contains(text(),'Internal Bank Transfers')]/ancestor::div[2]/following-sibling::div//table)[1]//tbody//tr")
    this.rows_BillingSystemTransfers = page.locator("//h4[contains(text(),'Billing System Transfers')]/parent::div//table//tbody//tr")
    this.rows_NSFTransactions = page.locator("//h3[contains(text(),'NSF Transactions')]/ancestor::div[2]/following-sibling::div//table//tbody//tr")

    this.markExceptionBtns_CashReceiptsInBilling = page.locator("//h3[contains(text(),'Cash Receipts in Billing System')]/ancestor::div[2]/following-sibling::div//table//tbody//tr//button//*[@data-component-name='SquareAsterisk']")
    this.markExceptionBtns_DepositsInBankRFMS = page.locator("//h3[contains(text(),'Deposits in Bank/RFMS')]/ancestor::div[2]/following-sibling::div//table//tbody//tr//button//*[@data-component-name='SquareAsterisk']")

    this.markTransferBtns_CashReceiptsInBilling = page.locator("//h3[contains(text(),'Cash Receipts in Billing System')]/ancestor::div[2]/following-sibling::div//table//tbody//tr//button//*[@data-component-name='ArrowLeftRight']")
    this.markTransferBtns_DepositsInBankRFMS = page.locator("//h3[contains(text(),'Deposits in Bank/RFMS')]/ancestor::div[2]/following-sibling::div//table//tbody//tr//button//*[@data-component-name='ArrowLeftRight']")

    this.markAsTransferBtn = page.locator("//button[normalize-space(text())='Mark as Transfer']")
    this.unmarkAsTransferBtn = page.locator("//button[normalize-space(text())='Unmark as Transfer']")

    this.dateElement = page.locator("//span[normalize-space(text())='Date:']/following-sibling::span")
    this.amountElement = page.locator("//span[normalize-space(text())='Amount:']/following-sibling::span")

    this.unmarkTransferBtns_InternalBankTransfers = page.locator("(//h3[contains(text(),'Internal Bank Transfers')]/ancestor::div[2]/following-sibling::div//table)[1]//tbody//tr//button")
    this.unmarkTransferBtns_BillingSystemTransfers = page.locator("//h4[contains(text(),'Billing System Transfers')]/parent::div//table//tbody//tr//button")
    this.unmarkNSFBtns_NSFTransactions = page.locator("//h3[contains(text(),'NSF Transactions')]/ancestor::div[2]/following-sibling::div//table//tbody//tr//button")

    this.closeBtn = page.locator("//button[normalize-space(text())='Close']")

    this.bulkMarkBtn_DepositsInBankRFMS = page.locator("//h3[text()='Deposits in Bank/RFMS Not Found in Billing System']/following-sibling::div//button[@title='Bulk exception mode']")
    this.bulkMarkBtn_CashReceiptsInBilling = page.locator("//h3[contains(text(),'Cash Receipts in Billing System')]/following-sibling::div//button[@title='Bulk exception mode']")

    this.selectAllCheckbox = page.locator("//button[@aria-label='Select all']")

    this.markAsTransfersBtn = page.locator("//button[normalize-space()='Mark as Transfers']")
    this.markAsExceptionsBtn = page.locator("//button[normalize-space()='Mark as Exceptions']")

    this.bulkMarkAsTransfers = (number) => page.locator(`//button[normalize-space()='Mark ${number} as Transfers']`)
    this.bulkTransfersMarked = page.locator("//div[normalize-space(text())='Bulk transfers marked']")

    this.reasonForException = page.locator("//textarea[@id='reason']")

    this.markAsExceptionBtn = page.locator("//button[normalize-space(text())='Mark as Exception']")
    this.exceptionMarked = page.locator("//div[normalize-space(text())='Exception marked']")

    this.bulkMarkAsExceptions = (number) => page.locator(`//button[normalize-space()='Mark ${number} as Exceptions']`)
    this.bulkExceptionsMarked = page.locator("//div[normalize-space(text())='Bulk exceptions marked']")

    this.unmarkAsNSFBtn = page.locator("//button[normalize-space(text())='Unmark as NSF']")

    this.noMatchesFoundMessage = page.locator("//div[normalize-space(text())='No matches found']")

    this.shareReconciliationBtns = page.locator("//button[@title='Share reconciliation']")
    this.editReconciliationBtns = page.locator("//*[contains(@data-component-name,'Edit')]/parent::button")
    this.deleteReconciliationBtns = page.locator("//*[contains(@data-component-name,'Trash')]/parent::button")

    this.usernameOrEmailInputBox = page.locator("//input[contains(@placeholder,'Search users by name or email')]")
    this.userOptionFromDropdown = (details) => page.locator(`//div[normalize-space(text())='${details}']`)
    this.shareBtn = page.locator("//button[contains(@data-component-file,'ShareReconciliation') and contains(text(),'Share')]")

    this.sharedSuccessfullyMessage = page.locator("//div[normalize-space(text())='Shared Successfully']")
    this.alreadySharedMessage = page.locator("//div[normalize-space(text())='Already Shared']")

    this.deleteBtn = page.locator("//button[normalize-space(text())='Delete']")
    this.sessionDeletedMessage = page.locator("//div[normalize-space(text())='Session Deleted']")

    this.recNameInputBox = page.locator("//div[@data-radix-scroll-area-viewport]/div/div/div//input")
    this.renameSuccesfulMessage = page.locator("//div[normalize-space(text())='Company name updated successfully']")

    this.tbl_MatchedTransactions = page.locator("//h3[text()='Matched Transactions']/ancestor::div[2]/following-sibling::div//table")
    this.rows_MatchedTransactions = page.locator("//h3[text()='Matched Transactions']/ancestor::div[2]/following-sibling::div//table//tbody//tr")
    this.rows_ExactMatchedTransactions = page.locator("//h3[text()='Matched Transactions']/ancestor::div[2]/following-sibling::div//table//tbody//tr[td[7][contains(normalize-space(),'Exact')]]")
    this.rows_ManyToOneMatchedTransactions = page.locator("//h3[text()='Matched Transactions']/ancestor::div[2]/following-sibling::div//table//tbody//tr[td[7][contains(normalize-space(),'Many-to-One')]]")

    this.unmatchSuccessMessage = page.locator("//div[contains(text(),'Transactions moved to unmatched sections')]")

    this.header_bankRFMSTransactionsInManualMatchCreator = page.locator("//div[normalize-space(text())='Bank/RFMS Transactions']")
    this.allBankTransactionsInManualMatchCreator = page.locator("//div[normalize-space(text())='Bank/RFMS Transactions']/following-sibling::div//button/following-sibling::div")
    this.header_billingSystemTransactionsInManualMatchCreator = page.locator("//div[normalize-space(text())='Billing System Transactions']")
    this.allBillingTransactionsInManualMatchCreator = page.locator("//div[normalize-space(text())='Billing System Transactions']/following-sibling::div//button/following-sibling::div")

    this.matchTypeLabel = page.locator("//div[normalize-space(text())='Match Type:']")
    this.createMatchBtn = page.locator("//button[normalize-space(text())='Create Match']")
    this.manualMatchCreatedSuccessfullyMessage = page.locator("//div[normalize-space(text())='Manual Match Created']")
    this.manualBadge = page.locator("//div[@data-component-name='Badge' and normalize-space(text())='Manual']")

    this.totalDeposits_CashReconciliationReport = page.locator("//h3[text()='Reconciliation:']//following::span[contains(text(), 'Total Deposits:')][1]//following::span[1]")
    this.totalCashReceipts_CashReconciliationReport = page.locator("//h3[text()='Reconciliation:']//following::span[contains(text(), 'Total Cash Receipts:')][1]//following::span[1]")
    this.difference_CashReconciliationReport = page.locator("//h3[text()='Reconciliation:']//following::span[contains(text(), 'Difference:')][1]//following::span[1]")

    this.lbl_BFRA_TotalTransactions = page.locator("(//h3[text()='Bank Feed / RFMS Analysis']//following::span[text()='Total Transactions:']//following::div[1])[1]")
    this.lbl_BFRA_MatchedTransactions = page.locator("(//h3[text()='Bank Feed / RFMS Analysis']//following::span[text()='Matched:']//following::div[1]//div[1])[1]")
    this.lbl_BFRA_Transfers = page.locator("//h3[text()='Bank Feed / RFMS Analysis']//following::span[text()='Transfers:']//following::div[1]//div[1]")
    this.lbl_BFRA_unMatched = page.locator("(//h3[text()='Bank Feed / RFMS Analysis']//following::span[text()='Unmatched:']//following::div[1]//div[1])[1]")
    this.lbl_BFRA_matchRate = page.locator("(//h3[text()='Bank Feed / RFMS Analysis']//following::span[text()='Match Rate:']//following::span[1])[1]")

    this.allBankStatementFiles = page.locator("//label[normalize-space(text())='Bank Statement Files']/following-sibling::div//span")
    this.allRFMSFiles = page.locator("//label[normalize-space(text())='RFMS Files']/following-sibling::div//span")
    this.allJournalBillingFiles = page.locator("//label[normalize-space(text())='Journal/Billing System Files']/following-sibling::div//span")

    this.deleteTransactionBtn = page.locator("//button[normalize-space(text())='Delete Transaction']")
    this.transactionDeletedSuccessfullyMessage = page.locator("//div[normalize-space(text())='Transaction Deleted']")

    this.exportToExcelBtn = page.locator("//button[normalize-space(text())='Cash Rec Report']/following-sibling::button[normalize-space(text())='Export to Excel']")
  }

  clickOnFirstCard = async () => {
    await this.noReconciltext.waitFor({ state: 'hidden' });
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    await excuteSteps(this.test, this.firstcard, "click", `Clicking on first card`);
    await this.totalTransaction.waitFor({ state: 'visible' });
  }

  scrollTillRunReconsillationBtn = async () => {
    await excuteSteps(this.test, this.runReconsillationBtn, "scroll", `Scrolling into view`);
  }
  scrollTillReversalTransactions = async () => {
    await excuteSteps(this.test, this.headerReversalTransactions, "scroll", `Scrolling into view`);
  }



  clickOnreconsillationBtn = async () => {
    await excuteSteps(this.test, this.runReconsillationBtn, "click", `Clicking on reconsilation button after file upload`);
  };

  clickOnMatchedTransactions = async () => {
    await excuteSteps(this.test, this.matchedTransactionsHeader, "click", `Clicking on matched transactions`)

  };
  scrollToHeaderNSFTransactions = async () => {
    await excuteSteps(this.test, this.headerNSFTransactions, "scroll", `Scrolling into view`);
  }

  clickOnCashPostingBtn = async () => {
    await excuteSteps(this.test, this.cashPostinggetStartedBtn, "click", `Clicking on cash posting button`);
  };

  clickOnInternalBankTransferHeader = async () => {
    await excuteSteps(this.test, this.headerInternalBankTransfers, "click", `Clicking on internal bank transfers`);
  };


  uploadingFilesInTest = async (bankStatementfilefolder, rfmsFileFolder, billingSystemFileFolder) => {
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
    await this.summaryCard.screenshot({ path: 'screenshots/test.png' });
  }

  uploadingFilesInProd = async (bankStatementfilefolder, rfmsFileFolder, billingSystemFileFolder) => {
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

  matchedTransactionsToExcelTest = async (path) => {
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


  compareAndExportMismatch = async (testEnvData, mismatchSheetName, mismatchFile) => {
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    await this.clickOnMatchedTransactions();
    await this.page.waitForTimeout(parseInt(process.env.largeWait));
    const rows = await this.page.locator("((//h3[text()='Matched Transactions']/following::table)[1]//tr)").all();
    console.log("now ", rows.length);
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

  writingBankFeedSummaryTest = async (path) => {
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

  gettingBandFeedDetailsAndCompareProd = async (testEnvData, mismatchSheetName, mismatchFile) => {
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

    console.log("first card,prod data", prodData);

    const testData = readDataFromExcel(testEnvData, "Bank Feed Analysis");
    console.log("test data", testData);
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
  scrollToHeaderNSFTransactions = async () => {
    await excuteSteps(this.test, this.headerNSFTransactions, "scroll", `Scrolling into view`);
  }

  scrollToHeaderCashReceiptBilling = async () => {
    await excuteSteps(this.test, this.header_cashReceiptsInBIllingSystem, "scroll", `Scrolling into view`);
  }
  CashReceiptsInBillingSystemNotFoundInBankExcelTest = async (path) => {
    await this.scrollToHeaderCashReceiptBilling()
    if (await this.tbl_cashReceiptsInBIllingSystem.isVisible()) {
      await expect.soft(this.tbl_cashReceiptsInBIllingSystem).toBeVisible()
    } else {
      await this.header_cashReceiptsInBIllingSystem.click();
      if (await this.lbl_NoTransactions.isVisible()) {
        console.log("No Internal Bank Transfer transactions found");
        await this.header_cashReceiptsInBIllingSystem.click()
        return; // stop further execution
      }
      await expect.soft(this.tbl_cashReceiptsInBIllingSystem).toBeVisible()
    }
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));

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


  compareAndExportCashReceiptsInBillingSystem = async (testEnvData, mismatchSheetName, mismatchFile) => {
    await scrollToElement(this.header_cashReceiptsInBIllingSystem);
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    if (await this.lbl_NoTransactions.isVisible()) {
      console.log("No Internal Bank Transfer transactions found");
      await this.header_cashReceiptsInBIllingSystem.click()
      return; // stop further execution
    }
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

  scrollToDepositInBankTransaction = async () => {
    await excuteSteps(this.test, this.headerDepositsInBankRFMS, "scroll", `Scrolling into view`);
  }
  DepositsInBankRFMSNotFoundTransactionsToExcelTest = async (path) => {
    await this.scrollToDepositInBankTransaction()
    if (await this.tblDepositsInBankRFMS.isVisible()) {
      await expect.soft(this.tblDepositsInBankRFMS).toBeVisible()
    } else {
      await this.headerDepositsInBankRFMS.click();
      if (await this.lbl_NoTransactions.isVisible()) {
        console.log("No Internal Bank Transfer transactions found");
        await this.headerDepositsInBankRFMS.click()
        return; // stop further execution
      }
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

  compareAndExportDepositsInBankRFMSNotFoundTransactions = async (testEnvData, mismatchSheetName, mismatchFile) => {
    await this.scrollToDepositInBankTransaction()
    if (await this.tblDepositsInBankRFMS.isVisible()) {
      await expect.soft(this.tblDepositsInBankRFMS).toBeVisible()
    } else {
      await this.headerDepositsInBankRFMS.click();
      if (await this.lbl_NoTransactions.isVisible()) {
        console.log("No Internal Bank Transfer transactions found");
        await this.headerDepositsInBankRFMS.click()
        return; // stop further execution
      }
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

  NSFTransactionsToExcelTest = async (path) => {
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    await this.scrollToHeaderNSFTransactions()
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    if (await this.tbl_NSFTransactions.isVisible()) {
      await expect.soft(this.tbl_NSFTransactions).toBeVisible()
    } else {
      await this.headerNSFTransactions.click();
      if (await this.lbl_NoTransactions.isVisible()) {
        console.log("No Internal Bank Transfer transactions found");
        await this.headerNSFTransactions.click()
        return; // stop further execution
      }
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
  compareAndExportNSFTransactionsMismatch = async (testEnvData, mismatchSheetName, mismatchFile) => {
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    await this.scrollToHeaderNSFTransactions()
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    if (await this.tbl_NSFTransactions.isVisible()) {

      await expect(this.tbl_NSFTransactions).toBeVisible()
    } else {
      await this.headerNSFTransactions.click();
      if (await this.lbl_NoTransactions.isVisible()) {
        console.log("No Internal Bank Transfer transactions found");
        await this.headerNSFTransactions.click()
        return; // stop further execution
      }
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

  scrollToReversalTransaction = async () => {
    await excuteSteps(this.test, this.headerReversalTransactions, "scroll", `Scrolling into view`);
  }

  ReversalTransactionsTransactionsToExcelTest = async (path) => {
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    if (await this.tbl_ReversalTransactions.isVisible()) {
      await expect.soft(this.tbl_ReversalTransactions).toBeVisible()
    } else {
      await this.headerReversalTransactions.click();
      if (await this.lbl_NoTransactions.isVisible()) {
        console.log("No Internal Bank Transfer transactions found");
        await this.headerReversalTransactions.click()
        return; // stop further execution
      }
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
  compareAndExportReversalTransactionMismatch = async (testEnvData, mismatchSheetName, mismatchFile) => {
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    if (await this.tbl_ReversalTransactions.isVisible()) {
      await expect.soft(this.tbl_ReversalTransactions).toBeVisible()
    } else {
      await this.headerReversalTransactions.click();
      if (await this.lbl_NoTransactions.isVisible()) {
        console.log("No Internal Bank Transfer transactions found");
        await this.headerReversalTransactions.click()
        return; // stop further execution
      }
      await expect.soft(this.tbl_ReversalTransactions).toBeVisible()
    }

    await this.page.waitForTimeout(parseInt(process.env.mediumWait));

    const rows = await this.tbl_ReversalTransactionsAll.all();
    const prodData = [];
    console.log("reversal transactions count", rows.length);
    for (let i = 1; i < rows.length; i++) {
      const cells = await rows[i].locator("td").allInnerTexts();
      prodData.push({
        BillingSysDate: cells[0].trim(),
        Amount: cells[1].trim(),
        BillingSysDesc: cells[2].trim(),
      });
    }
    console.log("Reversal transactions prod data", prodData);
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
      writeDataToExcel(mismatchFile, mismatchSheetName, mismatches);
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

  scrollToInternalBankTransfers = async () => {
    await excuteSteps(this.test, this.headerInternalBankTransfers, "scroll", `Scrolling into view`);
  }

  InternalBankTransfersTransactionsToExcelTest = async (path) => {
    await scrollToElement(this.headerInternalBankTransfers);
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    if (await this.tbl_InternalBankTransfers.isVisible()) {

      await expect.soft(this.tbl_InternalBankTransfers).toBeVisible()
    } else {
      await this.headerInternalBankTransfers.click();
      if (await this.lbl_NoTransactions.isVisible()) {
        console.log("No Internal Bank Transfer transactions found");
        await this.headerInternalBankTransfers.click()
        return; // stop further execution
      }
      await expect.soft(this.tbl_InternalBankTransfers).toBeVisible()
    }

    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    const rows = await this.tbl_InternalBankTransfersAll.all();
    const tableData = [];
    console.log("No of Rows in Internal Bank Transfers are:", rows.length)
    console.log(rows);
    for (let i = 1; i < rows.length - 1; i++) {
      const cells = await rows[i].locator("td").allInnerTexts();
      tableData.push({
        Date: cells[0].trim(),
        Amount: cells[1].trim(),
        TypeCode: cells[2].trim(),
        Description: cells[3].trim(),
        Status: cells[4].trim(),
      });
    }
    const rowsCount = rows.length - 1
    const cellsData = await rows[rowsCount].locator("td").allInnerTexts();
    console.log("Internal Transfer LastRow Data", rowsCount, cellsData);
    let data1 = cellsData[0].trim()
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
  compareAndExportInternalBankTransfersMismatch = async (testEnvData, mismatchSheetName, mismatchFile) => {
    await scrollToElement(this.headerInternalBankTransfers);
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    if (await this.tbl_InternalBankTransfers.isVisible()) {

      await expect.soft(this.tbl_InternalBankTransfers).toBeVisible()
    } else {
      await this.headerInternalBankTransfers.click();
      if (await this.lbl_NoTransactions.isVisible()) {
        console.log("No Internal Bank Transfer transactions found");
        await this.headerInternalBankTransfers.click()
        return; // stop further execution
      }
      await expect.soft(this.tbl_InternalBankTransfers).toBeVisible()
    }

    await this.page.waitForTimeout(parseInt(process.env.mediumWait));

    const rows = await this.tbl_InternalBankTransfersAll.all();
    const prodData = [];
    for (let i = 1; i < rows.length - 1; i++) {
      const cells = await rows[i].locator("td").allInnerTexts();
      prodData.push({
        Date: cells[0].trim(),
        Amount: cells[1].trim(),
        TypeCode: cells[2].trim(),
        Description: cells[3].trim(),
        Status: cells[4].trim(),
      });
    }
    const rowsCount = rows.length - 1;
    console.log(rowsCount);
    const cellsData = await rows[rowsCount].locator("td").allInnerTexts();
    console.log("Internal Transfer prod LastRow Data", rowsCount, cellsData);
    let data1 = cellsData[0].trim();
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


  scrollToHeaderNDCSweeps = async () => {
    await excuteSteps(this.test, this.headerNDCSweeps, "scroll", `Scrolling into view`);
  }
  NDCSweepsTransactionsToExcelTest = async (path) => {
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    await this.scrollToHeaderNDCSweeps();
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    if (await this.tbl_NDCSweeps.isVisible()) {
      await expect.soft(this.tbl_NDCSweeps).toBeVisible()
    } else {
      await this.headerNDCSweeps.click();
      if (await this.lbl_NoTransactions.isVisible()) {
        console.log("No Internal Bank Transfer transactions found");
        await this.headerNDCSweeps.click();
        return; // stop further execution
      }
      await expect.soft(this.tbl_NDCSweeps).toBeVisible()
    }

    const rows = await this.tbl_NDCSweepsAll.all();
    const tableData = [];
    console.log("NDC Sweeps Table RowCount: ", rows.length)
    console.log(rows);

    for (let i = 1; i < rows.length - 1; i++) {
      const cells = await rows[i].locator("td").allInnerTexts();
      tableData.push({
        Date: cells[0].trim(),
        Amount: cells[1].trim(),
        TypeCode: cells[2].trim(),
        Description: cells[3].trim(),
        BankAccount: cells[4].trim(),
      });
    }

    const rowsCount = rows.length - 1
    const cellsData = await rows[rowsCount].locator("td").allInnerTexts();
    let data1 = cellsData[0].trim()
    let data2 = cellsData[1].trim()
    console.log("The data to write excel", data1, ":", data2);
    tableData.push({
      Date: data1,
      Amount: data2,
      TypeCode: "",
      Description: "",
      BankAccount: "",

    })
    writeDataToExcel(path, "NDCSweeps", tableData);
  }
  compareAndExportNDCSweepsMismatch = async (testEnvData, mismatchSheetName, mismatchFile) => {

    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    await this.scrollToHeaderNDCSweeps();
    await this.page.waitForTimeout(parseInt(process.env.mediumWait));
    if (await this.tbl_NDCSweeps.isVisible()) {
      await expect.soft(this.tbl_NDCSweeps).toBeVisible()
    } else {
      await this.headerNDCSweeps.click()
      if (await this.lbl_NoTransactions.isVisible()) {
        console.log("No Internal Bank Transfer transactions found");
        await this.headerNDCSweeps.click()
        return; // stop further execution
      }
      await expect.soft(this.tbl_NDCSweeps).toBeVisible()
    }

    const rows = await this.tbl_NDCSweepsAll.all();
    const prodData = [];

    for (let i = 1; i < rows.length - 1; i++) {
      const cells = await rows[i].locator("td").allInnerTexts();
      prodData.push({
        Date: cells[0].trim(),
        Amount: cells[1].trim(),
        TypeCode: cells[2].trim(),
        Description: cells[3].trim(),
        BankAccount: cells[4].trim(),
      });
    }
    const rowsCount = rows.length
    const cellsData = await rows[rowsCount - 1].locator("td").allInnerTexts();
    let data1 = cellsData[0].trim()
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
          Test_BankAccount: testRow?.BankAccount || "Missing",
          Prod_Date: prodRow?.Date || "Missing",
          Prod_Amount: prodRow?.Amount || "Missing",
          Prod_TypeCode: prodRow?.TypeCode || "Missing",
          Prod_Description: prodRow?.Description || "Missing",
          Prod_BankAccount: prodRow?.BankAccount || "Missing",
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
            Test_BankAccount: testRow?.BankAccount || "Missing",
            Prod_Date: prodRow?.Date || "Missing",
            Prod_Amount: prodRow?.Amount || "Missing",
            Prod_TypeCode: prodRow?.TypeCode || "Missing",
            Prod_Description: prodRow?.Description || "Missing",
            Prod_BankAccount: prodRow?.BankAccount || "Missing",
          });
        }
      });
    }
    //const mismatchFile = path.resolve("./excel/Mismatches/Mismatch_NDCSweeps.xlsx");
    if (mismatches.length > 0) {
      console.log("Mismatches found. Exporting detailed report to Excel...");

      writeDataToExcel(mismatchFile, mismatchSheetName, mismatches);
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

  scrollToHeaderCashReconciliationReport = async () => {
    await excuteSteps(this.test, this.lbl_cashReconciliationReport, "scroll", `Scrolling into view`);
  }

  scrollToHeaderCashReconciliationDeductionReport = async () => {
    await excuteSteps(this.test, this.header_ReconciliationReport_Deductions, "scroll", `Scrolling into view`);
  }


  scrollToHeaderCashReconciliationBankDepositsByAccount = async () => {
    await excuteSteps(this.test, this.header_cashReconciliationReport_BankDepositsByAccount, "scroll", `Scrolling into view`);
  }
  scrollToHeaderCashReconciliationBillingSystemAdjustments = async () => {
    await excuteSteps(this.test, this.header_cashReconciliationReport_BillingSystemAdjustment, "scroll", `Scrolling into view`);
  }

  scrollToHeaderCashReconciliationReconciliation = async () => {
    await excuteSteps(this.test, this.header_cashReconciliationReport_Reconciliation, "scroll", `Scrolling into view`);
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
    const tableData = []
    const data = []
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

  compareAndExportReconReconciliationMismatch = async () => {
    const prodData = [];
    const prodDataWrite = [];
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

      await test.step("Verified, All rows in CashReconRep_Data matched successfully!", async () => {
        console.log("All rows matched successfully!");
      })
    }
  }

  compareAndExportReconBillingSystemAdjustmentsMismatch = async () => {
    const prodData = [];
    const prodDataWrite = [];
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

      await test.step("Verified, All rows in CashReconRep_Adjustments matched successfully!", async () => {
        console.log("All rows matched successfully!");
      })
    }
  }

  compareAndExportReconUnMatchedSystemTransMismatch = async () => {
    const prodData = [];
    const prodDataWrite = [];
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

      await test.step("Verified, All rows in CashReconRep_unMatchedSystemTrans matched successfully!", async () => {
        console.log("All rows matched successfully!");
      })
    }
  }

  compareAndExportReconTotalBillingTransMismatch = async () => {
    const prodData = [];
    const prodDataWrite = [];
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

      await test.step("Verified, All rows in CashReconRep_TotalBilling matched successfully!", async () => {
        console.log("All rows matched successfully!");
      })
    }
  }

  compareAndExportBankDepositByAccountMismatch = async () => {
    const prodData = [];
    const prodDataWrite = [];
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

      await test.step("Verified, All rows in CashReconRep_BankDepositByAccount matched successfully!", async () => {
        console.log("All rows matched successfully!");
      })
    }
  }


  compareAndExportDeductionsMismatch = async () => {
    const prodData = [];
    const prodDataWrite = [];
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

      await test.step("Verified, All rows in CashReconRep_BankDepositByAccount matched successfully!", async () => {
        console.log("All rows matched successfully!");
      })
    }
  }


  CashReconciliationReportToExcelTest = async (path) => {
    await scrollToElement(this.lbl_cashReconciliationReport);
    const tableData = [];
    const totalBillingrows = await this.tbl_cashReconciliationReport_TotalBillingTransactions.all();
    let rowsCount = await totalBillingrows.length;

    for (let i = 0; i < rowsCount; i += 2) {
      let dataText1 = await totalBillingrows[i].textContent();
      let dataText2 = await totalBillingrows[i + 1].textContent();

      tableData.push({
        "A": dataText1,
        "B": dataText2,
        "C": "",
      });
    }
    const rows_BankDeposits = await this.tbl_cashReconciliationReport_BankDepositsByAccount.all();

    for (let i = 0; i < rows_BankDeposits.length; i += 2) {
      let dataText1 = await rows_BankDeposits[i].textContent();
      let dataText2 = await rows_BankDeposits[i + 1].textContent();
      tableData.push({
        "A": dataText1,
        "B": dataText2,
        "C": "",
      });
    }

    const rows_Deductions = await this.tbl_cashReconciliationReport_Deductions.all();

    for (let i = 0; i < rows_Deductions.length; i += 2) {
      let dataText1 = await rows_Deductions[i].textContent();
      let dataText2 = await rows_Deductions[i + 1].textContent();

      tableData.push({
        "A": dataText1,
        "B": dataText2,
        "C": "",
      });
    }

    const rows_BillingSystemAdjustments = await this.tbl_cashReconciliationReport_BillingSystemAdjustments.all();

    for (let i = 0; i < rows_BillingSystemAdjustments.length; i += 2) {
      let dataText1 = await rows_BillingSystemAdjustments[i].textContent();
      let dataText2 = await rows_BillingSystemAdjustments[i + 1].textContent();
      tableData.push({
        "A": dataText1,
        "B": dataText2,
        "C": "",
      });
    }

    const rows_unMatchedBillingSystemTransactions = await this.tbl_cashReconciliationReport_unMatchedBillingSystemTransactions.all();
    let rows_unmatchedBilling = await rows_unMatchedBillingSystemTransactions.length;

    for (let i = 0; i < rows_unmatchedBilling - 3; i += 3) {
      let dataText1 = await rows_unMatchedBillingSystemTransactions[i].textContent();
      let dataText2 = await rows_unMatchedBillingSystemTransactions[i + 1].textContent();
      let dataText3 = await rows_unMatchedBillingSystemTransactions[i + 2].textContent();

      tableData.push({
        "A": dataText1,
        "B": dataText2,
        "C": dataText3,
      });
    }
    let rowCount = rows_unmatchedBilling - 2
    let dataText1 = await rows_unMatchedBillingSystemTransactions[rowCount].textContent();
    let dataText2 = await rows_unMatchedBillingSystemTransactions[rowCount + 1].textContent();

    tableData.push({
      "A": dataText1,
      "B": dataText2,
      "C": "",
    });

    const rows_cashReconciliationReport_Reconciliation = await this.tbl_cashReconciliationReport_Reconciliation.all();
    for (let i = 0; i < rows_cashReconciliationReport_Reconciliation.length; i += 2) {
      let dataText1 = await rows_cashReconciliationReport_Reconciliation[i].textContent();
      let dataText2 = await rows_cashReconciliationReport_Reconciliation[i + 1].textContent();

      tableData.push({
        "A": dataText1,
        "B": dataText2,
        "C": "",
      });
    }

    console.log(tableData);
    writeDataToExcel(path, "TotalBilling", tableData);

  }



  writingBankBillingSystemAnalysisTest = async (path) => {
    const totalTransactions = await this.lbl_BSA_TotalTransactions.innerText();
    const matched = await this.lbl_BSA_MatchedTransactions.innerText();
    const nSFs = await this.lbl_BSA_NSF.innerText();
    const reversals = await this.lbl_BSA_Reversals.innerText();
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



  writingBankBillingSystemAnalysisTotalMatchesFoundTest = async (path) => {
    const totalTransactions = await this.lbl_TotalMatchesFound_total.innerText();
    const exactMatches = await this.lbl_TotalMatchesFound_ExactMatches.innerText();
    const oneToMany = await this.lbl_TotalMatchesFound_oneToMany.innerText();
    const manyToMany = await this.lbl_TotalMatchesFound_manyToMany.innerText();

    const testData = [{
      "TotalTransactions": totalTransactions,
      "ExactMatches": exactMatches,
      "OneToMany": oneToMany,
      "ManyToMany": manyToMany
    }];

    writeDataToExcel(path, "TotalMatchesFound", testData);
  }


  writingBankBillingSystemAnalysisReconciliationStatusTest = async (path) => {
    const totalDeposits = await this.lbl_ReconStatus_totalDeposits.innerText();
    const totalCashReceipts = await this.lbl_ReconStatus_totalCashReceipts.innerText();
    const differences = await this.lbl_ReconStatus_differences.innerText();
    console.log({ totalDeposits, totalCashReceipts, differences });

    const testData = [{
      "TotalDeposits": totalDeposits,
      "TotalCashReceipts": totalCashReceipts,
      "Differences": differences
    }];

    writeDataToExcel(path, "ReconciliationStatus", testData);
  }


  gettingBandFeedDetailsBIllingSystemAnalysisAndCompareProd = async (testEnvData, mismatchSheetName, mismatchFile) => {
    const totalTransactions = await this.lbl_BSA_TotalTransactions.innerText();
    const matched = await this.lbl_BSA_MatchedTransactions.innerText();
    const nSFs = await this.lbl_BSA_NSF.innerText();
    const reversals = await this.lbl_BSA_Reversals.innerText();
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
    console.log("prodData", prodData);
    console.log("testdata", testData);
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

  gettingBandFeedDetailsTotalMatchesFoundAndCompareProd = async (testEnvData, mismatchSheetName, mismatchFile) => {
    const totalTransactions = await this.lbl_TotalMatchesFound_total.innerText();
    const exactMatches = await this.lbl_TotalMatchesFound_ExactMatches.innerText();
    const oneToMany = await this.lbl_TotalMatchesFound_oneToMany.innerText();
    const manyToMany = await this.lbl_TotalMatchesFound_manyToMany.innerText();

    const prodData = [{
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

  gettingBandFeedDetailsReconciliationStatusAndCompareProd = async (testEnvData, mismatchSheetName, mismatchFile) => {

    const totalDeposits = await this.lbl_ReconStatus_totalDeposits.innerText();
    const totalCashReceipts = await this.lbl_ReconStatus_totalCashReceipts.innerText();
    const differences = await this.lbl_ReconStatus_differences.innerText();
    const prodData = [{
      "TotalDeposits": totalDeposits,
      "TotalCashReceipts": totalCashReceipts,
      "Differences": differences
    }];

    const testData = readDataFromExcel(testEnvData, "ReconciliationStatus");
    const mismatches = getMismatches(prodData, testData);
    //const mismatchFile = path.resolve("./excel/Mismatches/Mismatch_ReconciliationStatus.xlsx");
    if (mismatches.length > 0) {
      console.log("Mismatches found:", { mismatches });
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
          XLSX.writeFile(wb, mismatchFile);
          console.log(`Old mismatch sheet "${mismatchSheetName}" deleted.`);
        }
      }

      await test.step("Verified, All test data matched", async () => {
        console.log("All data matched, old mismatch sheet removed if present!");
      });

    }
  }


  compareAndExportCashReconciliationReportMismatch = async (testEnvData, mismatchSheetName, mismatchFile) => {

    await scrollToElement(this.lbl_cashReconciliationReport);
    const prodData = [];
    const totalBillingrows = await this.tbl_cashReconciliationReport_TotalBillingTransactions.all();
    let rowsCount = await totalBillingrows.length;

    for (let i = 0; i < rowsCount; i += 2) {
      let dataText1 = await totalBillingrows[i].textContent();
      let dataText2 = await totalBillingrows[i + 1].textContent();

      prodData.push({
        "A": dataText1,
        "B": dataText2,
        "C": "",
      });
    }
    const rows_BankDeposits = await this.tbl_cashReconciliationReport_BankDepositsByAccount.all();

    for (let i = 0; i < rows_BankDeposits.length; i += 2) {
      let dataText1 = await rows_BankDeposits[i].textContent();
      let dataText2 = await rows_BankDeposits[i + 1].textContent();
      prodData.push({
        "A": dataText1,
        "B": dataText2,
        "C": "",
      });
    }

    const rows_Deductions = await this.tbl_cashReconciliationReport_Deductions.all();

    for (let i = 0; i < rows_Deductions.length; i += 2) {
      let dataText1 = await rows_Deductions[i].textContent();
      let dataText2 = await rows_Deductions[i + 1].textContent();

      prodData.push({
        "A": dataText1,
        "B": dataText2,
        "C": "",
      });
    }

    const rows_BillingSystemAdjustments = await this.tbl_cashReconciliationReport_BillingSystemAdjustments.all();

    for (let i = 0; i < rows_BillingSystemAdjustments.length; i += 2) {
      let dataText1 = await rows_BillingSystemAdjustments[i].textContent();
      let dataText2 = await rows_BillingSystemAdjustments[i + 1].textContent();
      prodData.push({
        "A": dataText1,
        "B": dataText2,
        "C": "",
      });
    }

    const rows_unMatchedBillingSystemTransactions = await this.tbl_cashReconciliationReport_unMatchedBillingSystemTransactions.all();
    let rows_unmatchedBilling = await rows_unMatchedBillingSystemTransactions.length;

    for (let i = 0; i < rows_unmatchedBilling - 3; i += 3) {
      let dataText1 = await rows_unMatchedBillingSystemTransactions[i].textContent();
      let dataText2 = await rows_unMatchedBillingSystemTransactions[i + 1].textContent();
      let dataText3 = await rows_unMatchedBillingSystemTransactions[i + 2].textContent();

      prodData.push({
        "A": dataText1,
        "B": dataText2,
        "C": dataText3,
      });
    }
    let rowCount = rows_unmatchedBilling - 2
    let dataText1 = await rows_unMatchedBillingSystemTransactions[rowCount].textContent();
    let dataText2 = await rows_unMatchedBillingSystemTransactions[rowCount + 1].textContent();

    prodData.push({
      "A": dataText1,
      "B": dataText2,
      "C": "",
    });

    const rows_cashReconciliationReport_Reconciliation = await this.tbl_cashReconciliationReport_Reconciliation.all();
    for (let i = 0; i < rows_cashReconciliationReport_Reconciliation.length; i += 2) {
      let dataText1 = await rows_cashReconciliationReport_Reconciliation[i].textContent();
      let dataText2 = await rows_cashReconciliationReport_Reconciliation[i + 1].textContent();

      prodData.push({
        "A": dataText1,
        "B": dataText2,
        "C": "",
      });
    }
    console.log("Last card prod data", prodData);
    const testData = readDataFromExcel(testEnvData, "TotalBilling");
    console.log("------Last card test data-------- ", testData);
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

  navigateToCashPosting = async () => {
    await this.clickOnCashPostingBtn()
    await this.loadingHistorySpinner.waitFor({ state: 'hidden' })
    await this.page.waitForTimeout(parseInt(process.env.smallWait));
  }

  clickMarkAsTransferBtn = async () => {
    await excuteSteps(this.test, this.markAsTransferBtn, "click", `Clicking on Mark as Transfer Button`)
  }

  clickMarkAsExceptionBtn = async () => {
    await excuteSteps(this.test, this.markAsExceptionBtn, "click", `Clicking on Mark as Exception Button`)
  }

  clickUnmarkAsTransferBtn = async () => {
    await excuteSteps(this.test, this.unmarkAsTransferBtn, "click", `Clicking on Unmark as Transfer Button`)
  }

  ensureCashReceiptsTableVisible = async () => {
    if (!(await this.tbl_CashReceiptsInBilling.isVisible())) {
      await this.header_cashReceiptsInBIllingSystem.click();
    }
  }

  ensureBankRFMSTableVisible = async () => {
    if (!(await this.tbl_DepositsInBankRFMS.isVisible())) {
      await this.headerDepositsInBankRFMS.click();
    }
  }

  ensureInternalBankTransfersTableVisible = async () => {
    if (!(await this.tbl_InternalBankTransfers.isVisible())) {
      await this.headerInternalBankTransfers.click();
    }
  }

  ensureBillingSystemTransfersTableVisible = async () => {
    if (!(await this.tbl_InternalBankTransfers.isVisible())) {
      await this.headerInternalBankTransfers.click();
    }
  }

  markAsBillingTransfer = async () => {
    await this.header_cashReceiptsInBIllingSystem.waitFor({ state: 'visible' });
    await this.ensureCashReceiptsTableVisible();

    // If match manually button is not visible, return null to continue with the next card
    //if (!(await this.matchManuallyBtn_CashReceiptsInBilling.isVisible())) return null;
    if (await this.markTransferBtns_CashReceiptsInBilling.count() === 0) return null;

    const description = await this.markTransferBtns_CashReceiptsInBilling.first().locator("//preceding::td[2]").innerText();
    await this.markTransferBtns_CashReceiptsInBilling.first().click();
    const date = await this.dateElement.innerText();
    const amount = await this.amountElement.innerText();
    await this.clickMarkAsTransferBtn();

    const transactionDetails = { description, date, amount };

    await this.ensureBillingSystemTransfersTableVisible();
    const transferredTransaction = this.rows_BillingSystemTransfers
      .filter({ hasText: transactionDetails.date })
      .filter({ hasText: transactionDetails.amount })
      .filter({ hasText: transactionDetails.description })
      .first();

    await expect(transferredTransaction).toBeVisible();
    const transactionStatus = await transferredTransaction.locator("//td[5]").innerText();
    expect.soft(transactionStatus.trim()).toBe("Manually marked as transfer from billing system")
    return transactionDetails; // Marked as transfer successfully
  }

  markAsBankTransfer = async () => {
    await this.headerDepositsInBankRFMS.waitFor({ state: 'visible' });
    await this.ensureBankRFMSTableVisible();

    //if (!(await this.matchManuallyBtn_DepositsInBankRFMS.isVisible())) return null;
    if (await this.markTransferBtns_DepositsInBankRFMS.count() === 0) return null;

    const description = await this.markTransferBtns_DepositsInBankRFMS.first()
      .locator("//preceding::td[2]").innerText();
    await this.markTransferBtns_DepositsInBankRFMS.first().click();
    const date = await this.dateElement.innerText();
    const amount = await this.amountElement.innerText();
    await this.clickMarkAsTransferBtn();

    const transactionDetails = { description, date, amount }

    await this.ensureInternalBankTransfersTableVisible();
    const transferredTransaction = this.rows_InternalBankTransfers
      .filter({ hasText: transactionDetails.date })
      .filter({ hasText: transactionDetails.amount })
      .filter({ hasText: transactionDetails.description })
      .first();

    await expect(transferredTransaction).toBeVisible();
    const transactionStatus = await transferredTransaction.locator("//td[5]").innerText();
    expect.soft(transactionStatus.trim()).toBe("Manually marked as transfer")
    return transactionDetails;
  }

  clickCloseBtn = async () => {
    await excuteSteps(this.test, this.closeBtn, "click", `Clicking on Close Button`)
  }

  unmarkAsBankTransfer = async () => {
    await this.headerDepositsInBankRFMS.waitFor({ state: 'visible' });
    await this.ensureInternalBankTransfersTableVisible();
    if (await this.unmarkTransferBtns_InternalBankTransfers.count() === 0) return false;

    const description = await this.unmarkTransferBtns_InternalBankTransfers.first().locator("//preceding::td[2]").innerText();
    await this.unmarkTransferBtns_InternalBankTransfers.first().click();
    const date = await this.dateElement.innerText();
    const amount = await this.amountElement.innerText();
    await this.clickUnmarkAsTransferBtn();
    await this.page.waitForTimeout(parseInt(process.env.smallWait));

    if (await this.closeBtn.isVisible()) await this.clickCloseBtn();

    await this.ensureBankRFMSTableVisible();
    const transferredTransaction = this.rows_DepositsInBankRFMS
      .filter({ hasText: date })
      .filter({ hasText: amount })
      .filter({ hasText: description })
      .first();

    await expect(transferredTransaction).toBeVisible();
    const transactionStatus = await transferredTransaction.locator("//td[4]").innerText();
    expect.soft(transactionStatus.trim()).toBe("Manually marked as non-transfer")
    return true;
  }

  unmarkAsBillingTransfer = async () => {
    await this.header_cashReceiptsInBIllingSystem.waitFor({ state: 'visible' });
    await this.ensureBillingSystemTransfersTableVisible();

    if (await this.unmarkTransferBtns_BillingSystemTransfers.count() === 0) return false;

    const description = await this.unmarkTransferBtns_BillingSystemTransfers.first().locator("//preceding::td[2]").innerText();
    await this.unmarkTransferBtns_BillingSystemTransfers.first().click();
    const date = await this.dateElement.innerText();
    const amount = await this.amountElement.innerText();
    await this.clickUnmarkAsTransferBtn();
    await this.page.waitForTimeout(parseInt(process.env.smallWait));

    if (await this.closeBtn.isVisible()) await this.clickCloseBtn();

    await this.ensureCashReceiptsTableVisible();
    const transferredTransaction = this.rows_CashReceiptsInBilling
      .filter({ hasText: date })
      .filter({ hasText: amount })
      .filter({ hasText: description })
      .first();

    await expect(transferredTransaction).toBeVisible();
    const transactionStatus = await transferredTransaction.locator("//td[4]").innerText();
    expect.soft(transactionStatus.trim()).toBe("Unmatched");
    return true;
  }

  clickBulkMarkBtn_DepositsInBankRFMs = async () => {
    await excuteSteps(this.test, this.bulkMarkBtn_DepositsInBankRFMS, "click", `Clicking on Bulk mark in Bank/RFMS Deposits Table`)
  }

  clickBulkMarkBtn_CashReceiptsInBilling = async () => {
    await excuteSteps(this.test, this.bulkMarkBtn_CashReceiptsInBilling, "click", `Clicking on Bulk mark in Cash Receipts Billing Table`)
  }

  checkSelectAllCheckbox = async () => {
    await excuteSteps(this.test, this.selectAllCheckbox, "check", `Checking Select All Checkbox to Bulk Mark`)
  }

  clickMarkAsTransfersBtn = async () => {
    await excuteSteps(this.test, this.markAsTransfersBtn, "click", `Clicking on Mark as Transfers Button to Bulk Mark`)
  }

  clickMarkAsExceptionsBtn = async () => {
    await excuteSteps(this.test, this.markAsExceptionsBtn, "click", `Clicking on Mark as Exceptions Button to Bulk Mark`)
  }

  clickBulkMarkAsTransfersBtn = async (number) => {
    await excuteSteps(this.test, this.bulkMarkAsTransfers(number), "click", `Clicking on Mark ${number} as Transfers Btn`, number)
  }

  clickBulkMarkAsExceptionsBtn = async (number) => {
    await excuteSteps(this.test, this.bulkMarkAsExceptions(number), "click", `Clicking on Mark ${number} as Exceptions Btn`, number)
  }

  bulkMarkAsBankTransfer = async () => {
    await this.headerDepositsInBankRFMS.waitFor({ state: 'visible' });
    await this.ensureBankRFMSTableVisible();

    if (!(await this.bulkMarkBtn_DepositsInBankRFMS.isVisible())) return false;

    const numOfTransactions = await this.rows_DepositsInBankRFMS.count();

    let transactions = [];
    for (let i = 0; i < numOfTransactions; i++) {
      const description = await this.markTransferBtns_DepositsInBankRFMS.nth(i).locator("//preceding::td[2]").innerText();
      const amount = await this.markTransferBtns_DepositsInBankRFMS.nth(i).locator("//preceding::td[3]").innerText();
      const date = await this.markTransferBtns_DepositsInBankRFMS.nth(i).locator("//preceding::td[4]").innerText();

      transactions.push({ description, amount, date });
    }

    await this.clickBulkMarkBtn_DepositsInBankRFMs();
    await this.checkSelectAllCheckbox();

    if (!(await this.markAsTransfersBtn.isVisible())) return false;

    await this.clickMarkAsTransfersBtn();
    await this.clickBulkMarkAsTransfersBtn(numOfTransactions);

    await this.bulkTransfersMarked.waitFor({ state: 'visible' });
    await this.bulkTransfersMarked.waitFor({ state: 'hidden' });

    await this.ensureInternalBankTransfersTableVisible();

    for (const transaction of transactions) {
      const matchingRows = this.page.locator(
        `(//h3[contains(text(),'Internal Bank Transfers')]/ancestor::div[2]/following-sibling::div//table)[1]//tbody//tr[
                normalize-space(td[1]) = "${transaction.date}" and
                normalize-space(td[2]) = "${transaction.amount}" and
                normalize-space(td[4]) = "${transaction.description}"
                ]`
      );
      const count = await matchingRows.count();
      expect(count).toBeGreaterThan(0);

      for (let j = 0; j < count; j++) {
        const statusText = await matchingRows
          .nth(j)
          .locator("//td[5]")
          .innerText();

        expect.soft(statusText.trim()).toBe("Manually marked as transfer");
      }
    }
    return true;
  }

  bulkMarkAsBillingTransfer = async () => {
    await this.header_cashReceiptsInBIllingSystem.waitFor({ state: 'visible' });
    await this.ensureCashReceiptsTableVisible();

    if (!(await this.bulkMarkBtn_CashReceiptsInBilling.isVisible())) return false;

    const numOfTransactions = await this.rows_CashReceiptsInBilling.count();

    let transactions = [];
    for (let i = 0; i < numOfTransactions; i++) {
      const description = await this.markTransferBtns_CashReceiptsInBilling.nth(i).locator("//preceding::td[2]").innerText();
      const amount = await this.markTransferBtns_CashReceiptsInBilling.nth(i).locator("//preceding::td[3]").innerText();
      const date = await this.markTransferBtns_CashReceiptsInBilling.nth(i).locator("//preceding::td[4]").innerText();

      transactions.push({ description, amount, date });
    }

    await this.clickBulkMarkBtn_CashReceiptsInBilling();
    await this.checkSelectAllCheckbox();

    if (!(await this.markAsTransfersBtn.isVisible())) return false;

    await this.clickMarkAsTransfersBtn();
    await this.clickBulkMarkAsTransfersBtn(numOfTransactions);

    await this.bulkTransfersMarked.waitFor({ state: 'visible' });
    await this.bulkTransfersMarked.waitFor({ state: 'hidden' });

    await this.ensureBillingSystemTransfersTableVisible();

    for (const transaction of transactions) {
      const matchingRows = this.page.locator(
        `//h4[contains(text(),'Billing System Transfers')]/parent::div//table//tbody//tr[
                normalize-space(td[1]) = "${transaction.date}" and
                normalize-space(td[2]) = "${transaction.amount}" and
                normalize-space(td[4]) = "${transaction.description}"
                ]`
      );
      const count = await matchingRows.count();
      expect(count).toBeGreaterThan(0);

      for (let j = 0; j < count; j++) {
        const statusText = await matchingRows
          .nth(j)
          .locator("//td[5]")
          .innerText();

        expect.soft(statusText.trim()).toBe("Manually marked as transfer from billing system");
      }
    }
    return true;
  }

  enterReasonForMarkingException = async (text) => {
    await excuteSteps(this.test, this.reasonForException, "fill", `Enter Reason for Marking as Exception`, [text]);
  }

  markAsBillingException = async (reason) => {
    await this.header_cashReceiptsInBIllingSystem.waitFor({ state: 'visible' });
    await this.ensureCashReceiptsTableVisible();

    // If there are no mark exception buttons, return null to continue with the next card
    if (await this.markExceptionBtns_CashReceiptsInBilling.count() === 0) return null;

    const transactionDetails = {
      description: await this.markExceptionBtns_CashReceiptsInBilling.first().locator("//preceding::td[2]").innerText(),
      date: await this.markExceptionBtns_CashReceiptsInBilling.first().locator("//preceding::td[4]").innerText(),
      amount: await this.markExceptionBtns_CashReceiptsInBilling.first().locator("//preceding::td[3]").innerText()
    }

    await this.markExceptionBtns_CashReceiptsInBilling.first().click();

    await this.enterReasonForMarkingException(reason);
    await this.clickMarkAsExceptionBtn();

    await expect(this.exceptionMarked).toBeVisible({ timeout: 10000 });

    const exemptedTransaction = this.rows_CashReceiptsInBilling
      .filter({ hasText: transactionDetails.date })
      .filter({ hasText: transactionDetails.amount })
      .filter({ hasText: transactionDetails.description })
      .first();

    const transactionStatus = exemptedTransaction.locator("//td[4]")
    await expect.soft(transactionStatus).toContainText(`Marked as Exception${reason}`)
    return transactionDetails; // Marked as exception successfully
  }

  markAsBankException = async (reason) => {
    await this.headerDepositsInBankRFMS.waitFor({ state: 'visible' });
    await this.ensureBankRFMSTableVisible();

    // If there are no mark exception buttons, return null to continue with the next card
    if (await this.markExceptionBtns_DepositsInBankRFMS.count() === 0) return null;

    const transactionDetails = {
      description: await this.markExceptionBtns_DepositsInBankRFMS.first()
        .locator("//preceding::td[2]").innerText(),
      date: await this.markExceptionBtns_DepositsInBankRFMS.first()
        .locator("//preceding::td[4]").innerText(),
      amount: await this.markExceptionBtns_DepositsInBankRFMS.first()
        .locator("//preceding::td[3]").innerText()
    }

    await this.markExceptionBtns_DepositsInBankRFMS.first().click();

    await this.enterReasonForMarkingException(reason);
    await this.clickMarkAsExceptionBtn();

    await expect(this.exceptionMarked).toBeVisible({ timeout: 10000 });

    const exemptedTransaction = this.rows_DepositsInBankRFMS
      .filter({ hasText: transactionDetails.date })
      .filter({ hasText: transactionDetails.amount })
      .filter({ hasText: transactionDetails.description })
      .first();

    const transactionStatus = exemptedTransaction.locator("//td[4]")
    await expect.soft(transactionStatus).toContainText(`Marked as Exception${reason}`)
    return transactionDetails; // Marked as exception successfully
  }

  bulkMarkAsBillingException = async (reason) => {
    await this.header_cashReceiptsInBIllingSystem.waitFor({ state: 'visible' });
    await this.ensureCashReceiptsTableVisible();

    if (!(await this.bulkMarkBtn_CashReceiptsInBilling.isVisible())) return false;

    const numOfTransactions = await this.rows_CashReceiptsInBilling.count();

    let transactions = [];
    for (let i = 0; i < numOfTransactions; i++) {
      const description = await this.markTransferBtns_CashReceiptsInBilling.nth(i)
        .locator("//preceding::td[2]").innerText();
      const amount = await this.markTransferBtns_CashReceiptsInBilling.nth(i)
        .locator("//preceding::td[3]").innerText();
      const date = await this.markTransferBtns_CashReceiptsInBilling.nth(i)
        .locator("//preceding::td[4]").innerText();

      transactions.push({ description, amount, date });
    }

    await this.clickBulkMarkBtn_CashReceiptsInBilling();
    await this.checkSelectAllCheckbox();

    if (!(await this.markAsExceptionsBtn.isVisible())) return false;
    await this.clickMarkAsExceptionsBtn();

    await this.enterReasonForMarkingException(reason);
    await this.clickBulkMarkAsExceptionsBtn(numOfTransactions);

    await this.bulkExceptionsMarked.waitFor({ state: 'visible', timeout: numOfTransactions * 10000 });

    for (const transaction of transactions) {
      const matchingRows = this.page.locator(
        `//h3[contains(text(),'Cash Receipts in Billing System')]/ancestor::div[2]/following-sibling::div//table//tbody//tr[
          normalize-space(td[1]) = "${transaction.date}" and
          normalize-space(td[2]) = "${transaction.amount}" and
          normalize-space(td[3]) = "${transaction.description}"
        ]`
      );
      const count = await matchingRows.count();
      expect(count).toBeGreaterThan(0);

      for (let j = 0; j < count; j++) {
        const status = await matchingRows
          .nth(j)
          .locator("//td[4]")

        await expect.soft(status).toContainText(`Marked as Exception${reason}`);
      }
    }
    return true;
  }

  bulkMarkAsBankException = async (reason) => {
    await this.headerDepositsInBankRFMS.waitFor({ state: 'visible' });
    await this.ensureBankRFMSTableVisible();

    if (!(await this.bulkMarkBtn_DepositsInBankRFMS.isVisible())) return false;

    const numOfTransactions = await this.rows_DepositsInBankRFMS.count();

    let transactions = [];
    for (let i = 0; i < numOfTransactions; i++) {
      const description = await this.markTransferBtns_DepositsInBankRFMS.nth(i).locator("//preceding::td[2]").innerText();
      const amount = await this.markTransferBtns_DepositsInBankRFMS.nth(i).locator("//preceding::td[3]").innerText();
      const date = await this.markTransferBtns_DepositsInBankRFMS.nth(i).locator("//preceding::td[4]").innerText();

      transactions.push({ description, amount, date });
    }

    await this.clickBulkMarkBtn_DepositsInBankRFMs();
    await this.checkSelectAllCheckbox();

    if (!(await this.markAsExceptionsBtn.isVisible())) return false;
    await this.clickMarkAsExceptionsBtn();

    await this.enterReasonForMarkingException(reason);
    await this.clickBulkMarkAsExceptionsBtn(numOfTransactions);

    await this.bulkExceptionsMarked.waitFor({ state: 'visible', timeout: 60000 });
    await this.bulkExceptionsMarked.waitFor({ state: 'hidden', timeout: 30000 });

    for (const transaction of transactions) {
      const matchingRows = this.page.locator(
        `//h3[contains(text(),'Deposits in Bank/RFMS')]/ancestor::div[2]/following-sibling::div//table//tbody//tr[
          normalize-space(td[1]) = "${transaction.date}" and
          normalize-space(td[2]) = "${transaction.amount}" and
          normalize-space(td[3]) = "${transaction.description}"
        ]`
      );
      const count = await matchingRows.count();
      expect(count).toBeGreaterThan(0);

      for (let j = 0; j < count; j++) {
        const status = await matchingRows
          .nth(j)
          .locator("//td[4]")

        await expect.soft(status).toContainText(`Marked as Exception${reason}`);
      }
    }
    return true;
  }

  ensureNSFTableVisible = async () => {
    if (!(await this.tbl_NSFTransactions.isVisible())) {
      await this.headerNSFTransactions.click();
    }
  }

  clickUnmarkAsNSFBtn = async () => {
    await excuteSteps(this.test, this.unmarkAsNSFBtn, "click", `Clicking on Unmark as NSF Button`)
  }

  unmarkNSFTransaction = async () => {
    await this.headerInternalBankTransfers.waitFor({ state: 'visible' });
    if (!(await this.headerNSFTransactions.isVisible())) return false;
    await this.ensureNSFTableVisible();

    const date = await this.rows_NSFTransactions.first().locator("//td[1]").innerText();
    const amount = await this.rows_NSFTransactions.first().locator("//td[2]").innerText();
    const description = await this.rows_NSFTransactions.first().locator("//td[3]").innerText();

    await this.unmarkNSFBtns_NSFTransactions.first().click();
    await this.clickUnmarkAsNSFBtn();

    await expect(this.noMatchesFoundMessage).toBeVisible({ timeout: 10000 });
    await this.noMatchesFoundMessage.waitFor({ state: 'hidden' });

    await this.ensureCashReceiptsTableVisible();

    const unmarkedTransaction = this.rows_CashReceiptsInBilling
      .filter({ hasText: date })
      .filter({ hasText: amount })
      .filter({ hasText: description })
      .first();

    await expect(unmarkedTransaction).toBeVisible();

    const statusText = await unmarkedTransaction.locator("//td[4]").innerText();
    expect.soft(statusText.trim()).toBe('NSF manually unmarked');

    return true;
  }

  hoverOnReconciliationCard = async (number) => {
    await excuteSteps(this.test, this.reconciliationCards.nth(number - 1), "hover", `Hovering on '${number}' card`, number)
  }

  clickShareReconciliationCard = async (number) => {
    await excuteSteps(this.test, this.shareReconciliationBtns.nth(number - 1), "click", `Clicking on Share Button of '${number}' card`, number)
  }

  clickDeleteReconciliationCard = async (number) => {
    await excuteSteps(this.test, this.deleteReconciliationBtns.nth(number - 1), "click", `Clicking on Delete Button of '${number}' card`, number)
  }

  clickEditReconciliationCard = async (number) => {
    await excuteSteps(this.test, this.editReconciliationBtns.nth(number - 1), "click", `Clicking on Edit Button of '${number}' card`, number)
  }

  enterUserDetails = async (details) => {
    await excuteSteps(this.test, this.usernameOrEmailInputBox, "fill", `Enter User Details to Share Reconciliation with`, [details])
  }

  selectUserFromOptions = async (details) => {
    await excuteSteps(this.test, this.userOptionFromDropdown(details), "click", `Selecting User from Options to Share Reconciliation`, details)
  }

  clickShareBtn = async () => {
    await excuteSteps(this.test, this.shareBtn, "click", `Clicking on Share Button`)
  }

  clickDeleteBtn = async () => {
    await excuteSteps(this.test, this.deleteBtn, "click", `Clicking on Delete Button`)
  }

  shareAReconciliation = async (cardNumber, userDetails) => {
    await this.hoverOnReconciliationCard(cardNumber);
    await this.clickShareReconciliationCard(cardNumber);
    await this.enterUserDetails(userDetails);
    await this.selectUserFromOptions(userDetails);
    await this.clickShareBtn();

    await Promise.race([
      this.alreadySharedMessage.waitFor({ state: 'visible', timeout: 30000 }),
      this.sharedSuccessfullyMessage.waitFor({ state: 'visible', timeout: 30000 })
    ]);
  }

  deleteAReconciliation = async (cardNumber) => {
    const cardsCountBefore = await this.reconciliationCards.count();

    await this.hoverOnReconciliationCard(cardNumber);
    await this.clickDeleteReconciliationCard(cardNumber);
    await this.clickDeleteBtn();

    await expect(this.sessionDeletedMessage).toBeVisible({ timeout: 30000 });
    await this.sessionDeletedMessage.waitFor({ state: 'hidden' });

    const cardsCountAfter = await this.reconciliationCards.count();

    expect(cardsCountAfter).toBe(cardsCountBefore - 1);
  }

  enterReconciliationName = async (name) => {
    await excuteSteps(this.test, this.recNameInputBox, "fill", `Enter a name to rename Reconciliation`, [name])
  }

  renameAReconciliation = async (cardNumber, cardName) => {
    await this.hoverOnReconciliationCard(cardNumber);
    await this.clickEditReconciliationCard(cardNumber);
    await this.enterReconciliationName(cardName);
    await this.page.keyboard.press("Enter");

    await expect(this.renameSuccesfulMessage).toBeVisible({ timeout: 30000 });
    await this.renameSuccesfulMessage.waitFor({ state: 'hidden' });
  }

  ensureMatchedTransactionsTableVisible = async () => {
    if (!(await this.tbl_MatchedTransactions.isVisible()))
      await this.clickOnMatchedTransactions();
  }

  unmatchATransaction = async () => {
    await this.matchedTransactionsHeader.waitFor({ state: 'visible' });
    await this.ensureMatchedTransactionsTableVisible();

    if (await this.rows_ExactMatchedTransactions.count() === 0) return null;

    const firstRow = this.rows_ExactMatchedTransactions.first();

    const transactionDetails = {
      billingDate: await firstRow.locator("//td[1]").innerText(),
      bankDate: await firstRow.locator("//td[2]").innerText(),
      billingAmt: await firstRow.locator("//td[3]").innerText(),
      bankAmt: await firstRow.locator("//td[4]").innerText(),
      billingDesc: await firstRow.locator("//td[5]").innerText(),
      bankDesc: await firstRow.locator("//td[6]").innerText()
    };

    await excuteSteps(this.test, firstRow, "hover", `Hovering on First Row`);

    const unmatchTransactionBtn = firstRow.locator("//button");
    await excuteSteps(this.test, unmatchTransactionBtn, "click", `Clicking on Unmatch Transaction Button`);

    await expect(this.unmatchSuccessMessage).toBeVisible({ timeout: 15000 });

    const transferredBankTransaction = this.rows_DepositsInBankRFMS
      .filter({ hasText: transactionDetails.bankDate })
      .filter({ hasText: transactionDetails.bankAmt })
      .filter({ hasText: transactionDetails.bankDesc })
      .first();

    await expect(transferredBankTransaction, 'Unmatched Transaction Should be in Billing Section')
      .toBeVisible();

    const transferredBillingTransaction = this.rows_CashReceiptsInBilling
      .filter({ hasText: transactionDetails.billingDate })
      .filter({ hasText: transactionDetails.billingAmt })
      .filter({ hasText: transactionDetails.billingDesc })
      .first();

    await expect(transferredBillingTransaction, 'Unmatched Transaction Should be in Bank RFMS Section')
      .toBeVisible();

    return transactionDetails;
  }

  clickMatchManuallyInCashReceipts = async () => {
    await excuteSteps(this.test, this.matchManuallyBtn_CashReceiptsInBilling, "click", `Clicking on Match manually Button beside Cash Receipts header`)
  }

  clickMatchManuallyInBankRFMS = async () => {
    await excuteSteps(this.test, this.matchManuallyBtn_DepositsInBankRFMS, "click", `Clicking on Match manually Button beside Deposits in Bank/RFMS header`)
  }

  clickCreateMatchBtn = async () => {
    await excuteSteps(this.test, this.createMatchBtn, "click", `Clicking on Create Match Button`)
  }

  createAManualMatch = async () => {
    const transactionDetails = await this.unmatchATransaction();
    if (!transactionDetails) return null;

    await this.clickMatchManuallyInCashReceipts();
    await this.header_bankRFMSTransactionsInManualMatchCreator.waitFor({ state: 'visible' });
    const billingCheckbox = this.allBillingTransactionsInManualMatchCreator
      .filter({ hasText: transactionDetails.billingDate })
      .filter({ hasText: transactionDetails.billingAmt })
      .locator("//preceding-sibling::button")
      .first();

    await excuteSteps(this.test, billingCheckbox, "check", `Clicking on Billing Transaction Checkbox`)

    const bankCheckbox = this.allBankTransactionsInManualMatchCreator
      .filter({ hasText: transactionDetails.bankDate })
      .filter({ hasText: transactionDetails.bankAmt })
      .locator("//preceding-sibling::button")
      .first();

    await excuteSteps(this.test, bankCheckbox, "check", `Clicking on Banking Transaction Checkbox`)

    await expect.soft(this.matchTypeLabel).toContainText('Exact Match');
    await this.clickCreateMatchBtn();

    await expect(this.manualMatchCreatedSuccessfullyMessage).toBeVisible({ timeout: 30000 });
    
    await this.ensureMatchedTransactionsTableVisible();

    const matchedTransactionRow = this.rows_ExactMatchedTransactions
      .filter({ hasText: transactionDetails.bankDate })
      .filter({ hasText: transactionDetails.bankAmt })
      .filter({ hasText: transactionDetails.bankDesc })
      .filter({ hasText: transactionDetails.billingDate })
      .filter({ hasText: transactionDetails.billingAmt })
      .filter({ hasText: transactionDetails.billingDesc })
      .first();

    // Verify Manual badge
    const manualBadge = matchedTransactionRow.locator(this.manualBadge);
    await expect(manualBadge).toBeVisible({ timeout: 10000 });

    console.log("Manual badge verified for the matched transaction.");
    return transactionDetails;
  }

  verifyReconciliationStatusTile = async () => {
    await this.reconciliationCards.nth(0).click();

    await this.lbl_ReconStatus_totalDeposits.waitFor({ state: 'visible' })
    const totalDeposits_ReconStatusTile =
      (await this.lbl_ReconStatus_totalDeposits.innerText()).trim();
    const totalCashReceipts_ReconStatusTile =
      (await this.lbl_ReconStatus_totalCashReceipts.innerText()).trim();
    const difference_ReconStatusTile =
      (await this.lbl_ReconStatus_differences.innerText()).trim();

    const totalDeposits_CashReconReport =
      (await this.totalDeposits_CashReconciliationReport.innerText()).trim();
    const totalCashReceipts_CashReconReport =
      (await this.totalCashReceipts_CashReconciliationReport.innerText()).trim();
    const difference_CashReconReport =
      (await this.difference_CashReconciliationReport.innerText()).trim();

    expect.soft(totalDeposits_ReconStatusTile).toBe(totalDeposits_CashReconReport)
    expect.soft(totalCashReceipts_ReconStatusTile).toBe(totalCashReceipts_CashReconReport)
    expect.soft(difference_ReconStatusTile).toBe(difference_CashReconReport)
  }

  verifyExceptionNotAvailableInManualMatch = async (transactionDetails) => {
    if (!(await this.matchManuallyBtn_CashReceiptsInBilling.isVisible())) return;

    await this.clickMatchManuallyInCashReceipts();
    await this.header_bankRFMSTransactionsInManualMatchCreator.waitFor({ state: 'visible' });

    const billingTransaction = this.allBillingTransactionsInManualMatchCreator
      .filter({ hasText: transactionDetails.date })
      .filter({ hasText: transactionDetails.amount })
      .filter({ hasText: transactionDetails.description });

    const bankTransaction = this.allBankTransactionsInManualMatchCreator
      .filter({ hasText: transactionDetails.date })
      .filter({ hasText: transactionDetails.amount })
      .filter({ hasText: transactionDetails.description });

    // Assertion - should NOT exist
    await expect(billingTransaction).toHaveCount(0);
    await expect(bankTransaction).toHaveCount(0);
  };

  verifyTransferNotAvailableInManualMatch = async (transactionDetails) => {
    if (!(await this.matchManuallyBtn_CashReceiptsInBilling.isVisible())) return;

    await this.clickMatchManuallyInCashReceipts();
    await this.header_bankRFMSTransactionsInManualMatchCreator.waitFor({ state: 'visible' });

    const billingTransaction = this.allBillingTransactionsInManualMatchCreator
      .filter({ hasText: transactionDetails.date })
      .filter({ hasText: transactionDetails.amount })
      .filter({ hasText: transactionDetails.description });

    const bankTransaction = this.allBankTransactionsInManualMatchCreator
      .filter({ hasText: transactionDetails.date })
      .filter({ hasText: transactionDetails.amount })
      .filter({ hasText: transactionDetails.description });

    // Assertion - should NOT exist
    await expect(billingTransaction).toHaveCount(0);
    await expect(bankTransaction).toHaveCount(0);
  };

  verifyMatchedNotAvailableInManualMatch = async (transactionDetails) => {
    if (!(await this.matchManuallyBtn_CashReceiptsInBilling.isVisible())) return;

    await this.clickMatchManuallyInCashReceipts();
    await this.header_bankRFMSTransactionsInManualMatchCreator.waitFor({ state: 'visible' });

    const billingTransaction = this.allBillingTransactionsInManualMatchCreator
      .filter({ hasText: transactionDetails.billingDate })
      .filter({ hasText: transactionDetails.billingAmt })
      .filter({ hasText: transactionDetails.billingDesc });

    const bankTransaction = this.allBankTransactionsInManualMatchCreator
      .filter({ hasText: transactionDetails.bankDate })
      .filter({ hasText: transactionDetails.bankAmt })
      .filter({ hasText: transactionDetails.bankDesc });

    // Validation
    await expect(billingTransaction).toHaveCount(0);
    await expect(bankTransaction).toHaveCount(0);
  }

  getTotalTransactionsBFRA = async () => {
    const totalTransactions = (await this.lbl_BFRA_TotalTransactions.innerText()).trim();
    return Number(totalTransactions);
  }

  getTotalTransactionsBSA = async () => {
    const totalTransactions = (await this.lbl_BSA_TotalTransactions.innerText()).trim();
    return Number(totalTransactions);
  }

  getMatchedCountBFRA = async () => {
    const matched = (await this.lbl_BFRA_MatchedTransactions.innerText()).trim();
    return Number(matched);
  }

  getMatchedCountBSA = async () => {
    const matched = (await this.lbl_BSA_MatchedTransactions.innerText()).trim();
    return Number(matched);
  }

  getTransfersCountBFRA = async () => {
    const transfers = (await this.lbl_BFRA_Transfers.innerText()).trim();
    return Number(transfers);
  }

  getUnmatchedCountBFRA = async () => {
    const unMatched = (await this.lbl_BFRA_unMatched.innerText()).trim();
    return Number(unMatched);
  }

  getUnmatchedCountBSA = async () => {
    const unMatched = (await this.lbl_BSA_unMatched.innerText()).trim();
    return Number(unMatched);
  }

  getMatchRateBFRA = async () => {
    const matchRate = (await this.lbl_BFRA_matchRate.innerText()).split('%')[0];
    return Number(matchRate);
  }

  getMatchRateBSA = async () => {
    const matchRate = (await this.lbl_BSA_matchRate.innerText()).split('%')[0];
    return Number(matchRate);
  }

  getNSFCountBSA = async () => {
    const nSFs = (await this.lbl_BSA_NSF.innerText()).trim();
    return Number(nSFs);
  }

  clickDeleteTransactionBtn = async () => {
    await excuteSteps(this.test, this.deleteTransactionBtn, "click", `Clicking on Delete Transaction Button`)
  }

  hoverAndClickDeleteTransaction = async () => {
    // Loop through both tables
    const tables = [this.rows_CashReceiptsInBilling, this.rows_DepositsInBankRFMS];

    for (const rows of tables) {
      for (let i = 0; i < await rows.count(); i++) {
        const row = rows.nth(i);

        // Hover the row
        await excuteSteps(this.test, row, "hover", `Hovering on row ${i + 1}`);
        await this.page.waitForTimeout(300);

        // Check if delete button exists
        const trashBtn = row.locator("//*[contains(@data-component-name,'Trash')]/parent::button");
        if (await trashBtn.count() > 0) {
          await excuteSteps(this.test, trashBtn, "click", `Clicking Trash button`);
          await this.clickDeleteTransactionBtn();
          await this.transactionDeletedSuccessfullyMessage.waitFor({ state: 'visible' });
          return true; // Transaction Deleted Successfully
        }
      }
    }

    // No delete button found in any row
    return false;
  };

  readUploadedPdfFiles = async () => {
    let pdfFound = false;
    let allPdfTransactions = [];
    let deleteBtnFoundInCard = false;

    const sections = [this.allBankStatementFiles, this.allRFMSFiles, this.allJournalBillingFiles];

    for (const section of sections) {
      const count = await section.count();

      for (let i = 0; i < count; i++) {
        const fileLocator = section.nth(i);
        const fileName = (await fileLocator.textContent()).trim().toLowerCase();
        if (!fileName.endsWith(".pdf")) continue;
        console.log(`PDF found: ${fileName}`);
        pdfFound = true;

        const downloadBtn = fileLocator.locator("//following-sibling::button");
        if (!(await downloadBtn.isVisible()) || !(await downloadBtn.isEnabled())) continue;

        const eventPromise = Promise.race([
          this.page.waitForEvent("download"),
          this.page.context().waitForEvent("page")
        ]);
        await excuteSteps(this.test, downloadBtn, "click", `Downloading/opening pdf: ${fileName}`);
        const event = await eventPromise;

        let buffer;
        if ("path" in event) {
          // pdf downloads (headless)
          const path = await event.path();
          buffer = fs.readFileSync(path);
        } else {
          // pdf opens in a new page (headed)
          const newPage = event;
          await newPage.waitForLoadState("domcontentloaded");
          const pdfUrl = newPage.url();
          // Fetch PDF content
          const response = await newPage.request.get(pdfUrl);
          buffer = await response.body();
          await newPage.close();
        }

        // Parse PDF
        const pdfData = await pdfParse(buffer);
        const pdfTransactions = await this.extractTransactions(pdfData.text);

        // Merge transactions from all PDFs
        allPdfTransactions.push(...pdfTransactions);
      }
    }

    if (!pdfFound) return false;

    console.log("All PDF Transactions Combined:");
    console.log(allPdfTransactions);

    await this.ensureBankRFMSTableVisible();
    const uiTransactions = await this.getUITableTransactions();

    console.log("UI Transactions:");
    console.log(uiTransactions);

    for (const tx of uiTransactions) {
      const transactionExistsInPDF = allPdfTransactions.some(t =>
        t.date === tx.date &&
        t.amount.replace(/\$/g, '').replace(/,/g, '').trim() ===
        tx.amount.replace(/\$/g, '').replace(/,/g, '').trim() &&
        (
          t.description.replace(/[^a-zA-Z\s]/g, '').replace(/\s+/g, ' ').trim()
            .includes(tx.description.replace(/[^a-zA-Z\s]/g, '').replace(/\s+/g, ' ').trim()) ||
          tx.description.replace(/[^a-zA-Z\s]/g, '').replace(/\s+/g, ' ').trim()
            .includes(t.description.replace(/[^a-zA-Z\s]/g, '').replace(/\s+/g, ' ').trim())
        )
      );

      await excuteSteps(this.test, tx.rowLocator, "hover", `Hovering on Bank Transaction Row`);
      await this.page.waitForTimeout(parseInt(process.env.smallWait));
      const deleteBtn = tx.rowLocator.locator(
        "//*[contains(@data-component-name,'Trash')]/parent::button"
      );

      if (transactionExistsInPDF) {
        console.log(`${tx.description} -> Transaction coming from PDF -> Delete allowed`);
        await expect(deleteBtn).toBeVisible();
        deleteBtnFoundInCard = true;
      } else {
        console.log(`${tx.description} -> Transaction not coming from PDF -> Delete NOT allowed`);
        await expect(deleteBtn).toBeHidden();
      }
    }

    return deleteBtnFoundInCard;
  };

  extractTransactions = async (pdfText) => {
    const pdfTransactions = [];
    // Remove lines with only long numbers
    const cleanedText = pdfText.replace(/^\d{10,}$/gm, "");
    const regex = /(\d{2}\/\d{2}\/\d{4})\s*([^\$]+)\$([\d,]+\.\d{2})\$([\d,]+\.\d{2})/g;

    let match;
    while ((match = regex.exec(cleanedText)) !== null) {
      pdfTransactions.push({
        date: match[1],
        description: match[2].trim(),
        amount: match[3]
      });
    }

    return pdfTransactions;
  };

  getUITableTransactions = async () => {
    const rowCount = await this.rows_DepositsInBankRFMS.count();

    const uiTransactions = [];
    for (let i = 0; i < rowCount; i++) {
      const row = this.rows_DepositsInBankRFMS.nth(i);
      const date = (await row.locator("//td[1]").textContent()).trim();
      const amount = (await row.locator("//td[2]").textContent()).trim();
      const description = (await row.locator("//td[3]").textContent()).trim();

      uiTransactions.push({
        date,
        description,
        amount,
        rowLocator: row
      });
    }

    return uiTransactions;
  };

  deleteTransactionIfPdfExists = async () => {
    // Get all file names from the three sections
    const bankFiles = await this.allBankStatementFiles.allTextContents();
    const rfmsFiles = await this.allRFMSFiles.allTextContents();
    const journalFiles = await this.allJournalBillingFiles.allTextContents();

    // Combine all file names
    const allFiles = [...bankFiles, ...rfmsFiles, ...journalFiles];

    // Check if any file ends with .pdf
    const hasPdf = allFiles.some(fileName => fileName.trim().toLowerCase().endsWith('.pdf'));
    if (!hasPdf) return false; // No PDF file found in this card, continue

    console.log("PDF found in this card");
    await this.ensureCashReceiptsTableVisible();
    await this.ensureBankRFMSTableVisible();

    return await this.hoverAndClickDeleteTransaction();
  }

  clickExportToExcelBtn = async (subfolder, fileName) => {
    const [download] = await Promise.all([
      this.page.waitForEvent('download'),
      excuteSteps(this.test, this.exportToExcelBtn, "click", `Clicking on Export to Excel Button`)
    ]);
    const outputFolder = path.join(process.cwd(), "output", subfolder);
    if (!fs.existsSync(outputFolder)) {
      fs.mkdirSync(outputFolder, { recursive: true });
    }
    const filePath = path.join(outputFolder, fileName);
    await download.saveAs(filePath);   // waits + saves properly
    return filePath;
  }

}
