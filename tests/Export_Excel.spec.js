const { test, expect } = require("@playwright/test");
const sections = require("../pageObjects/UI_Pages/pageIndex");
const { readDataFromExcel, writeDataToExcel } = require("../utilities/readExcel");
const { getMismatches } = require("../utilities/getMismatches");
const { compareExcelsSheetWise } = require("../utilities/excelCompare");
require("dotenv").config();

test("Compare Excel Export Report - Preview vs Prod", async ({ browser }) => {
    // ---------- PREVIEW ----------
    const previewContext = await browser.newContext();
    const previewPage = await previewContext.newPage();

    const previewLoginPage = new sections.LoginPage(test, previewPage);
    await previewLoginPage.launchingApplication([process.env.base_url_env]);
    await previewLoginPage.loginToLovable([process.env.lovableUsername], [process.env.lovablePassword]);
    await previewLoginPage.loginWithValidCredentials([process.env.user_name], [process.env.password]);

    const preview = new sections.CashPosting(test, previewPage);
    await preview.navigateToCashPosting();
    await preview.reconciliationCards.nth(0).click();
    await preview.exportToExcelBtn.waitFor({ state: 'visible' });

    const previewFile = await preview.clickExportToExcelBtn('preview.xlsx');

    await previewContext.close();

    // ---------- PROD ----------
    const prodContext = await browser.newContext();
    const prodPage = await prodContext.newPage();

    const prodLoginPage = new sections.LoginPage(test, prodPage);
    await prodLoginPage.launchingApplication([process.env.base_url_prod]);
    await prodLoginPage.loginWithValidCredentials([process.env.user_name], [process.env.password]);

    const prod = new sections.CashPosting(test, prodPage);
    await prod.navigateToCashPosting();
    await prod.reconciliationCards.nth(0).click();
    await prod.exportToExcelBtn.waitFor({ state: 'visible' });

    const prodFile = await prod.clickExportToExcelBtn('prod.xlsx');

    const hasMismatch = await compareExcelsSheetWise(previewFile, prodFile);

    expect.soft(hasMismatch, `Excel mismatch detected. See output/Mismatch_Report_SheetWise.xlsx`)
        .toBe(false);
});