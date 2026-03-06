const { test, expect } = require("@playwright/test");
const sections = require("../pageObjects/UI_Pages/pageIndex");
const path = require("path");
require("dotenv").config();

test.beforeEach("Login to AI Flow", async ({ page }) => {
    const loginPage = new sections.LoginPage(test, page);
    await loginPage.launchingApplication([process.env.base_url_env]);
    await loginPage.loginToLovable([process.env.lovableUsername], [process.env.lovablePassword]);
    await loginPage.loginWithValidCredentials([process.env.user_name], [process.env.password]);
});

test("Verify Deleting a Bank Transaction is only possible for transactions from a PDF File", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();

    const cardCount = await cashPosting.reconciliationCards.count();

    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        await cashPosting.headerDepositsInBankRFMS.waitFor({ state: 'visible' });
    }
})

test("Verify Deleting a Transaction", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();

    const cardCount = await cashPosting.reconciliationCards.count();

    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        await cashPosting.headerDepositsInBankRFMS.waitFor({ state: 'visible' });

        const deleted = await cashPosting.deleteTransactionIfPdfExists();

        if (deleted) {
            console.log(`Transaction Delete in card #${i + 1}`);
            break; // Stop after first successful delete
        } else {
            console.log(`no Transaction to Delete in Card #${i + 1}, moving to next card`);
        }
    }
});