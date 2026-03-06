const { test, expect } = require("@playwright/test");
const sections = require("../pageObjects/UI_Pages/pageIndex");
const path = require("path");
require("dotenv").config();

let reasonForException = 'Testing Purpose'

test.beforeEach("Login to AI Flow", async ({ page }) => {
    const loginPage = new sections.LoginPage(test, page);
    await loginPage.launchingApplication([process.env.base_url_env]);
    await loginPage.loginToLovable([process.env.lovableUsername], [process.env.lovablePassword]);
    await loginPage.loginWithValidCredentials([process.env.user_name], [process.env.password]);
});

test("Verify Unmatching a Transaction", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        const details = await cashPosting.unmatchATransaction();
        if (details) {
            return; // exit loop once unmatched
        }
    }
})

test("Verify Creating a Manual Match", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        const details = await cashPosting.createAManualMatch();
        if (details) {
            return; // exit if match was created
        }
    }
})

test("Verify Creating a Manual Match for Transaction marked as Bank Exception", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        const details = await cashPosting.markAsBankException(reasonForException);
        if (details) {
            await cashPosting.verifyExceptionNotAvailableInManualMatch(details);
            return;
        }
    }
})

test("Verify Creating a Manual Match for Transaction marked as Billing Exception", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        const details = await cashPosting.markAsBillingException(reasonForException);
        if (details) {
            await cashPosting.verifyExceptionNotAvailableInManualMatch(details);
            return;
        }
    }
})

test("Verify Creating a Manual Match for Transaction marked as Bank Transfer", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        const details = await cashPosting.markAsBankTransfer();
        if (details) {
            await cashPosting.verifyTransferNotAvailableInManualMatch(details);
            return;
        }
    }
})

test("Verify Creating a Manual Match for Transaction marked as Billing Transfer", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        const details = await cashPosting.markAsBillingTransfer();
        if (details) {
            await cashPosting.verifyTransferNotAvailableInManualMatch(details);
            return;
        }
    }
})

test("Verify Creating a Manual Match for Transaction which is already Matched manually", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        const details = await cashPosting.createAManualMatch();
        if (details) {
            await cashPosting.verifyMatchedNotAvailableInManualMatch(details);
            return;
        }
    }
})