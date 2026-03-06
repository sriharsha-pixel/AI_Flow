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

test("Mark as Billing Transfer", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        
        const details = await cashPosting.markAsBillingTransfer();
        if (details) {
            return; // exit loop once marked transfer
        }
    }
})

test("Unmark as Billing Transfer", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        const isUnmarked = await cashPosting.unmarkAsBillingTransfer();
        if (isUnmarked) {
            // If the transfer was successfully unmarked, exit the loop
            return;
        } else {
            continue; // Continue to next card if unmark transfer was not done
        }
    }
})

test("Bulk Mark as Billing Transfer", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        const bulkMarked = await cashPosting.bulkMarkAsBillingTransfer();
        if (bulkMarked) {
            // If transfers were successfully bulk marked, exit the loop
            return;
        } else {
            continue; // Continue to next card if bulk transfer was not done
        }
    }
})