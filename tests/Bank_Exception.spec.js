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

test("Mark as Bank Exception", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();

        const details = await cashPosting.markAsBankException(reasonForException);
        if (details) {
            return; // exit loop once marked exception
        }
    }
})

test("Bulk Mark as Bank Exception", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        
        const bulkMarked = await cashPosting.bulkMarkAsBankException(reasonForException);
        if (bulkMarked) {
            // If Exceptions were successfully bulk marked, exit the loop
            return;
        } else {
            continue; // Continue to next card if bulk Exception was not done
        }
    }
})