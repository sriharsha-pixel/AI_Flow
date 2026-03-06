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

// Enter the card number to share a particular Reconciliation and userdetails (Username or password) 
// to share to that particular person before Running
test("Share a Reconciliation", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    await cashPosting.shareAReconciliation(5, 'Esther Furst');
})

// Enter Card Number to Delete that particular Reconciliation
test("Delete a Reconciliation", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    await cashPosting.deleteAReconciliation(10);
})

// Enter card number and a name to rename a particular Reconciliation
test("Rename a Reconciliation", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    await cashPosting.renameAReconciliation(15, 'New Name');
})