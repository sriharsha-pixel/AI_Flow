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

// BFRA - Bank Feed/RFMS Analysis, BSA - Billing System Analysis
test("Verify BFRA and BSA Matched, Unmatched and Match Rate update after Unmatching a Transaction", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        await cashPosting.matchedTransactionsHeader.waitFor({ state: 'visible' });

        const matchedCountBFRABefore = await cashPosting.getMatchedCountBFRA();
        const matchedCountBSABefore = await cashPosting.getMatchedCountBSA();
        const unMatchedCountBFRABefore = await cashPosting.getUnmatchedCountBFRA();
        const unMatchedCountBSABefore = await cashPosting.getUnmatchedCountBSA();
        const matchRateBFRABefore = await cashPosting.getMatchRateBFRA();
        const matchRateBSABefore = await cashPosting.getMatchRateBSA();

        const details = await cashPosting.unmatchATransaction();
        if (details) {
            const matchedCountBFRAAfter = await cashPosting.getMatchedCountBFRA();
            const matchedCountBSAAfter = await cashPosting.getMatchedCountBSA();
            const unMatchedCountBFRAAfter = await cashPosting.getUnmatchedCountBFRA();
            const unMatchedCountBSAAfter = await cashPosting.getUnmatchedCountBSA();
            const matchRateBFRAAfter = await cashPosting.getMatchRateBFRA();
            const matchRateBSAAfter = await cashPosting.getMatchRateBSA();

            // Assertion
            expect(matchedCountBFRAAfter).toBe(matchedCountBFRABefore - 1);
            expect(matchedCountBSAAfter).toBe(matchedCountBSABefore - 1);
            expect(unMatchedCountBFRAAfter).toBe(unMatchedCountBFRABefore + 1);
            expect(unMatchedCountBSAAfter).toBe(unMatchedCountBSABefore + 1);
            expect(matchRateBFRAAfter).toBeLessThan(matchRateBFRABefore);
            expect(matchRateBSAAfter).toBeLessThan(matchRateBSABefore);

            return;
        }
    }
})

test("Verify BFRA Transfers count update after Marking as Bank Transfer", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        await cashPosting.matchedTransactionsHeader.waitFor({ state: 'visible' });

        const transfersCountBFRABefore = await cashPosting.getTransfersCountBFRA();

        const details = await cashPosting.markAsBankTransfer();
        if (details) {
            const transfersCountBFRAAfter = await cashPosting.getTransfersCountBFRA();
            // Assertion
            expect(transfersCountBFRAAfter).toBe(transfersCountBFRABefore + 1);

            return;
        }
    }
})

test("Verify BSA NSFs Count update after Unmarking NSF", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();
    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        await cashPosting.matchedTransactionsHeader.waitFor({ state: 'visible' });

        const nSFsCountBefore = await cashPosting.getNSFCountBSA();

        const details = await cashPosting.unmarkNSFTransaction();
        if (details) {
            const nSFsCountAfter = await cashPosting.getNSFCountBSA();

            // Assertion
            expect(nSFsCountAfter).toBe(nSFsCountBefore - 1);

            return;
        }
    }
})

test("Verify BFRA/BSA Total Transactions Count Update after Deleting a Transaction", async ({ page }) => {
    const cashPosting = new sections.CashPosting(test, page);
    await cashPosting.navigateToCashPosting();

    const cardCount = await cashPosting.reconciliationCards.count();
    for (let i = 0; i < cardCount; i++) {
        await cashPosting.reconciliationCards.nth(i).click();
        await cashPosting.headerDepositsInBankRFMS.waitFor({ state: 'visible' });
        const totalTransactionsBefore = await cashPosting.getTotalTransactionsBFRA() + await cashPosting.getTotalTransactionsBSA();

        const deleted = await cashPosting.deleteTransactionIfPdfExists();

        if (deleted) {
            const totalTransactionsAfter = await cashPosting.getTotalTransactionsBFRA() + await cashPosting.getTotalTransactionsBSA();
            expect(totalTransactionsAfter).toBe(totalTransactionsBefore - 1)
            break; // Stop after first successful delete
        } else {
            console.log(`Card #${i + 1} has no Transaction to Delete, moving to next card`);
        }
    }
});

