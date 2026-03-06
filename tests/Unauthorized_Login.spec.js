const { test, expect } = require("@playwright/test");
const sections = require("../pageObjects/UI_Pages/pageIndex");
const { log } = require("console");
require("dotenv").config();

test("Verify Signing in to AI Flow using Unauthorized Credentials", async ({ page }) => {
    const loginPage = new sections.LoginPage(test, page);
    // prod
    await loginPage.launchingApplication([process.env.base_url_prod]);
    await loginPage.loginWithValidUnauthorizedCredentials(
        [process.env.unauthorized_username],
        [process.env.unauthorized_password]
    );
});
