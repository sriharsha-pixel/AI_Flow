// @ts-check
import { defineConfig, devices } from "@playwright/test";
const path = require('path');
/**
 * Read environment variables from file.
 * https://github.com/motdotla/dotenv
 */
// import dotenv from 'dotenv';
// import path from 'path';
// dotenv.config({ path: path.resolve(__dirname, '.env') });
 const timestamp = new Date()
  .toLocaleString('en-GB', { timeZone: 'Asia/Kolkata' })              // Get local time in 'en-GB' format
  .replace(/[/, ]/g, '-')                // Replace slashes and spaces with dashes
  .replace(/:/g, '-');  
/**
 * @see https://playwright.dev/docs/test-configuration
 */
export default defineConfig({
  testDir: "./tests",
  timeout: 6* 60 * 1000,
  /* Deletes the existing allure report files before test execution*/
  globalSetup: require.resolve("./global_setup.js"),
  /* Run tests in files in parallel */
  fullyParallel: false,
  /* Fail the build on CI if you accidentally left test.only in the source code. */
  forbidOnly: !!process.env.CI,
  /* Retry on CI only */
  retries: process.env.CI ? 1 : 0,
  /* Opt out of parallel tests on CI. */
  workers: process.env.CI ? 1 :1,
  /* Reporter to use. See https://playwright.dev/docs/test-reporters */
  reporter: [["html",{outputFolder: path.join(__dirname, 'playwright-report', `html-report-${timestamp}`),open: 'always',}], ["allure-playwright"]],
  /* Shared settings for all the projects below. See https://playwright.dev/docs/api/class-testoptions.  ,['./utilities/customHtmlReporter']   */
  use: {
    /* Base URL to use in actions like `await page.goto('/')`. */
    // baseURL: 'http://localhost:3000',
 
    /* Collect trace when retrying the failed test. See https://playwright.dev/docs/trace-viewer */
    trace: "on-first-retry",
    headless: false,
    screenshot: "only-on-failure",
    video: "retain-on-failure",
  },
 
  /* Configure projects for major browsers */
  projects: [
    {
      name: "chromium",
      use: { ...devices["Desktop Chrome"] },
    },
 
    // {
    //   name: 'firefox',
    //   use: { ...devices['Desktop Firefox'] },
    // },
 
    // {
    //   name: 'webkit',
    //   use: { ...devices['Desktop Safari'] },
    // },
 
    /* Test against mobile viewports. */
    // {
    //   name: 'Mobile Chrome',
    //   use: { ...devices['Pixel 5'] },
    // },
    // {
    //   name: 'Mobile Safari',
    //   use: { ...devices['iPhone 12'] },
    // },
 
    /* Test against branded browsers. */
    // {
    //   name: 'Microsoft Edge',
    //   use: { ...devices['Desktop Edge'], channel: 'msedge' },
    // },
    // {
    //   name: 'Google Chrome',
    //   use: { ...devices['Desktop Chrome'], channel: 'chrome' },
    // },
  ],
 
  /* Run your local dev server before starting the tests */
  // webServer: {
  //   command: 'npm run start',
  //   url: 'http://localhost:3000',
  //   reuseExistingServer: !process.env.CI,
  // },
});
 