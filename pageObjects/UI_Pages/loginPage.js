const { excuteSteps } = require("../../utilities/actions");
const { test,expect } = require("@playwright/test");
exports.LoginPage = class LoginPage {
  constructor(test, page) {
    this.test = test;
    this.page = page;
    this.signInBtn=page.locator("//button/*[contains(text(),'Sign in')]");
    this.googleSignIn=page.locator("//button[contains(text(),'Continue with Google')]");
    this.googleNextBtn=page.locator("//span[text()='Next']");
    this.googleEmail=page.locator("//input[@aria-label='Email or phone']");
    this.googlePassword=page.locator("//input[@aria-label='Enter your password']");
    this.microsoftSignIn=page.locator("//button[contains(.,'Sign in with')]")
    this.email=page.locator("//input[@type='email']");
    this.nextBtn=page.locator("//input[@value='Next']");
    this.password=page.locator("//input[@type='password']");
    this.microsoftsignInButton=page.locator("//input[@value='Sign in']");
    this.cashPostingHeader=page.locator("//h3[contains(text(),'Cash Posting')]");
    this.yesBtn=page.locator("//input[@value='Yes']");
    this.dontShowAgainBtn=page.locator("//input[@name='DontShowAgain']");
    this.continueBtn=page.locator("//button[text()='Continue']");
    this.loginBtn=page.locator("(//button[text()='Log in'])[1]");
  }
  launchingApplication = async (baseUrl) => {
    await excuteSteps(
      this.test,
      await this.page,
      "navigate",
      `Navigate to AI flow`,
      baseUrl
    );
  };
  fillingEmail = async (email) => {
    await excuteSteps(
      this.test,
      this.email,
      "fill",
      `Enter username in username field`,
      email
    );
  };
  fillingPassword = async (pwd) => {
    await excuteSteps(
      this.test,
      this.password,
      "fill",
      `Entering password in password field`,
      pwd
    );
  };

  
  fillinggoogleEmail = async (email) => {
    await excuteSteps(
      this.test,
      this.googleEmail,
      "fill",
      `Enter username in username field`,
      email
    );
  };

  fillinggooglePassword = async (pwd) => {
    await excuteSteps(
      this.test,
      this.googlePassword,
      "fill",
      `Entering password in password field`,
      pwd
    );
  };

  clickOnGoogleNextBtn=async()=>{
    await excuteSteps(this.test,this.googleNextBtn,"click",`Clicking on google next button`);
  };

  clickOnGoogleSignInBtn=async()=>{
    await excuteSteps(this.test,this.googleSignIn,"click",`Clicking on google sign in button`);
  }

  clickOnDontShowAgainBtn=async()=>{
    await excuteSteps(this.test,this.dontShowAgainBtn,"click",`Clicking on dont show again button`);
  };

  clickOnContinueBtn=async()=>{
    await excuteSteps(test,this.continueBtn,"click",`Clicking on Continue button`);
  };

  clickOnSignBtn=async()=>{
    await excuteSteps(this.test,this.signInBtn,"click",`Clicking on sign in button`);
  };

  clickOnLoginBtn=async()=>{
    await excuteSteps(test,this.loginBtn,"click",`Clicking on Login in Button`);
  };
  
  clickOnNextBtn = async()=>{
    await excuteSteps(this.test,this.nextBtn,"click",`Clicking on next button`);
  };

  clickOnMicrosoftLogInBtn=async()=>{
    await excuteSteps(this.test,this.microsoftSignIn,"click",`Clicking on microsoft log in button`);
  };
  clickOnMicrosoftSignInBtn =async()=>{
    await excuteSteps(this.test,this.microsoftsignInButton,"click",`Clicking on microsoft sign in`)
  };
  clickOnYesBtn=async()=>{
    await excuteSteps(this.test,this.yesBtn,"click",`Clicking on yes button for log in`);
  };

  loginToLovable=async(email,pwd)=>{
    await this.test.step("The page is loading, please wait", async () => {
      await this.page.waitForTimeout(parseInt(process.env.smallWait));
    });
    
    await this.fillingEmail(email);
    await this.clickOnContinueBtn();
    await this.fillingPassword(pwd);
    await this.clickOnLoginBtn();

  };

  loginWithValidCredentials = async (email, pwd) => {
    await this.test.step("The page is loading, please wait", async () => {
      await this.page.waitForTimeout(parseInt(process.env.smallWait));
    });
    //await this.clickOnSignBtn();
    await this.clickOnMicrosoftLogInBtn();
    await this.fillingEmail(email);
    await this.test.step("The page is loading, please wait", async () => {
      await this.page.waitForTimeout(parseInt(process.env.smallWait));
    });
    await this.test.step("The page is loading, please wait", async () => {
      await this.page.waitForTimeout(parseInt(process.env.smallWait));
    });
    await this.clickOnNextBtn();
    await this.test.step("The page is loading, please wait", async () => {
      await this.page.waitForTimeout(parseInt(process.env.smallWait));
    });
    await this.test.step("The page is loading, please wait", async () => {
      await this.page.waitForTimeout(parseInt(process.env.smallWait));
    });
    await this.fillingPassword(pwd);
    await this.test.step("The page is loading, please wait", async () => {
      await this.page.waitForTimeout(parseInt(process.env.smallWait));
    });

    await this.clickOnMicrosoftSignInBtn(); 
    await this.clickOnDontShowAgainBtn();
    await this.clickOnYesBtn();
    await this.cashPostingHeader.waitFor({state: 'visible'});
  };




};
