/**
 * Page Object Model for Login Page
 * Contains locators and methods for Login page interactions
 */
import { WebDriver, By } from 'selenium-webdriver';
import { BasePage } from './BasePage';
export class LoginPage extends BasePage {
    // Locators using XPath from provided list
    private usernameField = By.xpath("//input[@id='user-name']");
    private passwordField = By.xpath("//input[@id='password']");
    private loginButton = By.xpath("//input[@id='login-button']");
    private loginCredentialsDiv = By.xpath("//div[@id='login_credentials']");
    private loginButtonContainer = By.xpath("//div[@id='login_button_container']");
    constructor(driver: WebDriver) {
        super(driver);
    }
    /**
     * Navigate to the login page
     * @param url - Base URL from test data using {{base_url}} placeholder
     */
    async navigateToLoginPage(url: string): Promise<void> {
        await this.navigateTo(url);
        console.log(`Navigated to login page: ${url}`);
    }
    /**
     * Verify that login page has loaded successfully
     */
    async verifyLoginPageDisplayed(): Promise<void> {
        await this.waitForElementVisible(this.usernameField);
        await this.waitForElementVisible(this.passwordField);
        await this.waitForElementVisible(this.loginButton);
        this.logVerification('Login page loads successfully');
    }
    /**
     * Enter username in the username field
     * @param username - Username to enter (from {{username}} placeholder)
     */
    async enterUsername(username: string): Promise<void> {
        await this.type(this.usernameField, username);
        console.log(`Entered username: ${username}`);
    }
    /**
     * Verify username field accepts input
     */
    async verifyUsernameFieldEnabled(): Promise<void> {
        const isEnabled = await this.isElementEnabled(this.usernameField);
        if (!isEnabled) {
            throw new Error('Username field is not enabled');
        }
        this.logVerification('Username field accepts input');
    }
    /**
     * Get the current value in username field
     */
    async getUsernameValue(): Promise<string | null> {
        return await this.getValue(this.usernameField);
    }
    /**
     * Enter password in the password field
     * @param password - Password to enter (from {{password}} placeholder)
     */
    async enterPassword(password: string): Promise<void> {
        await this.type(this.passwordField, password);
        const maskedPassword = '*'.repeat(password.length);
        console.log(`Entered password: ${maskedPassword}`);
    }
    /**
     * Verify password field accepts input and masks characters
     */
    async verifyPasswordFieldMasked(): Promise<void> {
        const fieldType = await this.getAttribute(this.passwordField, 'type');
        if (fieldType !== 'password') {
            throw new Error('Password field is not masked');
        }
        this.logVerification('Password field accepts input and masks characters');
    }
    /**
     * Click the Login button
     */
    async clickLoginButton(): Promise<void> {
        await this.click(this.loginButton);
        console.log('Clicked Login button');
    }
    /**
     * Verify login button is clickable
     */
    async verifyLoginButtonClickable(): Promise<void> {
        const isVisible = await this.isElementVisible(this.loginButton);
        const isEnabled = await this.isElementEnabled(this.loginButton);
        if (!isVisible || !isEnabled) {
            throw new Error('Login button is not clickable');
        }
        this.logVerification('Login button is clickable');
    }
    /**
     * Perform complete login action
     * @param username - Username for login (from {{username}} placeholder)
     * @param password - Password for login (from {{password}} placeholder)
     */
    async login(username: string, password: string): Promise<void> {
        await this.enterUsername(username);
        await this.enterPassword(password);
        await this.clickLoginButton();
        console.log('Login action completed');
    }
}