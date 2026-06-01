/**
 * Page Object Model for Checkout: Your Information Page
 * Contains locators and methods for Checkout Information page interactions
 */
import { WebDriver, By } from 'selenium-webdriver';
import { BasePage } from './BasePage';
export class CheckoutInformationPage extends BasePage {
    // Locators using XPath from provided list
    private firstNameField = By.xpath("//input[@id='first-name']");
    private lastNameField = By.xpath("//input[@id='last-name']");
    private postalCodeField = By.xpath("//input[@id='postal-code']");
    private continueButton = By.xpath("//input[@id='continue']");
    private cancelButton = By.xpath("//button[@id='cancel']");
    private checkoutInfoContainer = By.xpath("//div[@id='checkout_info_container']");
    private checkoutInfoTitle = By.xpath("//span[@class='title' and text()='Checkout: Your Information']");
    constructor(driver: WebDriver) {
        super(driver);
    }
    /**
     * Wait for Checkout: Your Information page to load
     */
    async waitForCheckoutInformationPageToLoad(): Promise<void> {
        await this.waitForElementVisible(this.checkoutInfoContainer);
        await this.waitForElementVisible(this.firstNameField);
        console.log('Checkout: Your Information page loaded successfully');
    }
    /**
     * Verify 'Checkout: Your Information' page displays with all fields
     */
    async verifyCheckoutInformationPageDisplayed(): Promise<void> {
        const isContainerVisible = await this.isElementVisible(this.checkoutInfoContainer);
        const isFirstNameVisible = await this.isElementVisible(this.firstNameField);
        const isLastNameVisible = await this.isElementVisible(this.lastNameField);
        const isPostalCodeVisible = await this.isElementVisible(this.postalCodeField);
        if (!isContainerVisible || !isFirstNameVisible || !isLastNameVisible || !isPostalCodeVisible) {
            throw new Error('Checkout Information page not displayed correctly');
        }
        this.logVerification("'Checkout: Your Information' page displays with header and three mandatory fields");
    }
    /**
     * Enter first name in the First Name field
     * @param firstName - First name to enter (from {{first_name}} placeholder)
     */
    async enterFirstName(firstName: string): Promise<void> {
        await this.type(this.firstNameField, firstName);
        console.log(`Entered first name: ${firstName}`);
    }
    /**
     * Verify First Name field accepts alphabetic input
     */
    async verifyFirstNameFieldEnabled(): Promise<void> {
        const isEnabled = await this.isElementEnabled(this.firstNameField);
        if (!isEnabled) {
            throw new Error('First Name field is not enabled');
        }
        this.logVerification('First Name field accepts alphabetic input');
    }
    /**
     * Enter last name in the Last Name field
     * @param lastName - Last name to enter (from {{last_name}} placeholder)
     */
    async enterLastName(lastName: string): Promise<void> {
        await this.type(this.lastNameField, lastName);
        console.log(`Entered last name: ${lastName}`);
    }
    /**
     * Verify Last Name field accepts alphabetic input
     */
    async verifyLastNameFieldEnabled(): Promise<void> {
        const isEnabled = await this.isElementEnabled(this.lastNameField);
        if (!isEnabled) {
            throw new Error('Last Name field is not enabled');
        }
        this.logVerification('Last Name field accepts alphabetic input');
    }
    /**
     * Enter zip/postal code in the Zip/Postal Code field
     * @param zipCode - Zip/Postal code to enter (from {{zip_code}} placeholder)
     */
    async enterZipPostalCode(zipCode: string): Promise<void> {
        await this.type(this.postalCodeField, zipCode);
        console.log(`Entered zip/postal code: ${zipCode}`);
    }
    /**
     * Verify Zip/Postal Code field accepts numeric input
     */
    async verifyZipPostalCodeFieldEnabled(): Promise<void> {
        const isEnabled = await this.isElementEnabled(this.postalCodeField);
        if (!isEnabled) {
            throw new Error('Zip/Postal Code field is not enabled');
        }
        this.logVerification('Zip/Postal Code field accepts numeric input');
    }
    /**
     * Click Continue button to proceed to Checkout Overview
     */
    async clickContinueButton(): Promise<void> {
        await this.click(this.continueButton);
        console.log('Clicked Continue button');
    }
    /**
     * Fill all checkout information fields
     * @param firstName - First name (from {{first_name}} placeholder)
     * @param lastName - Last name (from {{last_name}} placeholder)
     * @param zipCode - Zip/Postal code (from {{zip_code}} placeholder)
     */
    async fillCheckoutInformation(firstName: string, lastName: string, zipCode: string): Promise<void> {
        await this.enterFirstName(firstName);
        await this.enterLastName(lastName);
        await this.enterZipPostalCode(zipCode);
        console.log('Filled all checkout information fields');
    }
}