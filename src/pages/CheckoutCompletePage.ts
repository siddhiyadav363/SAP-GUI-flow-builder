/**
 * Page Object Model for Checkout Complete (Finish) Page
 * Contains locators and methods for Checkout Complete page interactions
 */
import { WebDriver, By } from 'selenium-webdriver';
import { BasePage } from './BasePage';
export class CheckoutCompletePage extends BasePage {
    // Locators using XPath from provided list
    private checkoutCompleteContainer = By.xpath("//div[@id='checkout_complete_container']");
    private thankYouMessage = By.xpath("//h2");
    private backHomeButton = By.xpath("//button[@id='back-to-products']");
    private ponyExpressImage = By.xpath("//img[@class='pony_express']");
    constructor(driver: WebDriver) {
        super(driver);
    }
    /**
     * Wait for Finish page to load
     */
    async waitForFinishPageToLoad(): Promise<void> {
        await this.waitForElementVisible(this.checkoutCompleteContainer);
        await this.waitForElementVisible(this.thankYouMessage);
        console.log('Finish page loaded successfully');
    }
    /**
     * Verify Finish page loads successfully
     */
    async verifyFinishPageDisplayed(): Promise<void> {
        const isVisible = await this.isElementVisible(this.checkoutCompleteContainer);
        if (!isVisible) {
            throw new Error('Finish page not displayed');
        }
        this.logVerification('Finish page loads successfully');
    }
    /**
     * Get the thank you message text
     */
    async getThankYouMessage(): Promise<string> {
        await this.waitForElementVisible(this.thankYouMessage);
        const message = await this.getText(this.thankYouMessage);
        console.log(`Thank you message: ${message}`);
        return message;
    }
    /**
     * Verify 'Thank you for your order!' message displays
     */
    async verifyThankYouMessageDisplayed(): Promise<void> {
        const message = await this.getThankYouMessage();
        if (!message.includes('Thank you for your order!')) {
            throw new Error(`Expected message not found. Actual: ${message}`);
        }
        this.logVerification("'Thank you for your order!' message displays");
    }
    /**
     * Verify Pony Express Sauce Labs logo displays
     */
    async verifyPonyExpressLogoDisplayed(): Promise<void> {
        const isVisible = await this.isElementVisible(this.ponyExpressImage);
        if (!isVisible) {
            throw new Error('Pony Express logo not displayed');
        }
        this.logVerification('Pony Express Sauce Labs logo displays');
    }
    /**
     * Verify success message and logo display
     */
    async verifyOrderCompletionConfirmed(): Promise<void> {
        await this.verifyThankYouMessageDisplayed();
        await this.verifyPonyExpressLogoDisplayed();
        console.log('Order completion confirmed - Message and Logo displayed');
    }
}