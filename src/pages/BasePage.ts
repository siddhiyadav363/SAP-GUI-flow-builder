/**
 * Base Page Object Model
 * Contains common methods and utilities for all page objects
 */
import { WebDriver, By, until, WebElement } from 'selenium-webdriver';
export class BasePage {
    protected driver: WebDriver;
    protected timeout: number = 10000;
    constructor(driver: WebDriver) {
        this.driver = driver;
    }
    /**
     * Navigate to a URL
     * @param url - URL to navigate to
     */
    async navigateTo(url: string): Promise<void> {
        await this.driver.get(url);
        console.log(`Navigated to: ${url}`);
    }
    /**
     * Wait for element to be visible
     * @param locator - Element locator
     * @param timeout - Wait timeout in milliseconds
     */
    async waitForElementVisible(locator: By, timeout: number = this.timeout): Promise<WebElement> {
        return await this.driver.wait(until.elementLocated(locator), timeout);
    }
    /**
     * Wait for element to be clickable
     * @param locator - Element locator
     * @param timeout - Wait timeout in milliseconds
     */
    async waitForElementClickable(locator: By, timeout: number = this.timeout): Promise<WebElement> {
        const element = await this.waitForElementVisible(locator, timeout);
        await this.driver.wait(until.elementIsEnabled(element), timeout);
        return element;
    }
    /**
     * Click on an element
     * @param locator - Element locator
     */
    async click(locator: By): Promise<void> {
        const element = await this.waitForElementClickable(locator);
        await element.click();
    }
    /**
     * Type text into an element
     * @param locator - Element locator
     * @param text - Text to type
     */
    async type(locator: By, text: string): Promise<void> {
        const element = await this.waitForElementVisible(locator);
        await element.clear();
        await element.sendKeys(text);
    }
    /**
     * Get text from an element
     * @param locator - Element locator
     */
    async getText(locator: By): Promise<string> {
        const element = await this.waitForElementVisible(locator);
        return await element.getText();
    }
    /**
     * Get attribute value from an element
     * @param locator - Element locator
     * @param attribute - Attribute name
     */
    async getAttribute(locator: By, attribute: string): Promise<string | null> {
        const element = await this.waitForElementVisible(locator);
        return await element.getAttribute(attribute);
    }
    /**
     * Check if element is visible
     * @param locator - Element locator
     */
    async isElementVisible(locator: By): Promise<boolean> {
        try {
            const element = await this.driver.findElement(locator);
            return await element.isDisplayed();
        } catch (error) {
            return false;
        }
    }
    /**
     * Check if element is enabled
     * @param locator - Element locator
     */
    async isElementEnabled(locator: By): Promise<boolean> {
        try {
            const element = await this.driver.findElement(locator);
            return await element.isEnabled();
        } catch (error) {
            return false;
        }
    }
    /**
     * Get value from input element
     * @param locator - Element locator
     */
    async getValue(locator: By): Promise<string | null> {
        return await this.getAttribute(locator, 'value');
    }
    /**
     * Log step information
     * @param stepNumber - Step number
     * @param description - Step description
     */
    logStep(stepNumber: number, description: string): void {
        console.log(`\nStep ${stepNumber}: ${description}`);
    }
    /**
     * Log verification information
     * @param description - Verification description
     */
    logVerification(description: string): void {
        console.log(`✓ Verified: ${description}`);
    }
}