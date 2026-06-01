/**
 * Page Object Model for Products Page
 * Contains locators and methods for Products page interactions
 */
import { WebDriver, By } from 'selenium-webdriver';
import { BasePage } from './BasePage';
export class ProductsPage extends BasePage {
    // Locators using XPath from provided list
    private productsPageTitle = By.xpath("//span[text()='Products']");
    private addToCartSauceLabsBackpack = By.xpath("//button[@id='add-to-cart-sauce-labs-backpack']");
    private removeSauceLabsBackpack = By.xpath("//button[@id='remove-sauce-labs-backpack']");
    private shoppingCartContainer = By.xpath("//div[@id='shopping_cart_container']");
    private cartBadge = By.xpath("//span[@class='shopping_cart_badge']");
    private inventoryContainer = By.xpath("//div[@id='inventory_container']");
    constructor(driver: WebDriver) {
        super(driver);
    }
    /**
     * Wait for Products page to load
     */
    async waitForProductsPageToLoad(): Promise<void> {
        await this.waitForElementVisible(this.productsPageTitle);
        await this.waitForElementVisible(this.inventoryContainer);
        console.log('Products page loaded successfully');
    }
    /**
     * Verify Products page loads with product listings
     */
    async verifyProductsPageDisplayed(): Promise<void> {
        const isTitleVisible = await this.isElementVisible(this.productsPageTitle);
        const isInventoryVisible = await this.isElementVisible(this.inventoryContainer);
        if (!isTitleVisible || !isInventoryVisible) {
            throw new Error('Products page not displayed correctly');
        }
        this.logVerification('Products page loads with product listings');
    }
    /**
     * Click 'Add to cart' button for Sauce Labs Backpack
     */
    async clickAddToCartForSauceLabsBackpack(): Promise<void> {
        await this.click(this.addToCartSauceLabsBackpack);
        console.log("Clicked 'Add to cart' for Sauce Labs Backpack");
    }
    /**
     * Verify 'Add to cart' button changes to 'Remove' after clicking
     */
    async verifyRemoveButtonDisplayed(): Promise<void> {
        await this.waitForElementVisible(this.removeSauceLabsBackpack);
        const isVisible = await this.isElementVisible(this.removeSauceLabsBackpack);
        if (!isVisible) {
            throw new Error("'Remove' button not displayed");
        }
        this.logVerification("'Add to cart' button changes to 'Remove' after clicking");
    }
    /**
     * Get cart badge count
     */
    async getCartBadgeCount(): Promise<string> {
        await this.waitForElementVisible(this.cartBadge);
        return await this.getText(this.cartBadge);
    }
    /**
     * Verify cart icon shows badge with expected count
     * @param expectedCount - Expected badge count
     */
    async verifyCartBadgeCount(expectedCount: string): Promise<void> {
        const actualCount = await this.getCartBadgeCount();
        if (actualCount !== expectedCount) {
            throw new Error(`Cart badge count mismatch. Expected: ${expectedCount}, Actual: ${actualCount}`);
        }
        console.log(`Cart badge count - Expected: ${expectedCount}, Actual: ${actualCount}`);
        this.logVerification(`Cart icon shows badge with '${expectedCount}'`);
    }
    /**
     * Click on Cart icon to navigate to Cart page
     */
    async clickCartIcon(): Promise<void> {
        await this.click(this.shoppingCartContainer);
        console.log('Clicked Cart icon');
    }
}