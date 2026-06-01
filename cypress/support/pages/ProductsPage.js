/**
 * Page Object Model for Products Page
 * Contains locators and methods for Products page interactions
 */
class ProductsPage {
    // Locators using CSS selectors
    elements = {
        productsPageTitle: () => cy.get('span').contains('Products'),
        addToCartSauceLabsBackpack: () => cy.get('#add-to-cart-sauce-labs-backpack'),
        removeSauceLabsBackpack: () => cy.get('#remove-sauce-labs-backpack'),
        shoppingCartContainer: () => cy.get('#shopping_cart_container'),
        cartBadge: () => cy.get('.shopping_cart_badge'),
        inventoryContainer: () => cy.get('#inventory_container')
    };
    /**
     * Wait for Products page to load
     */
    waitForProductsPageToLoad() {
        this.elements.productsPageTitle().should('be.visible');
        this.elements.inventoryContainer().should('be.visible');
        cy.log('Products page loaded successfully');
    }
    /**
     * Verify Products page loads with product listings
     */
    verifyProductsPageDisplayed() {
        this.elements.productsPageTitle().should('be.visible');
        this.elements.inventoryContainer().should('be.visible');
        cy.log('✓ Verified: Products page loads with product listings');
    }
    /**
     * Click 'Add to cart' button for Sauce Labs Backpack
     */
    clickAddToCartForSauceLabsBackpack() {
        this.elements.addToCartSauceLabsBackpack().click();
        cy.log("Clicked 'Add to cart' for Sauce Labs Backpack");
    }
    /**
     * Verify 'Add to cart' button changes to 'Remove' after clicking
     */
    verifyRemoveButtonDisplayed() {
        this.elements.removeSauceLabsBackpack().should('be.visible');
        cy.log("✓ Verified: 'Add to cart' button changes to 'Remove' after clicking");
    }
    /**
     * Get cart badge count
     * @returns {Cypress.Chainable<string>}
     */
    getCartBadgeCount() {
        return this.elements.cartBadge().invoke('text');
    }
    /**
     * Verify cart icon shows badge with expected count
     * @param {string} expectedCount - Expected badge count
     */
    verifyCartBadgeCount(expectedCount) {
        this.elements.cartBadge().should('have.text', expectedCount);
        cy.log(`✓ Verified: Cart icon shows badge with '${expectedCount}'`);
    }
    /**
     * Click on Cart icon to navigate to Cart page
     */
    clickCartIcon() {
        this.elements.shoppingCartContainer().click();
        cy.log('Clicked Cart icon');
    }
}
export default ProductsPage;