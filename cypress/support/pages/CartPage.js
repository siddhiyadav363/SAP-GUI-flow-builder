/**
 * Page Object Model for Cart Page
 * Contains locators and methods for Cart page interactions
 */
class CartPage {
    // Locators using CSS selectors
    elements = {
        cartContentsContainer: () => cy.get('#cart_contents_container'),
        checkoutButton: () => cy.get('#checkout'),
        continueShoppingButton: () => cy.get('#continue-shopping'),
        removeSauceLabsBackpack: () => cy.get('#remove-sauce-labs-backpack'),
        inventoryItemName: () => cy.get('.inventory_item_name'),
        cartQuantity: () => cy.get('.cart_quantity')
    };
    /**
     * Wait for Cart page to load
     */
    waitForCartPageToLoad() {
        this.elements.cartContentsContainer().should('be.visible');
        cy.log('Cart page loaded successfully');
    }
    /**
     * Verify Cart page displays with correct product name
     * @param {string} productName - Expected product name (from {{product_name}} placeholder)
     */
    verifyProductDisplayedInCart(productName) {
        this.elements.inventoryItemName().should('contain.text', productName);
        cy.log(`Product in cart verified: ${productName}`);
    }
    /**
     * Verify product appears in cart with quantity 1
     * @param {string} productName - Expected product name (from {{product_name}} placeholder)
     */
    verifyProductWithQuantityInCart(productName) {
        cy.get('.cart_quantity').contains('1').should('be.visible');
        this.verifyProductDisplayedInCart(productName);
        cy.log(`✓ Verified: Cart page displays with correct product name '${productName}' and quantity 1`);
    }
    /**
     * Verify Checkout button is visible and clickable
     */
    verifyCheckoutButtonVisibleAndClickable() {
        this.elements.checkoutButton().should('be.visible').and('be.enabled');
        cy.log('✓ Verified: Checkout button is visible and clickable');
    }
    /**
     * Click Checkout button to proceed to checkout
     */
    clickCheckoutButton() {
        this.elements.checkoutButton().click();
        cy.log('Clicked Checkout button');
    }
}
export default CartPage;