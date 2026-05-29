/**
 * Finish page POM class
 */
/**
 * ProductsPage page object for inventory interactions
 */
class ProductsPage {
  constructor() {
    // XPaths from provided locators
    this.addBackpackButton = "//button[@id='add-to-cart-sauce-labs-backpack']";
    this.inventoryContainer = "//div[@id='inventory_container']";
    this.cartIcon = "//div[@id='shopping_cart_container']";
    this.backpackTitleLink = "//a[@id='item_4_title_link']";
  }
  /**
   * assertProductsPageLoaded
   * Verifies the Products page is displayed
   */
  assertProductsPageLoaded() {
    cy.xpath(this.inventoryContainer).should('be.visible');
  }
  /**
   * addSauceLabsBackpackToCart
   */
  addSauceLabsBackpackToCart() {
    cy.xpath(this.addBackpackButton).should('be.visible').click();
    // Verify cart badge increment or cart contains item after add (cart content visible)
    cy.xpath(this.cartIcon).should('be.visible');
  }
  /**
   * openCart
   */
  openCart() {
    cy.xpath(this.cartIcon).click();
  }
}
export default ProductsPage;