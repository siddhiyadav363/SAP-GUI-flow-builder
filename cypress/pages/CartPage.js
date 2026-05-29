/**
 * Finish page POM class
 */
/**
 * CartPage page object
 */
class CartPage {
  constructor() {
    this.cartContentsContainer = "//div[@id='cart_contents_container']";
    this.checkoutButton = "//button[@id='checkout']";
    // Specific remove or item references can be extended if needed
    this.cartItemTitle = "//div[@id='cart_contents_container']//a[@id='item_4_title_link' or contains(text(),'Sauce Labs Backpack')] | //a[@id='item_4_title_link']";
  }
  /**
   * assertCartHasProduct
   * @param {string} productName
   */
  assertCartHasProduct(productName) {
    // We use a flexible XPath that matches the provided item title link
    cy.xpath(this.cartContentsContainer).should('be.visible');
    cy.xpath(this.cartItemTitle).should('contain.text', productName);
  }
  /**
   * clickCheckout
   */
  clickCheckout() {
    cy.xpath(this.checkoutButton).should('be.visible').click();
  }
}
export default CartPage;