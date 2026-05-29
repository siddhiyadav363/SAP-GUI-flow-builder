/**
 * Finish page POM class
 */
/**
 * CheckoutOverviewPage page object
 */
class CheckoutOverviewPage {
  constructor() {
    this.checkoutSummaryContainer = "//div[@id='checkout_summary_container']";
    this.finishButton = "//button[@id='finish']";
  }
  /**
   * assertOverviewPageLoaded
   * Verifies the Checkout: Overview screen is displayed
   */
  assertOverviewPageLoaded() {
    cy.xpath(this.checkoutSummaryContainer).should('be.visible');
    // Also ensure finish button exists as extra verification
    cy.xpath(this.finishButton).should('be.visible');
  }
}
export default CheckoutOverviewPage;