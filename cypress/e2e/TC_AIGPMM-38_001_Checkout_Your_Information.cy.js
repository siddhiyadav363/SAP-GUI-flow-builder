/**
 * Finish page POM class
 */
/// <reference types="cypress" />
// Ensure the cypress-xpath plugin is available in your Cypress setup (install & include in support/index.js)
// This test uses data-driven placeholders per project requirements: {{base_url}}, {{username}}, {{password}}, {{firstName}}, {{lastName}}, {{postalCode}}
require('cypress-xpath');
import LoginPage from '../pages/LoginPage';
import ProductsPage from '../pages/ProductsPage';
import CartPage from '../pages/CartPage';
import CheckoutYourInformationPage from '../pages/CheckoutYourInformationPage';
import CheckoutOverviewPage from '../pages/CheckoutOverviewPage';
describe('TC_AIGPMM-38_001 Checkout: Your Information Flow', () => {
  const loginPage = new LoginPage();
  const productsPage = new ProductsPage();
  const cartPage = new CartPage();
  const checkoutInfo = new CheckoutYourInformationPage();
  const checkoutOverview = new CheckoutOverviewPage();
  /**
   * testCheckoutOverviewFlow
   *
   * Verifies the user can login, add Sauce Labs Backpack to cart,
   * proceed to Checkout: Your Information, fill required fields and continue to Overview.
   */
  it('testCheckoutOverviewFlow', () => {
    // Navigate to application (data-driven placeholder)
    cy.visit('{{base_url}}');
    // Login
    loginPage.enterUsername('{{username}}');
    loginPage.enterPassword('{{password}}');
    loginPage.clickLogin();
    // Verify Products page loaded by checking inventory container visible
    productsPage.assertProductsPageLoaded();
    // Add product and navigate to cart
    productsPage.addSauceLabsBackpackToCart();
    productsPage.openCart();
    // Cart: verify product present and click Checkout
    cartPage.assertCartHasProduct('Sauce Labs Backpack');
    cartPage.clickCheckout();
    // Checkout: Your Information - fill fields using placeholders
    checkoutInfo.enterFirstName('{{firstName}}');
    checkoutInfo.enterLastName('{{lastName}}');
    checkoutInfo.enterPostalCode('{{postalCode}}');
    checkoutInfo.clickContinue();
    // Wait/assert for Checkout: Overview to load
    checkoutOverview.assertOverviewPageLoaded();
    // Final assertion: URL includes checkout-step-two.html (navigation to Overview)
    cy.url().should('include', 'checkout-step-two.html');
  });
});