/**
 * Checkout Overview Page POM class
 * Contains locators and methods for Checkout: Overview page interactions
 */
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
namespace SwagLabsAutomation.POM
{
    /// <summary>
    /// Page Object Model for Checkout: Overview Page
    /// </summary>
    public class CheckoutOverviewPagePOM
    {
        private readonly IWebDriver _driver;
        private readonly WebDriverWait _wait;
        // Locators using XPath from provided list
        private readonly By checkoutSummaryContainer = By.XPath("//div[@id='checkout_summary_container']");
        private readonly By finishButton = By.XPath("//button[@id='finish']");
        private readonly By cancelButton = By.XPath("//button[@id='cancel']");
        private readonly By checkoutOverviewTitle = By.XPath("//span[@class='title' and text()='Checkout: Overview']");
        private readonly By hamburgerMenuButton = By.XPath("//button[@id='react-burger-menu-btn']");
        private readonly By appLogo = By.XPath("//div[@class='app_logo']");
        private readonly By shoppingCartContainer = By.XPath("//div[@id='shopping_cart_container']");
        private readonly By cartItem = By.XPath("//div[@class='cart_item']");
        private readonly By inventoryItemName = By.XPath("//div[@class='inventory_item_name']");
        private readonly By summarySubtotalLabel = By.XPath("//div[@class='summary_subtotal_label']");
        private readonly By summaryTaxLabel = By.XPath("//div[@class='summary_tax_label']");
        private readonly By summaryTotalLabel = By.XPath("//div[@class='summary_total_label']");
        private readonly By paymentInfoLabel = By.XPath("//div[@class='summary_info' and contains(., 'Payment Information')]");
        private readonly By shippingInfoLabel = By.XPath("//div[@class='summary_info' and contains(., 'Shipping Information')]");
        /// <summary>
        /// Constructor for CheckoutOverviewPagePOM
        /// </summary>
        /// <param name="driver">WebDriver instance</param>
        public CheckoutOverviewPagePOM(IWebDriver driver)
        {
            _driver = driver;
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
        }
        /// <summary>
        /// Wait for Checkout: Overview page to load
        /// </summary>
        public void WaitForCheckoutOverviewPageToLoad()
        {
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(checkoutSummaryContainer));
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(finishButton));
        }
        /// <summary>
        /// Verify 'Checkout: Overview' page displays with correct header
        /// </summary>
        /// <returns>True if page header with hamburger menu, SWAGLABS logo, and cart icon is displayed</returns>
        public bool IsCheckoutOverviewHeaderDisplayed()
        {
            try
            {
                return _driver.FindElement(hamburgerMenuButton).Displayed &&
                       _driver.FindElement(appLogo).Displayed &&
                       _driver.FindElement(shoppingCartContainer).Displayed;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        /// <summary>
        /// Verify product details display correctly in overview
        /// </summary>
        /// <param name="productName">Expected product name</param>
        /// <returns>True if product is displayed</returns>
        public bool IsProductDisplayedInOverview(string productName)
        {
            try
            {
                _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(cartItem));
                string actualProductName = _driver.FindElement(inventoryItemName).Text;
                return actualProductName.Contains(productName);
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }
        /// <summary>
        /// Verify product table shows quantity and description correctly
        /// </summary>
        /// <param name="productName">Expected product name</param>
        /// <returns>True if product with quantity is displayed</returns>
        public bool IsProductTableDisplayedCorrectly(string productName)
        {
            try
            {
                By cartQuantity = By.XPath("//div[@class='cart_quantity' and text()='1']");
                _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(cartQuantity));
                return IsProductDisplayedInOverview(productName) && 
                       _driver.FindElement(cartQuantity).Displayed;
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }
        /// <summary>
        /// Verify Payment Information and Shipping Information sections display
        /// </summary>
        /// <returns>True if both sections are displayed</returns>
        public bool ArePaymentAndShippingInfoDisplayed()
        {
            try
            {
                return _driver.FindElement(paymentInfoLabel).Displayed &&
                       _driver.FindElement(shippingInfoLabel).Displayed;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        /// <summary>
        /// Verify Item Total, Tax, and Total amounts are calculated and displayed
        /// </summary>
        /// <returns>True if all price elements are displayed</returns>
        public bool ArePriceCalculationsDisplayed()
        {
            try
            {
                return _driver.FindElement(summarySubtotalLabel).Displayed &&
                       _driver.FindElement(summaryTaxLabel).Displayed &&
                       _driver.FindElement(summaryTotalLabel).Displayed;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        /// <summary>
        /// Get Item Total value
        /// </summary>
        /// <returns>Item Total as string</returns>
        public string GetItemTotal()
        {
            return _driver.FindElement(summarySubtotalLabel).Text;
        }
        /// <summary>
        /// Get Tax value
        /// </summary>
        /// <returns>Tax as string</returns>
        public string GetTax()
        {
            return _driver.FindElement(summaryTaxLabel).Text;
        }
        /// <summary>
        /// Get Total value
        /// </summary>
        /// <returns>Total as string</returns>
        public string GetTotal()
        {
            return _driver.FindElement(summaryTotalLabel).Text;
        }
        /// <summary>
        /// Click Finish button to complete the order
        /// </summary>
        public void ClickFinishButton()
        {
            IWebElement finishBtn = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(finishButton));
            finishBtn.Click();
        }
    }
}