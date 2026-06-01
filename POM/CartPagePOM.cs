/**
 * Cart Page POM class
 * Contains locators and methods for Cart page interactions
 */
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
namespace SwagLabsAutomation.POM
{
    /// <summary>
    /// Page Object Model for Cart Page
    /// </summary>
    public class CartPagePOM
    {
        private readonly IWebDriver _driver;
        private readonly WebDriverWait _wait;
        // Locators using XPath from provided list
        private readonly By cartContentsContainer = By.XPath("//div[@id='cart_contents_container']");
        private readonly By checkoutButton = By.XPath("//button[@id='checkout']");
        private readonly By continueShoppingButton = By.XPath("//button[@id='continue-shopping']");
        private readonly By removeSauceLabsBackpack = By.XPath("//button[@id='remove-sauce-labs-backpack']");
        private readonly By cartItemLabel = By.XPath("//div[@class='cart_item_label']");
        private readonly By inventoryItemName = By.XPath("//div[@class='inventory_item_name']");
        /// <summary>
        /// Constructor for CartPagePOM
        /// </summary>
        /// <param name="driver">WebDriver instance</param>
        public CartPagePOM(IWebDriver driver)
        {
            _driver = driver;
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
        }
        /// <summary>
        /// Wait for Cart page to load
        /// </summary>
        public void WaitForCartPageToLoad()
        {
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(cartContentsContainer));
        }
        /// <summary>
        /// Verify Cart page displays with correct product name and quantity
        /// </summary>
        /// <param name="productName">Expected product name</param>
        /// <returns>True if product is displayed in cart</returns>
        public bool IsProductDisplayedInCart(string productName)
        {
            try
            {
                _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(inventoryItemName));
                string actualProductName = _driver.FindElement(inventoryItemName).Text;
                return actualProductName.Contains(productName);
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }
        /// <summary>
        /// Verify product appears in cart with quantity 1
        /// </summary>
        /// <param name="productName">Expected product name</param>
        /// <returns>True if product with quantity 1 is in cart</returns>
        public bool VerifyProductWithQuantityInCart(string productName)
        {
            try
            {
                By cartQuantity = By.XPath("//div[@class='cart_quantity' and text()='1']");
                _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(cartQuantity));
                return IsProductDisplayedInCart(productName) && _driver.FindElement(cartQuantity).Displayed;
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }
        /// <summary>
        /// Verify Checkout button is visible and clickable
        /// </summary>
        /// <returns>True if Checkout button is enabled and displayed</returns>
        public bool IsCheckoutButtonVisibleAndClickable()
        {
            try
            {
                IWebElement checkoutBtn = _driver.FindElement(checkoutButton);
                return checkoutBtn.Displayed && checkoutBtn.Enabled;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        /// <summary>
        /// Click Checkout button to proceed to checkout
        /// </summary>
        public void ClickCheckoutButton()
        {
            IWebElement checkoutBtn = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(checkoutButton));
            checkoutBtn.Click();
        }
    }
}