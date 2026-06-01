/**
 * Products Page POM class
 * Contains locators and methods for Products page interactions
 */
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
namespace SwagLabsAutomation.POM
{
    /// <summary>
    /// Page Object Model for Products Page
    /// </summary>
    public class ProductsPagePOM
    {
        private readonly IWebDriver _driver;
        private readonly WebDriverWait _wait;
        // Locators using XPath from provided list
        private readonly By productsPageTitle = By.XPath("//span[text()='Products']");
        private readonly By addToCartSauceLabsBackpack = By.XPath("//button[@id='add-to-cart-sauce-labs-backpack']");
        private readonly By removeSauceLabsBackpack = By.XPath("//button[@id='remove-sauce-labs-backpack']");
        private readonly By shoppingCartContainer = By.XPath("//div[@id='shopping_cart_container']");
        private readonly By cartBadge = By.XPath("//span[@class='shopping_cart_badge']");
        private readonly By inventoryContainer = By.XPath("//div[@id='inventory_container']");
        /// <summary>
        /// Constructor for ProductsPagePOM
        /// </summary>
        /// <param name="driver">WebDriver instance</param>
        public ProductsPagePOM(IWebDriver driver)
        {
            _driver = driver;
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
        }
        /// <summary>
        /// Wait for Products page to load
        /// </summary>
        public void WaitForProductsPageToLoad()
        {
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(productsPageTitle));
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(inventoryContainer));
        }
        /// <summary>
        /// Verify Products page loads with product listings
        /// </summary>
        /// <returns>True if Products page is displayed</returns>
        public bool IsProductsPageDisplayed()
        {
            try
            {
                return _driver.FindElement(productsPageTitle).Displayed &&
                       _driver.FindElement(inventoryContainer).Displayed;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        /// <summary>
        /// Click 'Add to cart' button for Sauce Labs Backpack
        /// </summary>
        public void ClickAddToCartForSauceLabsBackpack()
        {
            IWebElement addToCartBtn = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(addToCartSauceLabsBackpack));
            addToCartBtn.Click();
        }
        /// <summary>
        /// Verify 'Add to cart' button changes to 'Remove' after clicking
        /// </summary>
        /// <returns>True if Remove button is displayed</returns>
        public bool IsRemoveButtonDisplayedForSauceLabsBackpack()
        {
            try
            {
                _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(removeSauceLabsBackpack));
                return _driver.FindElement(removeSauceLabsBackpack).Displayed;
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }
        /// <summary>
        /// Get cart badge count
        /// </summary>
        /// <returns>Cart badge count as string</returns>
        public string GetCartBadgeCount()
        {
            try
            {
                IWebElement badge = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(cartBadge));
                return badge.Text;
            }
            catch (WebDriverTimeoutException)
            {
                return "0";
            }
        }
        /// <summary>
        /// Verify cart icon shows badge with expected count
        /// </summary>
        /// <param name="expectedCount">Expected badge count</param>
        /// <returns>True if badge matches expected count</returns>
        public bool IsCartBadgeCountCorrect(string expectedCount)
        {
            string actualCount = GetCartBadgeCount();
            return actualCount.Equals(expectedCount);
        }
        /// <summary>
        /// Click on Cart icon to navigate to Cart page
        /// </summary>
        public void ClickCartIcon()
        {
            IWebElement cartIcon = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(shoppingCartContainer));
            cartIcon.Click();
        }
    }
}