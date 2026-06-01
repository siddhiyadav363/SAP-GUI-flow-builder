/**
 * Checkout Complete Page POM class
 * Contains locators and methods for Checkout Complete/Finish page interactions
 */
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
namespace SwagLabsAutomation.POM
{
    /// <summary>
    /// Page Object Model for Checkout Complete (Finish) Page
    /// </summary>
    public class CheckoutCompletePagePOM
    {
        private readonly IWebDriver _driver;
        private readonly WebDriverWait _wait;
        // Locators using XPath from provided list
        private readonly By checkoutCompleteContainer = By.XPath("//div[@id='checkout_complete_container']");
        private readonly By thankYouMessage = By.XPath("//h2");
        private readonly By backHomeButton = By.XPath("//button[@id='back-to-products']");
        private readonly By ponyExpressImage = By.XPath("//img[@class='pony_express']");
        /// <summary>
        /// Constructor for CheckoutCompletePagePOM
        /// </summary>
        /// <param name="driver">WebDriver instance</param>
        public CheckoutCompletePagePOM(IWebDriver driver)
        {
            _driver = driver;
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
        }
        /// <summary>
        /// Wait for Finish page to load
        /// </summary>
        public void WaitForFinishPageToLoad()
        {
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(checkoutCompleteContainer));
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(thankYouMessage));
        }
        /// <summary>
        /// Verify Finish page loads successfully
        /// </summary>
        /// <returns>True if Finish page is displayed</returns>
        public bool IsFinishPageDisplayed()
        {
            try
            {
                return _driver.FindElement(checkoutCompleteContainer).Displayed;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        /// <summary>
        /// Get the thank you message text
        /// </summary>
        /// <returns>Thank you message text</returns>
        public string GetThankYouMessage()
        {
            IWebElement messageElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(thankYouMessage));
            return messageElement.Text;
        }
        /// <summary>
        /// Verify 'Thank you for your order!' message displays
        /// </summary>
        /// <returns>True if success message is displayed</returns>
        public bool IsThankYouMessageDisplayed()
        {
            try
            {
                string message = GetThankYouMessage();
                return message.Contains("Thank you for your order!");
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }
        /// <summary>
        /// Verify Pony Express Sauce Labs logo displays
        /// </summary>
        /// <returns>True if logo is displayed</returns>
        public bool IsPonyExpressLogoDisplayed()
        {
            try
            {
                IWebElement logoElement = _driver.FindElement(ponyExpressImage);
                return logoElement.Displayed;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        /// <summary>
        /// Verify success message and logo display
        /// </summary>
        /// <returns>True if both success message and logo are displayed</returns>
        public bool IsOrderCompletionConfirmed()
        {
            return IsThankYouMessageDisplayed() && IsPonyExpressLogoDisplayed();
        }
    }
}