/**
 * Checkout Information Page POM class
 * Contains locators and methods for Checkout: Your Information page interactions
 */
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
namespace SwagLabsAutomation.POM
{
    /// <summary>
    /// Page Object Model for Checkout: Your Information Page
    /// </summary>
    public class CheckoutInformationPagePOM
    {
        private readonly IWebDriver _driver;
        private readonly WebDriverWait _wait;
        // Locators using XPath from provided list
        private readonly By firstNameField = By.XPath("//input[@id='first-name']");
        private readonly By lastNameField = By.XPath("//input[@id='last-name']");
        private readonly By postalCodeField = By.XPath("//input[@id='postal-code']");
        private readonly By continueButton = By.XPath("//input[@id='continue']");
        private readonly By cancelButton = By.XPath("//button[@id='cancel']");
        private readonly By checkoutInfoContainer = By.XPath("//div[@id='checkout_info_container']");
        private readonly By checkoutInfoTitle = By.XPath("//span[@class='title' and text()='Checkout: Your Information']");
        /// <summary>
        /// Constructor for CheckoutInformationPagePOM
        /// </summary>
        /// <param name="driver">WebDriver instance</param>
        public CheckoutInformationPagePOM(IWebDriver driver)
        {
            _driver = driver;
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
        }
        /// <summary>
        /// Wait for Checkout: Your Information page to load
        /// </summary>
        public void WaitForCheckoutInformationPageToLoad()
        {
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(checkoutInfoContainer));
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(firstNameField));
        }
        /// <summary>
        /// Verify 'Checkout: Your Information' page displays with header and three mandatory fields
        /// </summary>
        /// <returns>True if page is displayed with all required fields</returns>
        public bool IsCheckoutInformationPageDisplayed()
        {
            try
            {
                return _driver.FindElement(checkoutInfoContainer).Displayed &&
                       _driver.FindElement(firstNameField).Displayed &&
                       _driver.FindElement(lastNameField).Displayed &&
                       _driver.FindElement(postalCodeField).Displayed;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        /// <summary>
        /// Enter first name in the First Name field
        /// </summary>
        /// <param name="firstName">First name to enter</param>
        public void EnterFirstName(string firstName)
        {
            IWebElement firstNameElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(firstNameField));
            firstNameElement.Clear();
            firstNameElement.SendKeys(firstName);
        }
        /// <summary>
        /// Verify First Name field accepts alphabetic input
        /// </summary>
        /// <returns>True if First Name field accepts input</returns>
        public bool IsFirstNameFieldAcceptsInput()
        {
            return _driver.FindElement(firstNameField).Enabled;
        }
        /// <summary>
        /// Enter last name in the Last Name field
        /// </summary>
        /// <param name="lastName">Last name to enter</param>
        public void EnterLastName(string lastName)
        {
            IWebElement lastNameElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(lastNameField));
            lastNameElement.Clear();
            lastNameElement.SendKeys(lastName);
        }
        /// <summary>
        /// Verify Last Name field accepts alphabetic input
        /// </summary>
        /// <returns>True if Last Name field accepts input</returns>
        public bool IsLastNameFieldAcceptsInput()
        {
            return _driver.FindElement(lastNameField).Enabled;
        }
        /// <summary>
        /// Enter zip/postal code in the Zip/Postal Code field
        /// </summary>
        /// <param name="zipCode">Zip/Postal code to enter</param>
        public void EnterZipPostalCode(string zipCode)
        {
            IWebElement zipCodeElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(postalCodeField));
            zipCodeElement.Clear();
            zipCodeElement.SendKeys(zipCode);
        }
        /// <summary>
        /// Verify Zip/Postal Code field accepts numeric input
        /// </summary>
        /// <returns>True if Zip/Postal Code field accepts input</returns>
        public bool IsZipPostalCodeFieldAcceptsInput()
        {
            return _driver.FindElement(postalCodeField).Enabled;
        }
        /// <summary>
        /// Click Continue button to proceed to Checkout Overview
        /// </summary>
        public void ClickContinueButton()
        {
            IWebElement continueBtn = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(continueButton));
            continueBtn.Click();
        }
        /// <summary>
        /// Fill all checkout information fields
        /// </summary>
        /// <param name="firstName">First name</param>
        /// <param name="lastName">Last name</param>
        /// <param name="zipCode">Zip/Postal code</param>
        public void FillCheckoutInformation(string firstName, string lastName, string zipCode)
        {
            EnterFirstName(firstName);
            EnterLastName(lastName);
            EnterZipPostalCode(zipCode);
        }
    }
}