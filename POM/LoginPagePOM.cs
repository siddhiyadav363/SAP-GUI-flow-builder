/**
 * Login Page POM class
 * Contains locators and methods for Login page interactions
 */
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
namespace SwagLabsAutomation.POM
{
    /// <summary>
    /// Page Object Model for Login Page
    /// </summary>
    public class LoginPagePOM
    {
        private readonly IWebDriver _driver;
        private readonly WebDriverWait _wait;
        // Locators using XPath from provided list
        private readonly By usernameField = By.XPath("//input[@id='user-name']");
        private readonly By passwordField = By.XPath("//input[@id='password']");
        private readonly By loginButton = By.XPath("//input[@id='login-button']");
        private readonly By loginCredentialsDiv = By.XPath("//div[@id='login_credentials']");
        /// <summary>
        /// Constructor for LoginPagePOM
        /// </summary>
        /// <param name="driver">WebDriver instance</param>
        public LoginPagePOM(IWebDriver driver)
        {
            _driver = driver;
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
        }
        /// <summary>
        /// Navigate to the login page
        /// </summary>
        /// <param name="url">Base URL from test data</param>
        public void NavigateToLoginPage(string url)
        {
            _driver.Navigate().GoToUrl(url);
        }
        /// <summary>
        /// Verify that login page has loaded successfully
        /// </summary>
        /// <returns>True if login page elements are displayed</returns>
        public bool IsLoginPageDisplayed()
        {
            try
            {
                _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(usernameField));
                return _driver.FindElement(usernameField).Displayed &&
                       _driver.FindElement(passwordField).Displayed &&
                       _driver.FindElement(loginButton).Displayed;
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }
        /// <summary>
        /// Enter username in the username field
        /// </summary>
        /// <param name="username">Username to enter</param>
        public void EnterUsername(string username)
        {
            IWebElement usernameElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(usernameField));
            usernameElement.Clear();
            usernameElement.SendKeys(username);
        }
        /// <summary>
        /// Verify username field accepts input
        /// </summary>
        /// <returns>True if username field accepts input</returns>
        public bool IsUsernameFieldAcceptsInput()
        {
            return _driver.FindElement(usernameField).Enabled;
        }
        /// <summary>
        /// Enter password in the password field
        /// </summary>
        /// <param name="password">Password to enter</param>
        public void EnterPassword(string password)
        {
            IWebElement passwordElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(passwordField));
            passwordElement.Clear();
            passwordElement.SendKeys(password);
        }
        /// <summary>
        /// Verify password field accepts input and masks characters
        /// </summary>
        /// <returns>True if password field type is 'password'</returns>
        public bool IsPasswordFieldMasked()
        {
            IWebElement passwordElement = _driver.FindElement(passwordField);
            return passwordElement.GetAttribute("type").Equals("password");
        }
        /// <summary>
        /// Click the Login button
        /// </summary>
        public void ClickLoginButton()
        {
            IWebElement loginBtn = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(loginButton));
            loginBtn.Click();
        }
        /// <summary>
        /// Verify login button is clickable
        /// </summary>
        /// <returns>True if login button is enabled and displayed</returns>
        public bool IsLoginButtonClickable()
        {
            IWebElement loginBtn = _driver.FindElement(loginButton);
            return loginBtn.Enabled && loginBtn.Displayed;
        }
        /// <summary>
        /// Perform complete login action
        /// </summary>
        /// <param name="username">Username for login</param>
        /// <param name="password">Password for login</param>
        public void Login(string username, string password)
        {
            EnterUsername(username);
            EnterPassword(password);
            ClickLoginButton();
        }
    }
}