/**
 * Base test class for NUnit Selenium test automation
 * Provides WebDriver initialization, configuration, and cleanup
 */
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
namespace SwagLabsAutomation.Tests.TestBase
{
    /// <summary>
    /// Base class for all test fixtures providing WebDriver setup and teardown
    /// </summary>
    public class TestBase
    {
        protected IWebDriver Driver;
        protected WebDriverWait Wait;
        /// <summary>
        /// Setup method executed before each test
        /// Initializes WebDriver and navigates to base URL
        /// </summary>
        [SetUp]
        public void SetUp()
        {
            // Initialize Chrome WebDriver
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            options.AddArgument("--disable-notifications");
            Driver = new ChromeDriver(options);
            // Initialize WebDriverWait with 30 seconds timeout
            Wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30));
            // Set implicit wait
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
        }
        /// <summary>
        /// Teardown method executed after each test
        /// Closes browser and cleans up WebDriver resources
        /// </summary>
        [TearDown]
        public void TearDown()
        {
            if (Driver != null)
            {
                Driver.Quit();
                Driver.Dispose();
            }
        }
        /// <summary>
        /// Helper method to wait for element to be visible
        /// </summary>
        /// <param name="by">By locator for the element</param>
        /// <returns>The visible WebElement</returns>
        protected IWebElement WaitForElementVisible(By by)
        {
            return Wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(by));
        }
        /// <summary>
        /// Helper method to wait for element to be clickable
        /// </summary>
        /// <param name="by">By locator for the element</param>
        /// <returns>The clickable WebElement</returns>
        protected IWebElement WaitForElementClickable(By by)
        {
            return Wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(by));
        }
    }
}