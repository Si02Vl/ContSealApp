using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.IE;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;

namespace ContSealAuth
{
    public class FirstScriptTest
    {
        public static void EdgeSession()
        {
            
            new DriverManager().SetUpDriver(new InternetExplorerConfig());
            var driver = new InternetExplorerDriver();
            driver.Manage().Window.Maximize();

            driver.Navigate().GoToUrl("https://google.com");

            var title = driver.Title;
            Assert.AreEqual("Google", title);

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(500);

            var searchBox = driver.FindElement(By.Name("q"));
            var searchButton = driver.FindElement(By.Name("btnK"));
            
            searchBox.SendKeys("Selenium");
            searchButton.Click();
            
            searchBox = driver.FindElement(By.Name("q"));
            var value = searchBox.GetAttribute("value");
            Assert.AreEqual("Selenium", value);

            driver.Quit();
        }
    }
}
