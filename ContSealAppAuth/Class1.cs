using OpenQA.Selenium.Firefox;

namespace SeleniumDocs.Hello
{
    public class HelloSelenium
    {
        public static void Auth()
        {
            var driver = new FirefoxDriver();
            
            driver.Navigate().GoToUrl("https://selenium.dev");
            
            driver.Quit();
        }
    }
}
