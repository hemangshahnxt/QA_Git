using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;

namespace SeleniumAutomation
{

    
    class MPVAdmin
    {
        public static readonly string _baseUsername = "Hemang1";
        public static readonly string _basePassword = "Hemang.78";
        public static ChromeDriver Chrome { get; private set; }
        
        public void CreateAdminClients()
        {
            string strName = Verification.GenerateRandomString();
            Chrome = new ChromeDriver();
            Chrome.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
            Chrome.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            Chrome.Manage().Window.Maximize();
            Chrome.Navigate().GoToUrl("https://devqa.mypatientvisit.com");
            Chrome.FindElement(By.XPath(".//input[@placeholder='Enter Username']")).SendKeys( _baseUsername);
            Chrome.FindElement(By.XPath(".//input[@placeholder='Password']")).SendKeys(_basePassword);
            Chrome.FindElement(By.XPath(".//input[@value='Login']")).Click();
            Verification.Sleep();

            Chrome.FindElement(By.XPath(".//button[@ng-click='clients()']")).Click();

        }

    }
}
