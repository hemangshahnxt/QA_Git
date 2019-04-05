using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using SeleniumAutomation;

namespace MyPatientVisit.FunctionalTest
{
    [TestClass]
    public class OnlineBillPay
    {

        //public static SqlConnection AppConnection { get; set; }
        public static readonly string _baseSite = "devqa"; //"www";
        public static readonly string _practicID = "BIKATH"; //"RWSAFH";
        public static string AppConnectionString { get; set; }

        [TestMethod]
        public void VerifyOnlineBillPay()
        {
            IWebDriver ChromeDriver = new ChromeDriver();
            ChromeDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
            ChromeDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            ChromeDriver.Manage().Window.Maximize();
            ChromeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com");
            IWebElement verifyLoginPage = ChromeDriver.FindElement(By.XPath("//input[@placeholder='Enter Username']"));
            WriteFile.WriteToFile("Login Page is Confirmed");
            Assert.IsTrue(verifyLoginPage.Enabled, "Login page is displayed");
            Verification.Sleep();

            //Login in MPV with correct credentials
            ChromeDriver.FindElement(By.XPath("//input[@placeholder='Enter Username']")).SendKeys("selenium");
            ChromeDriver.FindElement(By.XPath("//input[@name='Password']")).SendKeys("Password.1");
            ChromeDriver.FindElement(By.XPath("//input[@value='Login']")).SendKeys(Keys.Enter);
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//*[contains(text(), 'SELECT THE ')]")).Enabled);
            WriteFile.WriteToFile("Select the Medical Record Page is confirmed");
            ChromeDriver.FindElement(By.XPath("(//img[starts-with(@src, 'data:image/JPEG;base64,iVBORw')])[1]")).Click();
            Verification.Sleep();

            //Verify that the dashboard is displayed
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@class='ng-binding']")).Enabled, "Failed :- Dashboard is not present");
            WriteFile.WriteToFile("Dashboard is Present - REGRESSION TEST");
            System.Threading.Thread.Sleep(5000);

            //Navigating to myBillPay tab
            ChromeDriver.FindElement(By.XPath("//a[@class='dropdown-toggle']//span[@translate-once='BILLPAY']//following-sibling::span//parent::a")).Click();
            ChromeDriver.FindElement(By.XPath("//a[@href='#billPay']")).Click();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//span[contains(text(), 'Account Balance')]")).Enabled);
            WriteFile.WriteToFile("Online Bill Pay Page is confirmed");

            //Making a partial Payment
            ChromeDriver.FindElement(By.XPath("//div[@class='md-on']//ancestor::md-radio-button[@aria-label='Partial Payment']")).Click();
            ChromeDriver.FindElement(By.XPath("//input[@id='partialPaymentAmount']")).SendKeys("20.20");

            //Selecting Visa as the card type
            ChromeDriver.FindElement(By.XPath("//div[@class='md-on']//ancestor::md-radio-button[@aria-label='Visa']")).Click();
            ChromeDriver.SwitchTo().Frame(ChromeDriver.FindElement(By.XPath("//iframe[@id='tokenframe']")));

            ChromeDriver.FindElement(By.XPath("//input[@id='ccnumfield']")).SendKeys("4111111111111111");
            ChromeDriver.SwitchTo().DefaultContent();
            ChromeDriver.FindElement(By.XPath("//input[@name='expiration']")).SendKeys("102020");
            ChromeDriver.FindElement(By.XPath("//input[@type='submit']")).Click();

            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//*[contains(text(),'Yes')]")).Enabled);
            WriteFile.WriteToFile("Confirmation Dialog is Present");
            ChromeDriver.FindElement(By.XPath("//*[contains(text(),'Yes')]")).Click();

            //Verifying if the Payment Confirmation Page is displayed
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//*[contains(text(), 'Thank you')]")).Enabled);
            WriteFile.WriteToFile("Payment Confirmation Page is displayed");

            //Clicking the return to dashboard button
            ChromeDriver.FindElement(By.XPath("//i[@aria-label='Go back to the dashboard']")).Click();

            //Verifying that the Dashboard is present
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@class='ng-binding']")).Displayed);

            //Navigating to myBillPay tab
            ChromeDriver.FindElement(By.XPath("//a[@class='dropdown-toggle']//span[@translate-once='BILLPAY']//following-sibling::span//parent::a")).Click();
            ChromeDriver.FindElement(By.XPath("//a[@href='#billPay']")).Click();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//span[contains(text(), 'Account Balance')]")).Enabled);
            WriteFile.WriteToFile("Online Bill Pay Page is confirmed");

            //Making a Full Payment
            ChromeDriver.FindElement(By.XPath("//div[@class='md-on']//ancestor::md-radio-button[@aria-label='Full Payment']")).Click();

            //Selecting Visa as the card type
            ChromeDriver.FindElement(By.XPath("//div[@class='md-on']//ancestor::md-radio-button[@aria-label='Visa']")).Click();
            ChromeDriver.SwitchTo().Frame(ChromeDriver.FindElement(By.XPath("//iframe[@id='tokenframe']")));

            ChromeDriver.FindElement(By.XPath("//input[@id='ccnumfield']")).SendKeys("4111111111111111");
            ChromeDriver.SwitchTo().DefaultContent();
            ChromeDriver.FindElement(By.XPath("//input[@name='expiration']")).SendKeys("102020");
            ChromeDriver.FindElement(By.XPath("//input[@type='submit']")).Click();

            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//*[contains(text(),'Yes')]")).Enabled);
            WriteFile.WriteToFile("Confirmation Dialog is Present");
            ChromeDriver.FindElement(By.XPath("//*[contains(text(),'Yes')]")).Click();

            //Verifying if the Payment Confirmation Page is displayed
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//*[contains(text(), 'Thank you')]")).Enabled);
            WriteFile.WriteToFile("Payment Confirmation Page is displayed");

            //Clicking the return to dashboard button
            ChromeDriver.FindElement(By.XPath("//i[@aria-label='Go back to the dashboard']")).Click();

            //Verifying that the Dashboard is present
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@class='ng-binding']")).Displayed);

            //Logging off from the application
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='SETTINGS']")).Click();
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='LOGOFF']")).Click();
            ChromeDriver.Close();

        }


    }
}

