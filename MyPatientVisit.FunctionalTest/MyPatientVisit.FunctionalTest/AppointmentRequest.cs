using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using SeleniumAutomation;

namespace MyPatientVisit.FunctionalTest
{
    [TestClass]
    public class AppointmentRequest
    { 
        public static readonly string _baseSite = "devqa"; //"www";
        public static readonly string _practicID = "BIKATH"; //"RWSAFH";
        public static string AppConnectionString { get; set; }

        [TestMethod]
        public void VerifyAppointmentRequest()
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

            Actions action = new Actions(ChromeDriver);
            
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

            //Navigatinig to the appointment request tab
            ChromeDriver.FindElement(By.XPath("//a[@class='dropdown-toggle']//span[@translate-once='APPOINTMENTS']//following-sibling::span//parent::a")).Click();
            ChromeDriver.FindElement(By.XPath("//a[@href='#appointmentrequest']//strong[@translate-once='REQUEST_APPT']")).Click();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(), 'Appointment Request')]")).Enabled);
            WriteFile.WriteToFile("Appointment Request Page is Confirmed");

            //Selecting a Provider
            //ChromeDriver.FindElement(By.XPath("//md-select[@aria-label='Provider']")).SendKeys(Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//md-select[@aria-label='Provider']")).Click();
            ChromeDriver.FindElement(By.XPath("//div[contains(text(), 'Hemang')]")).Click();

            //Selecting a Location
            ChromeDriver.FindElement(By.XPath("//md-select[@aria-label='Location:']")).Click();
            System.Threading.Thread.Sleep(1000);
            ChromeDriver.FindElement(By.XPath("//md-option[@value='1' and @ng-repeat='location in Request.locations']")).Click();

            //Selecting a Reason for Visit
            ChromeDriver.FindElement(By.XPath("//md-select[@aria-label='Reason for Visit']")).Click();
            System.Threading.Thread.Sleep(1000);
            ChromeDriver.FindElement(By.XPath("//md-option[@value=33 and @ng-repeat='apptType in Request.apptTypes']")).Click();

            //Selecting Time of Day
            ChromeDriver.FindElement(By.XPath("//md-select[@aria-label='Time of Day']")).Click();
            System.Threading.Thread.Sleep(1000);
            ChromeDriver.FindElement(By.XPath("//md-option[@tabindex=3 and @ng-repeat='time in apptTimes']")).Click();
            action.SendKeys(Keys.Escape).Perform();

            //Selecting a a few Days
           
            System.Threading.Thread.Sleep(1000);
            ChromeDriver.FindElement(By.XPath("//div[contains(text(), 'Monday')]")).Click();
            ChromeDriver.FindElement(By.XPath("//div[contains(text(), 'Wednesday')]")).Click();
            ChromeDriver.FindElement(By.XPath("//div[contains(text(), 'Friday')]")).Click();

            //Enterting some comments
            ChromeDriver.FindElement(By.XPath("//textarea[@name='comments']")).SendKeys("Testing appointment request via Selenium");

            ChromeDriver.FindElement(By.XPath("//input[@value='Request Appointment']")).Click();
            ChromeDriver.FindElement(By.XPath("//button[@type='button' and contains(text(), 'OK')]")).Click();
            //Verifying that the Dashboard is present
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@class='ng-binding']")).Displayed);
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='SETTINGS']")).Click();
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='LOGOFF']")).Click();
            ChromeDriver.Close();
        }
    }
}
