using System;
using System.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Edge;

namespace SeleniumAutomation
{
    [TestClass]
    public class MPV
    {
        public static readonly string _baseSite = "devqa"; //"www";
        public static readonly string _practicID = "BIKATH"; //"RWSAFH";
        
        // TODO: add more entries here for the remaining hard-coded stuff, e.g. username "hemang", passwords, etc.
              
        string methodName = System.Reflection.MethodBase.GetCurrentMethod().Name;
        [TestMethod]


        [Priority(2)]
        public void TestFireFoxDriver()
        {

            // FirefoxDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
            FirefoxDriver FirefoxDriver = new FirefoxDriver();
            FirefoxDriver.Navigate().GoToUrl($"https://{ConfigurationManager.AppSettings["_baseSite"]}.mypatientvisit.com");
            FirefoxDriver.Manage().Window.Maximize();
            Verification.Sleep();
            IWebElement usernameField = FirefoxDriver.FindElement(By.XPath("//input[@placeholder='Enter Username']"));
            Console.WriteLine("MPV successful in FireFox");
            WriteFile.WriteToFile($"MPV successful in FireFox with URL - https://{_baseSite}.mypatientvisit.com");
            System.Diagnostics.Debug.WriteLine($"MPV SUCCESSFUL IN FIREFOX - https://{_baseSite}.mypatientvisit.com");
            FirefoxDriver.FindElement(By.XPath("//input[@placeholder='Enter Username']")).SendKeys("hemang");
            FirefoxDriver.FindElement(By.XPath("//input[@name='Password']")).SendKeys("Password.1");
            FirefoxDriver.FindElement(By.XPath("//input[@value='Login']")).SendKeys(Keys.Enter);
            FirefoxDriver.Close();
            
            Verification.Sleep();

            FirefoxDriver = new FirefoxDriver();
            FirefoxDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com/#/register");
            FirefoxDriver.Manage().Window.Maximize();
            Verification.Sleep();
            Assert.IsTrue(FirefoxDriver.FindElement(By.XPath("//input[@id='Username']")).Displayed);
            Console.WriteLine("MPV successful in FireFox");
            WriteFile.WriteToFile($"MPV successful in FireFox with URL - https://{_baseSite}.mypatientvisit.com/#/register");
            System.Diagnostics.Debug.WriteLine($"MPV SUCCESSFUL IN FIREFOX - https://{_baseSite}.mypatientvisit.com/#/register");
            FirefoxDriver.Close();
            Verification.Sleep();

            FirefoxDriver = new FirefoxDriver();
            FirefoxDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com/#/register?practiceID=BIKATH");
            FirefoxDriver.Manage().Window.Maximize();
            Verification.Sleep();
            Assert.IsTrue(FirefoxDriver.FindElement(By.XPath("//input[@id='Username']")).Displayed);
            Console.WriteLine("MPV successful in FireFox");
            WriteFile.WriteToFile($"MPV successful in FireFox with URL - https://{_baseSite}.mypatientvisit.com/#/register?practiceID={_practicID}");
            System.Diagnostics.Debug.WriteLine($"MPV SUCCESSFUL IN FIREFOX - https://{_baseSite}.mypatientvisit.com/#/register?practiceID={_practicID}");

            FirefoxDriver.Close();
            // ChromeDriver.Close();

        }

        [Priority(3)]
        [TestMethod]
        public void MPVRegression()
        {
            IWebDriver ChromeDriver = new ChromeDriver();
            ChromeDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
            ChromeDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            ChromeDriver.Manage().Window.Maximize();
            ChromeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com");
            IWebElement usernameField = ChromeDriver.FindElement(By.XPath("//input[@placeholder='Enter Username']"));
            Console.WriteLine("Login Page is Confirmed");
            WriteFile.WriteToFile("Login Page is Confirmed");
            Assert.IsTrue(usernameField.Enabled, "Failed :- Login page is not found - REGRESSION TEST");
            Verification.Sleep();
            ChromeDriver.FindElement(By.XPath("//input[@placeholder='Enter Username']")).SendKeys("selenium");
            ChromeDriver.FindElement(By.XPath("//input[@name='Password']")).SendKeys("Password.1");
            ChromeDriver.FindElement(By.XPath("//input[@value='Login']")).SendKeys(Keys.Enter);
            ChromeDriver.FindElement(By.XPath("(//img[starts-with(@src, 'data:image/JPEG;base64,iVBORw')])[1]")).Click();
            ChromeDriver.FindElement(By.XPath("(//img[starts-with(@src, 'data:image/JPEG;base64,iVBORw')])[1]")).Click();

            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@class='ng-binding']")).Enabled, "Failed :- Dashboard is not present");
            WriteFile.WriteToFile("Dashboard is Present - REGRESSION TEST");
            System.Threading.Thread.Sleep(5000);
            //Navigating to demographics page
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='CHART']")).Click();
            ChromeDriver.FindElement(By.XPath("//a[@href='#generaldemographics']")).Click();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@translate-once='ENTER_INFO']")).Enabled, "Failed :- Demographic Page is not Found");


            if (ChromeDriver.FindElement(By.XPath("//div[@translate-once='PATIENT_INFO_NOTE_SAVE']")).Displayed)
            {
                Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//div[@translate-once='PHONE_REQUIRED']")).Enabled, "Failed :- Phone number is not missing");
                ChromeDriver.FindElement(By.XPath("//label[@for='input_17']//following-sibling::input[@name='homePhone']")).SendKeys("1234567890");
                ChromeDriver.FindElement(By.XPath("//a[@ng-click='save()']")).Click();
            }
            else
            {
                ChromeDriver.FindElement(By.XPath("//a[@ng-click='save()']")).Click();
            }
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@class='ng-binding']")).Enabled, "Failed :- Dashboard is not present");

            //Navigating to Insurance tab
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='CHART']")).Click();
            ChromeDriver.FindElement(By.XPath("//a[@href='#insurance']")).Click();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@translate-once='ENTER_INSURANCE_INFO']")).Displayed, "False:- Insurance Tab is not Present");
            System.Threading.Thread.Sleep(1000);
            ChromeDriver.FindElement(By.XPath("//a[@ng-click='save()']")).Click();
            System.Threading.Thread.Sleep(3000);
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@class='ng-binding']")).Enabled, "Failed :- Dashboard is not present");

            //Navigating to the Document Summary Page
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='CHART']")).Click();
            ChromeDriver.FindElement(By.XPath("//a[@href='#documents']")).Click();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@class='ng-binding']")).Displayed);
            ChromeDriver.FindElement(By.XPath("//md-checkbox[@ng-model='visit.selected']")).Click();
            ChromeDriver.FindElement(By.XPath("//input[@ng-click='view(listVisitsSelected)']")).Click();
            System.Threading.Thread.Sleep(3000);
            var tabs = ChromeDriver.WindowHandles;

            if (tabs.Count > 1)
            {
                System.Threading.Thread.Sleep(1000);
                ChromeDriver.SwitchTo().Window(tabs[1]);
                ChromeDriver.Close();
                ChromeDriver.SwitchTo().Window(tabs[0]);
            }

            //Transmitting a secure message
            ChromeDriver.FindElement(By.XPath("//md-checkbox[@ng-model='visit.selected']")).Click();
            ChromeDriver.FindElement(By.XPath("//input[@ng-click='transmitSecurely(listVisitsSelected)']")).Click();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(), 'Transmit Documents Securely')]")).Displayed);
            ChromeDriver.FindElement(By.XPath("//input[@id ='input_14']")).SendKeys("nextech@directaddress.net");


            //Navigating to the forms page
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='CHART']")).Click();
            ChromeDriver.FindElement(By.XPath("//a[@href='#emr']")).Click();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@ng-show='!appointmentId']")).Displayed);

            //Navigating to my Messages tab
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='CHART']")).Click();
            ChromeDriver.FindElement(By.XPath("//a[@href='#message']")).Click();
            Verification.Sleep();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@class='ng-binding ng-scope']")).Displayed);
            ChromeDriver.FindElement(By.XPath("//button[@ng-show='!sendNewMessage']")).Click();
            ChromeDriver.FindElement(By.XPath("//button[@ng-show='cancelFlag']")).Click();

            //Navigating to Upload Document tab
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='CHART']")).Click();
            ChromeDriver.FindElement(By.XPath("//a[@ng-click='uploadDocument()']")).Click();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//input[@value='Submit ']")).Displayed);
            WriteFile.WriteToFile("Upload Document Dialog is Present");
            ChromeDriver.FindElement(By.XPath("//input[@value='Cancel']")).Click();


            //Clicking the mySetting dropdown and changing the profile picture
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='SETTINGS']")).Click();
            ChromeDriver.FindElement(By.XPath("//a[@ng-click='changeProfilePicture()']")).Click();
            System.Threading.Thread.Sleep(1000);
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(), 'Change My Patient Picture')]")).Displayed);
            WriteFile.WriteToFile("CHange My Patient Picture Dialog is Visible");
            ChromeDriver.FindElement(By.XPath("//input[@value='Cancel']")).Click();

            //Verifying Manage My Patients 
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='SETTINGS']")).Click();
            ChromeDriver.FindElement(By.XPath("//strong[@translate-once='MANAGE_PATIENTS']")).Click();
            System.Threading.Thread.Sleep(1000);
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(), 'SELECT THE MEDICAL')]")).Displayed);
            WriteFile.WriteToFile("Select The Medical Record page is visible");
            ChromeDriver.Close();

        }

        [Priority(4)]
        [TestMethod]
        public void MPVregisterOnChrome()
        {
            ChromeDriver ChromeDriver = new ChromeDriver();
            ChromeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com/#/register");
            IWebElement usernameField1 = ChromeDriver.FindElement(By.XPath("//input[@id='Username']"));
            Assert.IsTrue(usernameField1.Enabled, "Fail :- Registration page is not found");
            Console.WriteLine("Registration Page is confirmed");
            WriteFile.WriteToFile($"Registration Page is confirmed with URL - https://{_baseSite}.mypatientvisit.com/#/register");
            // System.Diagnostics.Debug.WriteLine("MPV Registration successful");
            ChromeDriver.Close();
            Verification.Sleep();

            ChromeDriver = new ChromeDriver();
            ChromeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com");
            ChromeDriver.Manage().Window.Maximize();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(), 'WELCOME')]")).Displayed);
            Console.WriteLine("Link verified on Chrome");
            WriteFile.WriteToFile($"Registration Page is confirmed with URL - https://{_baseSite}.mypatientvisit.com");
            ChromeDriver.Close();
            Verification.Sleep();


            ChromeDriver = new ChromeDriver();
            ChromeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com/#/login");
            ChromeDriver.Manage().Window.Maximize();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(), 'WELCOME')]")).Displayed);
            Console.WriteLine("Link verified on Chrome");
            WriteFile.WriteToFile($"Registration Page is confirmed with URL - https://{_baseSite}.mypatientvisit.com/#/login");
            ChromeDriver.Close();
            Verification.Sleep();

            ChromeDriver = new ChromeDriver();
            ChromeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com/#/register?practiceID={_practicID}");
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//input[@id='Username']")).Displayed);
            Console.WriteLine("Registration Page is confirmed");
            WriteFile.WriteToFile($"Registration Page is confirmed with URL - https://{_baseSite}.mypatientvisit.com/#/register?practiceID={_practicID}");
            ChromeDriver.Close();
            Verification.Sleep();

        }

        //This Test ensures landing pages on Edge Browser
        [Priority(1)]
        [TestMethod]
        public void MPVRegisterwithEdge()
        {
            IWebDriver EdgeDriver = new EdgeDriver();
            EdgeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com/#/register?practiceID={_practicID}");
            EdgeDriver.Manage().Window.Maximize();
            Verification.Sleep();
            Assert.IsTrue(EdgeDriver.FindElement(By.XPath("//input[@type='submit']")).Displayed);
            Console.WriteLine("Registration page on Edge is confirmed with Practice ID");
            WriteFile.WriteToFile($"Registration page on EDGE is confirmed with Practice ID = {_practicID}");
            EdgeDriver.Close();
            Verification.Sleep();

            EdgeDriver = new EdgeDriver();
            EdgeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com/#/register");
            // EdgeDriver.Manage().Window.Maximize();
            Verification.Sleep();
            Assert.IsTrue(EdgeDriver.FindElement(By.XPath("//input[@type='submit']")).Displayed);
            Console.WriteLine("Registration page on Edge is confirmed WITHOUT Practice ID");
            WriteFile.WriteToFile("Registration page on Edge is confirmed WITHOUT Practice ID");
            EdgeDriver.Close();
            Verification.Sleep();

            EdgeDriver = new EdgeDriver();
            EdgeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com/");
            EdgeDriver.Manage().Window.Maximize();
            Verification.Sleep();
            Assert.IsTrue(EdgeDriver.FindElement(By.XPath("//input[@value='Login']")).Displayed);
            Console.WriteLine("Registration page on Edge is confirmed");
            WriteFile.WriteToFile($"Registration page on Edge is confirmed with URL - https://{_baseSite}.mypatientvisit.com/");
            EdgeDriver.Close();
            Verification.Sleep();

            EdgeDriver = new EdgeDriver();
            EdgeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com/#/login");
            EdgeDriver.Manage().Window.Maximize();
            Verification.Sleep();
            Assert.IsTrue(EdgeDriver.FindElement(By.XPath("//input[@value='Login']")).Displayed);
            Console.WriteLine("Registration page on Edge is confirmed");
            WriteFile.WriteToFile($"Registration page on Edge is confirmed with URL - https://{_baseSite}.mypatientvisit.com/#/login");
            EdgeDriver.Close();
            Verification.Sleep();
        }

        // Recover Username Test
        [Priority(5)]
        [TestMethod]
        public void RecoverUsername()
        {
            IWebDriver ChromeDriver = new ChromeDriver();
            ChromeDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
            ChromeDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            ChromeDriver.Manage().Window.Maximize();
            ChromeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com");
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(),'WELCOME')]")).Displayed);

            //Clicking on the Need help with login link
            ChromeDriver.FindElement(By.XPath("//a[@href='#recoverAccount']")).Click();
            Verification.Sleep();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2")).Displayed);
            ChromeDriver.FindElement(By.XPath("//md-radio-button[@ng-value='forgotUsername']")).Click();
            ChromeDriver.FindElement(By.XPath("//input[@id='FirstName']")).SendKeys("Selenium");
            ChromeDriver.FindElement(By.XPath("//input[@id='LastName']")).SendKeys("Automation");
            ChromeDriver.FindElement(By.XPath("//input[@name='DateOfBirth']")).SendKeys("10101985");
            ChromeDriver.FindElement(By.XPath("//input[@id='ZipCode']")).SendKeys("33609");

            if (ChromeDriver.FindElement(By.XPath("//input[@id='PracticeID']")).Displayed)
            {
                ChromeDriver.FindElement(By.XPath("//input[@id='PracticeID']")).SendKeys(_practicID);
            }
            ChromeDriver.FindElement(By.XPath("//input[@type='submit']")).Click();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(),'RECOVER')]")).Displayed);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 1:')]//following::input[1]")).SendKeys("1");
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 1:')]//following::input[2]")).SendKeys("1");
            ChromeDriver.FindElement(By.XPath("//input[@value='Continue']")).Click();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//p[@class='security-question-header ng-binding']//child::em[contains(text(), 'Selenium1')]")).Displayed);
            ChromeDriver.Close();
        }

        //Recover Password Test
        [Priority(6)]
        [TestMethod]
        public void RecoverPassword()
        {
            IWebDriver ChromeDriver = new ChromeDriver();
            ChromeDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
            ChromeDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            ChromeDriver.Manage().Window.Maximize();
            ChromeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com");
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(),'WELCOME')]")).Displayed);

            //Clicking on the Need help with login link
            ChromeDriver.FindElement(By.XPath("//a[@href='#recoverAccount']")).Click();
            Verification.Sleep();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2")).Displayed);
            //Selecting the forgot password radio button
            ChromeDriver.FindElement(By.XPath("//md-radio-button[@ng-value='forgotPassword']")).Click();
            ChromeDriver.FindElement(By.XPath("//input[@id='FirstName']")).SendKeys("Selenium");
            ChromeDriver.FindElement(By.XPath("//input[@id='LastName']")).SendKeys("Automation");
            ChromeDriver.FindElement(By.XPath("//input[@id='Username']")).SendKeys("selenium1");
            ChromeDriver.FindElement(By.XPath("//input[@name='DateOfBirth']")).SendKeys("10101985");
            ChromeDriver.FindElement(By.XPath("//input[@id='ZipCode']")).SendKeys("33609");
            if (ChromeDriver.FindElement(By.XPath("//input[@id='PracticeID']")).Displayed)
            {
                ChromeDriver.FindElement(By.XPath("//input[@id='PracticeID']")).SendKeys(_practicID);
            }
            ChromeDriver.FindElement(By.XPath("//input[@type='submit']")).Click();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(),'RECOVER')]")).Displayed);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 1:')]//following::input[1]")).SendKeys("1");
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 1:')]//following::input[2]")).SendKeys("1");
            ChromeDriver.FindElement(By.XPath("//input[@value='Continue']")).Click();
            Verification.Sleep();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(), 'RESET PASSWORD')]")).Displayed);
            ChromeDriver.FindElement(By.XPath("//input[@id='Password']")).SendKeys("Password.1");
            ChromeDriver.FindElement(By.XPath("//input[@id='ConfirmPassword']")).SendKeys("Password.1");
            ChromeDriver.FindElement(By.XPath("//input[@value='Submit']")).Click();
            Verification.Sleep();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h4[contains(text(), 'Password success')]")).Displayed);
            ChromeDriver.FindElement(By.XPath("//button[@ng-click='closeDialog()']")).Click();
            ChromeDriver.Close();
        }

        //Recover both Username and Password Test
        [Priority(7)]
        [TestMethod]
        public void RecoverCredential()
        {
            IWebDriver ChromeDriver = new ChromeDriver();
            ChromeDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
            ChromeDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            ChromeDriver.Manage().Window.Maximize();
            ChromeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com");
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(),'WELCOME')]")).Displayed);
            //Clicking on the Need help with login link
            ChromeDriver.FindElement(By.XPath("//a[@href='#recoverAccount']")).Click();
            Verification.Sleep();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(), 'FORGOT LOGIN CREDENTIALS?')]")).Displayed);
            //Selecting the Forgot Both Radio button
            ChromeDriver.FindElement(By.XPath("//md-radio-button[@ng-value='forgotBoth']")).Click();
            ChromeDriver.FindElement(By.XPath("//input[@id='FirstName']")).SendKeys("Selenium");
            ChromeDriver.FindElement(By.XPath("//input[@id='LastName']")).SendKeys("Automation");
            ChromeDriver.FindElement(By.XPath("//input[@name='DateOfBirth']")).SendKeys("10101985");
            ChromeDriver.FindElement(By.XPath("//input[@id='ZipCode']")).SendKeys("33609");

            if (ChromeDriver.FindElement(By.XPath("//input[@id='PracticeID']")).Displayed)
            {
                ChromeDriver.FindElement(By.XPath("//input[@id='PracticeID']")).SendKeys(_practicID);
            }
            ChromeDriver.FindElement(By.XPath("//input[@type='submit']")).Click();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(),'RECOVER')]")).Displayed);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 1:')]//following::input[1]")).SendKeys("1");
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 1:')]//following::input[2]")).SendKeys("1");
            ChromeDriver.FindElement(By.XPath("//input[@value='Continue']")).Click();
            Verification.Sleep();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//p[@class='security-question-header ng-binding']//child::em[contains(text(), 'Selenium1')]")).Displayed);
            ChromeDriver.FindElement(By.XPath("//input[@value='Reset Password']")).Click();
            Verification.Sleep();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(), 'RESET PASSWORD')]")).Displayed);
            ChromeDriver.FindElement(By.XPath("//input[@id='Password']")).SendKeys("Password.1");
            ChromeDriver.FindElement(By.XPath("//input[@id='ConfirmPassword']")).SendKeys("Password.1");
            ChromeDriver.FindElement(By.XPath("//input[@value='Submit']")).Click();
            Verification.Sleep();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h4[contains(text(), 'Password success')]")).Displayed);
            ChromeDriver.FindElement(By.XPath("//button[@ng-click='closeDialog()']")).Click();
            ChromeDriver.Close();
        }

        //Register a patient in MPV With practice ID in the URL
        [Priority(8)]
        [TestMethod]
        public void RegistrationWithPracticeID()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("headless");
            IWebDriver ChromeDriver = new ChromeDriver(options);
            ChromeDriver.Manage().Window.Maximize();
            ChromeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com/#/register?practiceID={_practicID}");
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//input[@id='Username']")).Displayed);
            Console.WriteLine("Registration Page is confirmed");
            //Registration page is confirmed
            string strName = Verification.GenerateRandomString();
            //Generating a random string for the username starting with SeleniumTest
            ChromeDriver.FindElement(By.XPath("//input[@id='Username']")).SendKeys(strName);
            ChromeDriver.FindElement(By.XPath("//input[@id='Password']")).SendKeys("Password.1");
            ChromeDriver.FindElement(By.XPath("//input[@id='ConfirmPassword']")).SendKeys("Password.1");

            ChromeDriver.FindElement(By.XPath("//input[@id='FirstName']")).SendKeys("Smoke");
            ChromeDriver.FindElement(By.XPath("//input[@id='LastName']")).SendKeys("Test");
            ChromeDriver.FindElement(By.XPath("//input[@name='DateOfBirth']")).SendKeys("10101985");
            ChromeDriver.FindElement(By.XPath("//input[@id='ZipCode']")).SendKeys("33609");
            ChromeDriver.FindElement(By.XPath("//input[@id='SecurityCode']")).SendKeys("433723727");
            ChromeDriver.FindElement(By.XPath("//input[@type='checkbox']")).Click();
            ChromeDriver.FindElement(By.XPath("//input[@value='Create Account']")).Click();
            Verification.Sleep();
            //Confirming if the Registration is successful by confirming the Forget password page is available
            // Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(),'JUST IN CASE')]")).Displayed);
            //Selecting and answering Question1
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion0']//parent::div")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion0']")).SendKeys(Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 1:')]//following::input[1]")).SendKeys("1");

            //Selecting and ansering Question2
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion1']")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion1']")).SendKeys(Keys.Down + Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 2:')]//following::input[1]")).SendKeys("1");

            //Selecting and answering Question 3
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion2']")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion2']")).SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 3:')]//following::input[1]")).SendKeys("1");
            Verification.Sleep();
            ChromeDriver.FindElement(By.XPath("//input[@value='Continue']")).Click();
            Verification.Sleep();
            Verification.Sleep();
            //Verifying that the Dashboard is present
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@class='ng-binding']")).Displayed);
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='SETTINGS']")).Click();
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='LOGOFF']")).Click();
            ChromeDriver.Close();
        }

        //Register a patient WITHOUT Practice ID in the URL
        [Priority(9)]
        [TestMethod]
        public void RegistrationWithoutPracticeID()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("headless");
            IWebDriver ChromeDriver = new ChromeDriver(options);            
            ChromeDriver.Manage().Window.Maximize();
            ChromeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com/#/register");
            Verification.Sleep();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//input[@id='Username']")).Displayed);
            Console.WriteLine("Registration page is confirmed");
            // Calling the function RegistrationProcess
            ChromeDriver.FindElement(By.XPath("//input[@id='PracticeID']")).SendKeys(_practicID);
            string strName = Verification.GenerateRandomString();
            //Generating a random string for the username starting with SeleniumTest
            ChromeDriver.FindElement(By.XPath("//input[@id='Username']")).SendKeys(strName);
            ChromeDriver.FindElement(By.XPath("//input[@id='Password']")).SendKeys("Password.1");
            ChromeDriver.FindElement(By.XPath("//input[@id='ConfirmPassword']")).SendKeys("Password.1");

            ChromeDriver.FindElement(By.XPath("//input[@id='FirstName']")).SendKeys("Smoke");
            ChromeDriver.FindElement(By.XPath("//input[@id='LastName']")).SendKeys("Test");
            ChromeDriver.FindElement(By.XPath("//input[@name='DateOfBirth']")).SendKeys("10101985");
            ChromeDriver.FindElement(By.XPath("//input[@id='ZipCode']")).SendKeys("33609");
            ChromeDriver.FindElement(By.XPath("//input[@id='SecurityCode']")).SendKeys("433723727");
            ChromeDriver.FindElement(By.XPath("//input[@type='checkbox']")).Click();
            ChromeDriver.FindElement(By.XPath("//input[@value='Create Account']")).Click();
            Verification.Sleep();
            //Confirming if the Registration is successful by confirming the Forget password page is available
            // Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(),'JUST IN CASE')]")).Displayed);
            //Selecting and answering Question1
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion0']//parent::div")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion0']")).SendKeys(Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 1:')]//following::input[1]")).SendKeys("1");

            //Selecting and ansering Question2
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion1']")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion1']")).SendKeys(Keys.Down + Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 2:')]//following::input[1]")).SendKeys("1");

            //Selecting and answering Question 3
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion2']")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion2']")).SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 3:')]//following::input[1]")).SendKeys("1");
            Verification.Sleep();
            ChromeDriver.FindElement(By.XPath("//input[@value='Continue']")).Click();
            Verification.Sleep();
            Verification.Sleep();
            //Verifying that the Dashboard is present
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@class='ng-binding']")).Displayed);
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='SETTINGS']")).Click();
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='LOGOFF']")).Click();
            ChromeDriver.Close();
        }

        private static void RegistrationProcess()
        {
            IWebDriver ChromeDriver = new ChromeDriver();

            string strName = Verification.GenerateRandomString();
            //Generating a random string for the username starting with SeleniumTest
            ChromeDriver.FindElement(By.XPath("//input[@id='Username']")).SendKeys(strName);
            ChromeDriver.FindElement(By.XPath("//input[@id='Password']")).SendKeys("Password.1");
            ChromeDriver.FindElement(By.XPath("//input[@id='ConfirmPassword']")).SendKeys("Password.1");

            ChromeDriver.FindElement(By.XPath("//input[@id='FirstName']")).SendKeys("Smoke");
            ChromeDriver.FindElement(By.XPath("//input[@id='LastName']")).SendKeys("Test");
            ChromeDriver.FindElement(By.XPath("//input[@name='DateOfBirth']")).SendKeys("10101985");
            ChromeDriver.FindElement(By.XPath("//input[@id='ZipCode']")).SendKeys("33609");
            ChromeDriver.FindElement(By.XPath("//input[@id='SecurityCode']")).SendKeys("433723727");
            ChromeDriver.FindElement(By.XPath("//input[@type='checkbox']")).Click();
            ChromeDriver.FindElement(By.XPath("//input[@value='Create Account']")).Click();
            Verification.Sleep();
            //Confirming if the Registration is successful by confirming the Forget password page is available
            // Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(),'JUST IN CASE')]")).Displayed);
            //Selecting and answering Question1
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion0']//parent::div")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion0']")).SendKeys(Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 1:')]//following::input[1]")).SendKeys("1");

            //Selecting and ansering Question2
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion1']")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion1']")).SendKeys(Keys.Down + Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 2:')]//following::input[1]")).SendKeys("1");

            //Selecting and answering Question 3
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion2']")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion2']")).SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 3:')]//following::input[1]")).SendKeys("1");
            Verification.Sleep();
            ChromeDriver.FindElement(By.XPath("//input[@value='Continue']")).Click();
            Verification.Sleep();
            Verification.Sleep();
            //Verifying that the Dashboard is present
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[@class='ng-binding']")).Displayed);
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='SETTINGS']")).Click();
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='LOGOFF']")).Click();
            ChromeDriver.Close();
        }


        //Registration Process in Incognito with Practice ID
        [Priority(10)]
        [TestMethod]
        public void IncognitoRegistrationWithPracticeID()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("headless");
            options.AddArguments("--incognito");
            IWebDriver ChromeDriver = new ChromeDriver(options);
            ChromeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com/#/register?practiceID={_practicID}");
            ChromeDriver.Manage().Window.Maximize();
            Verification.Sleep();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//input[@id='Username']")).Displayed);
            Console.WriteLine("Registration page is confirmed");
            // Calling the function RegistrationProcess
            string strName = Verification.GenerateRandomString();
            //Generating a random string for the username starting with SeleniumTest
            ChromeDriver.FindElement(By.XPath("//input[@id='Username']")).SendKeys(strName);
            ChromeDriver.FindElement(By.XPath("//input[@id='Password']")).SendKeys("Password.1");
            ChromeDriver.FindElement(By.XPath("//input[@id='ConfirmPassword']")).SendKeys("Password.1");
            //if (ChromeDriver.FindElement(By.XPath("//input[@id='PracticeID']")).Displayed)
            //{
            //    ChromeDriver.FindElement(By.XPath("//input[@id='PracticeID']")).SendKeys(_practicID);
            //}
            ChromeDriver.FindElement(By.XPath("//input[@id='FirstName']")).SendKeys("Smoke");
            ChromeDriver.FindElement(By.XPath("//input[@id='LastName']")).SendKeys("Test");
            ChromeDriver.FindElement(By.XPath("//input[@name='DateOfBirth']")).SendKeys("10101985");
            ChromeDriver.FindElement(By.XPath("//input[@id='ZipCode']")).SendKeys("33609");
            ChromeDriver.FindElement(By.XPath("//input[@id='SecurityCode']")).SendKeys("433723727");
            ChromeDriver.FindElement(By.XPath("//input[@type='checkbox']")).Click();
            ChromeDriver.FindElement(By.XPath("//input[@value='Create Account']")).Click();
            Verification.Sleep();
            //Confirming if the Registration is successful by confirming the Forget password page is available
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(),'JUST IN CASE')]")).Displayed);
            //Selecting and answering Question1
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion0']")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion0']")).SendKeys(Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 1:')]//following::input[1]")).SendKeys("1");

            //Selecting and ansering Question2
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion1']")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion1']")).SendKeys(Keys.Down + Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 2:')]//following::input[1]")).SendKeys("1");

            //Selecting and answering Question 3
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion2']")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion2']")).SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 3:')]//following::input[1]")).SendKeys("1");
            Verification.Sleep();
            ChromeDriver.FindElement(By.XPath("//input[@value='Continue']")).Click();
            Verification.Sleep();
            Verification.Sleep();
            //Verifying that the Dashboard is present
            // Assert.IsTrue(ChromeDrivers.FindElement(By.XPath("//h2[@class='ng-binding']")).Displayed);
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='SETTINGS']")).Click();
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='LOGOFF']")).Click();
            ChromeDriver.Close();

        }

        //Registration Process in Incognito WITHOUT Practice ID
        [Priority(11)]
        [TestMethod]
        public void IncognitoRegistrationWithoutPracticeID()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("headless");
            options.AddArguments("--incognito");
            IWebDriver ChromeDriver = new ChromeDriver(options);
            ChromeDriver.Navigate().GoToUrl($"https://{_baseSite}.mypatientvisit.com/#/register?practiceID={_practicID}");
            ChromeDriver.Manage().Window.Maximize();
            Verification.Sleep();
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//input[@id='Username']")).Displayed);
            Console.WriteLine("Registration page is confirmed");
            WriteFile.WriteToFile("Registration page is confirmed - Incognito");
            // Calling the function RegistrationProcess
            string strName = Verification.GenerateRandomString();
            //Generating a random string for the username starting with SeleniumTest
            ChromeDriver.FindElement(By.XPath("//input[@id='Username']")).SendKeys(strName);
            ChromeDriver.FindElement(By.XPath("//input[@id='Password']")).SendKeys("Password.1");
            ChromeDriver.FindElement(By.XPath("//input[@id='ConfirmPassword']")).SendKeys("Password.1");
            //if (ChromeDriver.FindElement(By.XPath("//input[@id='PracticeID']")).Displayed)
            //{
            //    ChromeDriver.FindElement(By.XPath("//input[@id='PracticeID']")).SendKeys(_practicID);
            //}
            ChromeDriver.FindElement(By.XPath("//input[@id='FirstName']")).SendKeys("Smoke");
            ChromeDriver.FindElement(By.XPath("//input[@id='LastName']")).SendKeys("Test");
            ChromeDriver.FindElement(By.XPath("//input[@name='DateOfBirth']")).SendKeys("10101985");
            ChromeDriver.FindElement(By.XPath("//input[@id='ZipCode']")).SendKeys("33609");
            ChromeDriver.FindElement(By.XPath("//input[@id='SecurityCode']")).SendKeys("433723727");
            ChromeDriver.FindElement(By.XPath("//input[@type='checkbox']")).Click();
            ChromeDriver.FindElement(By.XPath("//input[@value='Create Account']")).Click();
            Verification.Sleep();
            Verification.Sleep();
            WriteFile.WriteToFile("Secutiry Questions Page is verified - Incongito");
            //Confirming if the Registration is successful by confirming the Forget password page is available
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(),'JUST IN CASE')]")).Displayed);
            //Selecting and answering Question1
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion0']")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion0']")).SendKeys(Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 1:')]//following::input[1]")).SendKeys("1");

            //Selecting and ansering Question2
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion1']")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion1']")).SendKeys(Keys.Down + Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 2:')]//following::input[1]")).SendKeys("1");

            //Selecting and answering Question 3
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion2']")).Click();
            ChromeDriver.FindElement(By.XPath("//select[@name='securityQuestion2']")).SendKeys(Keys.Down + Keys.Down + Keys.Down + Keys.Enter);
            ChromeDriver.FindElement(By.XPath("//fieldset//child::legend[contains(text(), 'Question 3:')]//following::input[1]")).SendKeys("1");
            Verification.Sleep();
            ChromeDriver.FindElement(By.XPath("//input[@value='Continue']")).Click();
            Verification.Sleep();
            Verification.Sleep();
            //Verifying that the Dashboard is present
            Assert.IsTrue(ChromeDriver.FindElement(By.XPath("//h2[contains(text(), 'Welcome')]")).Displayed);
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='SETTINGS']")).Click();
            ChromeDriver.FindElement(By.XPath("//span[@translate-once='LOGOFF']")).Click();
            ChromeDriver.Close();

            WriteFile.WriteToFile("All the test have passed");
            //Verification.Email_send();
        }

    }

}

