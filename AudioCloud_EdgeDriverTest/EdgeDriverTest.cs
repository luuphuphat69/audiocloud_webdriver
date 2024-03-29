using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading.Tasks;
using System;
using static System.Net.WebRequestMethods;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Security.Cryptography;
using IdentityModel.Client;
using AudioCloud_EdgeDriverTest;
using OpenQA.Selenium.Support.UI;
using System.Security.Policy;
using System.Threading;
using OpenQA.Selenium.Interactions;



namespace AudioCloud_ChromeDriverTest
{
    [TestClass]
    public class ChromeDriverTest
    {
        private ChromeDriver _driver;
        public static string url = "http://54.161.251.210:8080";
        private static int rowIndex = 1;

        [TestInitialize]
        public void ChromeDriverInitialize()
        {
            // Initialize Chrome driver 
            var options = new ChromeOptions
            {
                PageLoadStrategy = PageLoadStrategy.Normal
            };
            _driver = new ChromeDriver(options);
        }


        [TestCleanup]
        public void ChromeDriverCleanup()
        {
            _driver.Quit();
        }

        // Update Test Status in Excel
        private void UpdateTestResultInExcel(string sheetName, int row, int column, string result)
        {
            using (var excelPackage = new ExcelPackage(new System.IO.FileInfo("DataTest.xlsx")))
            {
                var worksheet = excelPackage.Workbook.Worksheets[sheetName];
                worksheet.Cells[row, column].Value = result;
                excelPackage.Save();
            }
        }

        // Login data
        public static IEnumerable<object[]> ReadTestDataLoginFromExcel()
        {
            var testData = new List<object[]>();

            using (var excelPackage = new ExcelPackage(new System.IO.FileInfo("DataTest.xlsx")))
            {
                var worksheet = excelPackage.Workbook.Worksheets["LoginDataTest"];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 1; row <= rowCount; row++)
                {
                    var account = worksheet.Cells[row, 1].Value?.ToString();
                    var password = worksheet.Cells[row, 2].Value?.ToString();
                    testData.Add(new object[] { account, password });
                }
            }

            return testData;
        }

        // SignUp data
        public static IEnumerable<object[]> ReadTestDataSignUpFromExcel()
        {
            var testData = new List<object[]>();

            using (var excelPackage = new ExcelPackage(new System.IO.FileInfo("DataTest.xlsx")))
            {
                var worksheet = excelPackage.Workbook.Worksheets["SignUpDataTest"];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 1; row <= rowCount; row++)
                {
                    var account = worksheet.Cells[row, 1].Value?.ToString();
                    var password = worksheet.Cells[row, 2].Value?.ToString();
                    var repassword = worksheet.Cells[row, 3].Value?.ToString();
                    var email = worksheet.Cells[row, 4].Value?.ToString();
                    testData.Add(new object[] { account, password, repassword, email });
                }
            }

            return testData;
        }

        // Search data 

        public static IEnumerable<object[]> ReadTestDataSearchFromExcel()
        {
            var testData = new List<object[]>();

            using (var excelPackage = new ExcelPackage(new System.IO.FileInfo("DataTest.xlsx")))
            {
                var worksheet = excelPackage.Workbook.Worksheets["SearchDataTest"];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 1; row <= rowCount; row++)
                {
                    var query = worksheet.Cells[row, 1].Value?.ToString();
                    testData.Add(new object[] { query });
                }
            }

            return testData;
        }

        // Edit Information 
        public static IEnumerable<object[]> ReadTestDataEditInformationFromExcel()
        {
            var testData = new List<object[]>();

            using (var excelPackage = new ExcelPackage(new System.IO.FileInfo("DataTest.xlsx")))
            {
                var worksheet = excelPackage.Workbook.Worksheets["EditInformationDataTest"];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 1; row <= rowCount; row++)
                {   
                    var displayname = worksheet.Cells[row, 1].Value?.ToString();
                    var address = worksheet.Cells[row, 2].Value?.ToString();
                    var bio = worksheet.Cells[row, 3].Value?.ToString();
                    var photo = worksheet.Cells[row, 4].Value?.ToString();

                    testData.Add(new object[] { displayname, address, bio , photo});
                }
            }

            return testData;
        }

        private void UpdateAddAudioToFavouriteTestResultInExcel(int row, string result)
        {
            using (var excelPackage = new ExcelPackage(new System.IO.FileInfo("DataTest.xlsx")))
            {
                var worksheet = excelPackage.Workbook.Worksheets["AddAudioToFavouriteDataTest"];
                worksheet.Cells[row, 1].Value = result; 
                excelPackage.Save();
            }
        }

        [TestMethod]
        [DataRow("http://54.161.251.210:8080/home")]
        public void URLTest(string url)
        {
            _driver.Url = url;
            Assert.AreEqual(url, _driver.Url);
        }

        [TestMethod]
        [DynamicData(nameof(ReadTestDataLoginFromExcel), DynamicDataSourceType.Method)]
        public async Task LoginTest(string account, string password)
        {
            _driver.Url = url + "/login";

            _driver.FindElement(By.Id("account")).SendKeys(account);
            _driver.FindElement(By.Id("password")).SendKeys(password);

            _driver.FindElement(By.Id("submit")).Submit();
            await Task.Delay(TimeSpan.FromSeconds(5));

            IAlert alert = _driver.SwitchTo().Alert();
            alert.Accept();

            try
            {
                Assert.IsTrue(_driver.Url.Contains("/home"), "Login failed. User was not redirected to the home page.");
                UpdateTestResultInExcel("LoginDataTest", rowIndex, 3, "Passed");
            }
            catch (AssertFailedException ex)
            {
                UpdateTestResultInExcel("LoginDataTest", rowIndex, 3, "Failed");
                throw ex;
            }
            finally
            {
                rowIndex++;
            }
        }

        [TestMethod]
        [DynamicData(nameof(ReadTestDataSignUpFromExcel), DynamicDataSourceType.Method)]
        public async Task SignUpTest(string account, string password, string repassword, string email)
        {
            _driver.Url = url + "/register";

            _driver.FindElement(By.Name("account")).SendKeys(account);
            _driver.FindElement(By.Name("password")).SendKeys(password);
            _driver.FindElement(By.Name("repassword")).SendKeys(repassword);
            _driver.FindElement(By.Name("email")).SendKeys(email);

            _driver.FindElement(By.Name("submit")).Submit();
            await Task.Delay(TimeSpan.FromSeconds(5));

            IAlert alert = _driver.SwitchTo().Alert();
            alert.Accept();

            try
            {
                Assert.IsTrue(_driver.Url.Contains("/login"), "Sign up failed");
                UpdateTestResultInExcel("SignUpDataTest", rowIndex, 6, "Passed");
            }
            catch (AssertFailedException ex)
            {
                UpdateTestResultInExcel("SignUpDataTest", rowIndex, 6, "Failed");
                throw ex;
            }
            finally
            {
                rowIndex++;
            }
        }

        [TestMethod]
        [DynamicData(nameof(ReadTestDataSearchFromExcel), DynamicDataSourceType.Method)]
        public async Task SearchSongs(string query)
        {
            _driver.Url = url + "/home";
            var searchInput = _driver.FindElement(By.CssSelector("input.form-control.pl-3"));
            searchInput.SendKeys(query);
            searchInput.Submit();
            await Task.Delay(TimeSpan.FromSeconds(5));

            try
            {
                Assert.IsTrue(_driver.Url.Contains("/search"), "Search failed");
                UpdateTestResultInExcel("SearchDataTest", rowIndex, 2, "Passed");   
            }
            catch (AssertFailedException ex)
            {
                UpdateTestResultInExcel("SearchDataTest", rowIndex, 2, "Failed");
                throw ex;
            }
            finally
            {
                rowIndex++;
            }
        }

        [TestMethod]
        [DynamicData(nameof(ReadTestDataEditInformationFromExcel), DynamicDataSourceType.Method)]

        public async Task EditInformation(string displayname, string address, string bio, string photo)
        {
                       
                _driver.Url = url + "/login";

                _driver.FindElement(By.Id("account")).SendKeys("myaccount");
                _driver.FindElement(By.Id("password")).SendKeys("123123");

                _driver.FindElement(By.Id("submit")).Submit();
                await Task.Delay(TimeSpan.FromSeconds(5));

                IAlert alert = _driver.SwitchTo().Alert();
                alert.Accept();

                _driver.Url = url + "/profile";
                await Task.Delay(TimeSpan.FromSeconds(5));

                var editInput = _driver.FindElement(By.CssSelector("Span.edit-icon"));
                editInput.Click();
                await Task.Delay(TimeSpan.FromSeconds(5));

                var displayNameElement = _driver.FindElement(By.Id("displayname"));
                var addressElement = _driver.FindElement(By.Id("address"));
                var bioElement = _driver.FindElement(By.Id("bio"));
                
                //Clear
                displayNameElement.Clear();
                addressElement.Clear();
                bioElement.Clear();
                //photoElement.Clear();
                               
                await Task.Delay(TimeSpan.FromSeconds(2));

                displayNameElement.SendKeys(displayname);
                addressElement.SendKeys(address);
                bioElement.SendKeys(bio);
                _driver.FindElement(By.Id("UserPhoto")).SendKeys("C:\\Users\\ACER-PC\\Downloads\\cloud.jpg");
                //photoElement.SendKeys(photo);

                var buttonInput = _driver.FindElement(By.CssSelector("button.btn.btn-primary.mt-3"));
                buttonInput.Click();
                await Task.Delay(TimeSpan.FromSeconds(5));

                alert.Accept();

                // Refresh the page to avoid stale elements
                _driver.Navigate().Refresh();
           
                // Re-locate elements after page refresh
                displayNameElement = _driver.FindElement(By.Id("displayname"));
                addressElement = _driver.FindElement(By.Id("address"));
                bioElement = _driver.FindElement(By.Id("bio"));
                

                var _displayName = displayNameElement.GetAttribute("value").ToString();
                var _address = addressElement.GetAttribute("value").ToString();
                var _bio = bioElement.GetAttribute("value").ToString();

                var firstCondition = _displayName.Equals(displayname);
                var secondCondition = _address.Equals(address);
                var thirdCondition = _bio.Equals(bio);
            try 
            { 

                Assert.IsTrue(firstCondition || secondCondition || thirdCondition, "Failed");
                UpdateTestResultInExcel("EditInformationDataTest", rowIndex, 4, "Pass");
            }
            catch (Exception e)
            {
                Assert.Fail(e.ToString());
                UpdateTestResultInExcel("EditInformationDataTest", rowIndex, 4, "Failed");
            }
        }

        

        //Xóa bài hát 

        [TestMethod]
        public async Task DeleteSong()
        {
            _driver.Url = url + "/login";

            _driver.FindElement(By.Id("account")).SendKeys("myaccount");
            _driver.FindElement(By.Id("password")).SendKeys("123123");

            _driver.FindElement(By.Id("submit")).Submit();
            await Task.Delay(TimeSpan.FromSeconds(5));

            IAlert alert = _driver.SwitchTo().Alert();
            alert.Accept();

            _driver.Url = url + "/profile";
            await Task.Delay(TimeSpan.FromSeconds(5));
            //*[@id="long-menu"]/div[3]/ul/li[2]

            var chamInput = _driver.FindElement(By.CssSelector("#long-button > svg"));
            chamInput.Click();
            Thread.Sleep(1000);

            Actions actions = new Actions(_driver);
            actions.SendKeys(Keys.End).Perform();            
            Thread.Sleep(5000);

            var deleteInput = _driver.FindElement(By.CssSelector("#long-menu > div.MuiPaper-root.MuiPaper-elevation.MuiPaper-rounded.MuiPaper-elevation8.MuiMenu-paper.MuiPopover-paper.MuiMenu-paper.css-3dzjca-MuiPaper-root-MuiPopover-paper-MuiMenu-paper > ul > li:nth-child(2)"));
            deleteInput.Click();
            await Task.Delay(TimeSpan.FromSeconds(5));
            alert.Accept();
            await Task.Delay(TimeSpan.FromSeconds(5));
        }



        [TestMethod]
        public async Task SignOut()
        {
            _driver.Url = url + "/login";

            _driver.FindElement(By.Id("account")).SendKeys("myaccount");
            _driver.FindElement(By.Id("password")).SendKeys("123123");

            _driver.FindElement(By.Id("submit")).Submit();
            await Task.Delay(TimeSpan.FromSeconds(5));

            IAlert alert = _driver.SwitchTo().Alert();
            alert.Accept();
            
            await Task.Delay(TimeSpan.FromSeconds(3));
            _driver.Manage().Window.Maximize();
            await Task.Delay(TimeSpan.FromSeconds(5));

            var profileInput = _driver.FindElement(By.XPath("//*[@id=\"ftco-navbar\"]/div/form/div/button"));
            profileInput.Click();
            Thread.Sleep(2000);

            var logoutInput = _driver.FindElement(By.XPath("//*[@id=\"navbar-links\"]/li/div/a[4]"));
            logoutInput.Click();
            Thread.Sleep(2000);

        }
        [TestMethod]
        // Thêm vào danh sách yêu thích 
        public async Task AddAudioToFavourite()
        {
            _driver.Url = url + "/login";

            _driver.FindElement(By.Id("account")).SendKeys("myaccount");
            _driver.FindElement(By.Id("password")).SendKeys("123123");

            _driver.FindElement(By.Id("submit")).Submit();
            await Task.Delay(TimeSpan.FromSeconds(3));

            IAlert alert = _driver.SwitchTo().Alert();
            alert.Accept();            
            _driver.Manage().Window.Maximize();
            Thread.Sleep(3000);
            
            _driver.Url = url + "/details/AUDIO10448585";
            Thread.Sleep(2000);

            var playlistButton = _driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/button[2]/div[2]"));
            playlistButton.Click();
            Thread.Sleep(2000);

            var addplaylistButton = _driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[4]/div[1]/div[1]/div[1]/ul[1]/div[1]/div[1]/div[1]/button[1]"));
            addplaylistButton.Click();
            Thread.Sleep(3000);          

            _driver.Url = url + "/profile";
            Thread.Sleep(3000);

            var danhsachphatButton = _driver.FindElement(By.CssSelector("#root > div > div.card.card-body.mx-3.mx-md-4.mt-n6 > div.row.gx-4.mb-2 > div.col-md-6.col-md-6.my-sm-auto.ms-sm-auto.me-sm-0.mx-auto.mt-3 > div > ul > li:nth-child(2) > a"));
            danhsachphatButton.Click();
            Thread.Sleep(3000);

            var danhsachplalistButton = _driver.FindElement(By.CssSelector("Capa_1"));
            danhsachplalistButton.Click();
            Thread.Sleep(3000);

            try
            {
                Assert.IsTrue(_driver.Url.Contains("/profile"), "Yêu thích thành công ");
                UpdateAddAudioToFavouriteTestResultInExcel(rowIndex, "Passed");
            }
            catch (AssertFailedException ex)
            {
                UpdateAddAudioToFavouriteTestResultInExcel(rowIndex, "Failed");
                throw ex;
            }
            finally
            {
                rowIndex++;
            }

        }

        [TestMethod]

        public async Task AddAudioToFavouriteAndCreate()
        {
            _driver.Url = url + "/login";

            _driver.FindElement(By.Id("account")).SendKeys("myaccount");
            _driver.FindElement(By.Id("password")).SendKeys("123123");

            _driver.FindElement(By.Id("submit")).Submit();
            await Task.Delay(TimeSpan.FromSeconds(3));

            IAlert alert = _driver.SwitchTo().Alert();
            alert.Accept();
            _driver.Manage().Window.Maximize();
            Thread.Sleep(3000);

            _driver.Url = url + "/details/AUDIO10448585";
            Thread.Sleep(2000);

            var playlistButton = _driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/button[2]/div[2]"));
            playlistButton.Click();
            Thread.Sleep(2000);

            var addplaylistButton = _driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[4]/div[1]/div[1]/div[1]/ul[1]/div[1]/div[1]/div[1]/button[1]"));
            addplaylistButton.Click();
            Thread.Sleep(3000);
                
            _driver.Url = url + "/profile";
            Thread.Sleep(3000);

            var danhsachphatButton = _driver.FindElement(By.CssSelector("#root > div > div.card.card-body.mx-3.mx-md-4.mt-n6 > div.row.gx-4.mb-2 > div.col-md-6.col-md-6.my-sm-auto.ms-sm-auto.me-sm-0.mx-auto.mt-3 > div > ul > li:nth-child(2) > a"));
            danhsachphatButton.Click();
            Thread.Sleep(3000);

            var danhsachplalistButton = _driver.FindElement(By.CssSelector("Capa_1"));
            danhsachplalistButton.Click();
            Thread.Sleep(3000);

            try
            {
                Assert.IsTrue(_driver.Url.Contains("/profile"), "Yêu thích thành công ");
                UpdateAddAudioToFavouriteTestResultInExcel(rowIndex, "Passed");
            }
            catch (AssertFailedException ex)
            {
                UpdateAddAudioToFavouriteTestResultInExcel(rowIndex, "Failed");
                throw ex;
            }
            finally
            {
                rowIndex++;
            }

        }

        [TestMethod]
        public async Task PlayAudio()
        {
            _driver.Url = url + "/home";
            _driver.FindElement(By.CssSelector("button.btn-95")).Submit();
            await Task.Delay(TimeSpan.FromSeconds(5));

            var aplayer = _driver.FindElement(By.Id("aplayer"));
            Assert.IsTrue(aplayer.Displayed, "Failed");

            IAlert alert = _driver.SwitchTo().Alert();
            alert.Accept();
        }

        

    }
}
