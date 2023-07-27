using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using RPA.Models;
using System.Diagnostics;

namespace RPA.program

{
    public static class Program
    {
        public static ChromeDriver? _driver;
        public static IWebElement? _element;

        
        public static void Main()
        {
            //verificando se já existe o arquivo do site baixado. Caso sim, exclua.
            string downloadPath = @"C:\Users\fiska\Desktop\projetos\RPAChallengeSelenium\RPA\Excel";            
            string excelFile = "challenge.xlsx";            
            string fullPath = Path.Combine(downloadPath, excelFile);
            if (File.Exists(fullPath))
            {                
                File.Delete(fullPath);
            }

            // alterando as definições de salvamento de downloads
            ChromeOptions options = new ChromeOptions();
            options.AddUserProfilePreference("download.default_directory", @"C:\Users\fiska\Desktop\projetos\RPAChallengeSelenium\RPA\Excel");
            options.AddUserProfilePreference("prompt_for_download", false); //sobreposição de arquivos download

            // configurando acesso à pagina
            _driver = new ChromeDriver(options);
            _driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
            _driver.Navigate().GoToUrl("https://www.rpachallenge.com/");

            // download do arquivo com os dados
            var excelDownload = _driver.FindElement(By.CssSelector("a.btn[target='_blank']"));
            excelDownload.Click();

            Thread.Sleep(3000);

            var begin = _driver.FindElement(By.TagName("button"));
            begin.Click();     
            
            InputData();

            string screenshotPath = @"C:\Users\fiska\Desktop\projetos\RPAChallengeSelenium\RPA\screenshot\print.png";

            // Salvando o resultado
            Screenshot screenshot = ((ITakesScreenshot)_driver).GetScreenshot();           
            screenshot.SaveAsFile(screenshotPath, ScreenshotImageFormat.Png);

            Thread.Sleep(3000);

            _driver.Quit();


            //abrindo printscreen com resultado           
            if (File.Exists(screenshotPath))
            {
                
                Process.Start(new ProcessStartInfo(screenshotPath) { UseShellExecute = true });
            }

         

        }

        private static List<PersonModel> ExcelRead()
        {
            ExcelDataReader excelFileDataReader = new ExcelDataReader();
            var path = @"C:\Users\fiska\Desktop\projetos\RPAChallengeSelenium\RPA\Excel\challenge.xlsx";
            return excelFileDataReader.ReadingSpreadsheet(path);
        }


        private static void InputData()
        {

            var personList = ExcelRead();

            for (int i = 0; i <= personList.Count - 1; i++)
            {
                var inputList = _driver.FindElements(By.TagName("input"));
                for (int j = 0; j < inputList.Count - 1; j++)
                {
                    if (inputList[j].GetDomAttribute("ng-reflect-name").Contains("FirstName"))
                    {
                        inputList[j].SendKeys(personList[i].FirstName);
                        continue;
                    }

                    if (inputList[j].GetDomAttribute("ng-reflect-name").ToUpper().Contains("LASTNAME"))
                    {
                        inputList[j].SendKeys(personList[i].LastName);
                        continue;
                    }

                    if (inputList[j].GetDomAttribute("ng-reflect-name").ToUpper().Contains("COMPANY"))
                    {
                        inputList[j].SendKeys(personList[i].CompanyName);
                        continue;
                    }


                    if (inputList[j].GetDomAttribute("ng-reflect-name").ToUpper().Contains("ROLE"))
                    {
                        inputList[j].SendKeys(personList[i].RoleInCompany);
                        continue;
                    }


                    if (inputList[j].GetDomAttribute("ng-reflect-name").ToUpper().Contains("ADDRESS"))
                    {
                        inputList[j].SendKeys(personList[i].Address);
                        continue;
                    }

                    if (inputList[j].GetDomAttribute("ng-reflect-name").ToUpper().Contains("EMAIL"))
                    {
                        inputList[j].SendKeys(personList[i].Email);
                        continue;
                    }

                    if (inputList[j].GetDomAttribute("ng-reflect-name").ToUpper().Contains("PHONE"))
                    {
                        inputList[j].SendKeys(personList[i].PhoneNumber);
                        continue;
                    }
                }

                _element = inputList.ElementAt(inputList.Count - 1);
                _element.Click();
            }
        }


    }
}