using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using RPAChallengeInputForm_JRTF.Models;
using System.IO;

namespace RPAChallengeInputForm_JRTF.Main
{
    public static class Init
    {
        private static ChromeDriver _driver;

        private static IWebElement _submitButton;

        public static void Main()
        {
            _driver = new ChromeDriver();
            _driver.Manage().Timeouts().PageLoad = TimeSpan.FromMinutes(2);
            _driver.Navigate().GoToUrl("https://www.rpachallenge.com/");

            var startButtom = _driver.FindElement(By.TagName("button"));
            startButtom.Click();

            InputData();
            SaveOutput();
        }

        /* O método realiza a chamada do método ReadEmployeeDataFromSpreadsheet da classe ExcelFileDataReader
         * e retorna a lista de informações dos funcionários.
         */
        private static List<EmployeeModel> ReadData()
        {
            ExcelFileDataReader excelFileDataReader = new ExcelFileDataReader();
            var path = @"Data\challenge.xlsx";
            return excelFileDataReader.ReadEmployeeDataFromSpreadsheet(path);
        }

        // O método realiza a inserção dos dados da lista nos respectivos campos do formulário. 
        private static void InputData()
        {
            var employeeList = ReadData();

            for (int i = 0; i <= employeeList.Count - 1; i++)
            {
                var inputList = _driver.FindElements(By.TagName("input"));
                for (int j = 0; j < inputList.Count - 1; j++)
                {
                    if (inputList[j].GetDomAttribute("ng-reflect-name").ToUpper().Contains("FIRSTNAME"))
                    {
                        inputList[j].SendKeys(employeeList[i].FirstName);
                        continue;
                    }

                    if (inputList[j].GetDomAttribute("ng-reflect-name").ToUpper().Contains("LASTNAME"))
                    {
                        inputList[j].SendKeys(employeeList[i].LastName);
                        continue;
                    }

                    if (inputList[j].GetDomAttribute("ng-reflect-name").ToUpper().Contains("COMPANY"))
                    {
                        inputList[j].SendKeys(employeeList[i].CompanyName);
                        continue;
                    }


                    if (inputList[j].GetDomAttribute("ng-reflect-name").ToUpper().Contains("ROLE"))
                    {
                        inputList[j].SendKeys(employeeList[i].RoleInCompany);
                        continue;
                    }


                    if (inputList[j].GetDomAttribute("ng-reflect-name").ToUpper().Contains("ADDRESS"))
                    {
                        inputList[j].SendKeys(employeeList[i].Address);
                        continue;
                    }

                    if (inputList[j].GetDomAttribute("ng-reflect-name").ToUpper().Contains("EMAIL"))
                    {
                        inputList[j].SendKeys(employeeList[i].Email);
                        continue;
                    }

                    if (inputList[j].GetDomAttribute("ng-reflect-name").ToUpper().Contains("PHONE"))
                    {
                        inputList[j].SendKeys(employeeList[i].PhoneNumber);
                        continue;
                    }
                }

                _submitButton = inputList.ElementAt(inputList.Count - 1);
                _submitButton.Click();
            }
        }

        // O método salva o resultado na pasta Documentos do usuário atual do Windows
        private static void SaveOutput()
        {
            StreamWriter sw = new StreamWriter(@$"C:\Users\{Environment.UserName}\Documents\Output.txt");

            var outputMessageTitle = _driver.FindElement(By.ClassName("message1"));
            sw.WriteLine(outputMessageTitle.Text);

            var outputMessageContent = _driver.FindElement(By.ClassName("message2"));
            sw.WriteLine(outputMessageContent.Text);

            sw.Close();
        }
    }
}