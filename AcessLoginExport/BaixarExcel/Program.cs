using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using System.Threading;
using System.IO;
using System;

class Program
{
    static void Main(string[] args)
    {
        var chromeDriverPath = @"C:\Dev\BaixarExcel\chromedriver.exe";

        var chromeOptions = new ChromeOptions();
        chromeOptions.AddArgument("--headless");
        chromeOptions.AddArgument("--disable-gpu");
        chromeOptions.AddArgument("--start-maximized");

        using (var driver = new ChromeDriver(chromeDriverPath, chromeOptions))
        {
            driver.Navigate().GoToUrl("http://34.235.5.35:8080/los-dashboard"); // url que deseja
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            Thread.Sleep(2000);

            var timeout = TimeSpan.FromSeconds(10);
            var pollInterval = TimeSpan.FromMilliseconds(500);
            var startTime = DateTime.Now;

            IWebElement element1 = null;

            while (DateTime.Now - startTime < timeout)
            {
                try // tenta preencher o campo login
                {
                    element1 = driver.FindElement(By.CssSelector("#login\\:name"));
                    if (element1.Displayed)
                    {
                        break;
                    }
                }
                catch (NoSuchElementException) { }
                catch (StaleElementReferenceException) { }

                Thread.Sleep(pollInterval);
            }

            if (element1 != null)
            {
                element1.Click();
                element1.SendKeys("teste");
            }

            startTime = DateTime.Now;
            IWebElement element2 = null;

            while (DateTime.Now - startTime < timeout)
            {
                try // tenta preencher o campo senha
                {
                    element2 = driver.FindElement(By.CssSelector("#login\\:password"));
                    if (element2.Displayed)
                    {
                        break;
                    }
                }
                catch (NoSuchElementException) { }
                catch (StaleElementReferenceException) { }

                Thread.Sleep(pollInterval);
            }

            if (element2 != null)
            {
                element2.Click();
                element2.SendKeys("123");
                element2.SendKeys(Keys.Enter);
            }
            Thread.Sleep(3000);

            driver.Navigate().GoToUrl("http://34.235.5.35:8080/los-dashboard/products/expirationDate.xhtml"); // navega para outra parte do site
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete")); // Aguarda carregar aquela p[agina do site
            Thread.Sleep(3000);

            ((IJavaScriptExecutor)driver).ExecuteScript("console.clear();");

            string javascriptCommand = "mojarra.jsfcljs(document.getElementById('expirationForm'), {'expirationForm:expirationDataTableId:j_idt404':'expirationForm:expirationDataTableId:j_idt404'}, '');"; // executa o classe que exporta o excel que preciso

            ((IJavaScriptExecutor)driver).ExecuteScript(javascriptCommand);
            Thread.Sleep(5000);

            string sourcePath = @"C:\"; // local que e baixado o arquivo
            string targetPath = @"C:\"; // local para onde quero que mova

            DirectoryInfo dir = new DirectoryInfo(sourcePath);

            FileInfo[] files = dir.GetFiles("Vencimentos_*");

            foreach (FileInfo file in files)
            {
                string destFile = Path.Combine(targetPath, "Vencimentos.xlsx");

                if (File.Exists(destFile))
                {
                    File.Delete(destFile);
                }

                File.Move(file.FullName, destFile);
            }

        }
    }
}
