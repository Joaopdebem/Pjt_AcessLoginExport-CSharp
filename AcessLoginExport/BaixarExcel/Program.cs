﻿using OpenQA.Selenium;
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
            driver.Navigate().GoToUrl("http://34.235.5.35:8080/los-dashboard");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            Thread.Sleep(2000);
            
            var timeout = TimeSpan.FromSeconds(10);
            var pollInterval = TimeSpan.FromMilliseconds(500);
            var startTime = DateTime.Now;

            IWebElement element1 = null;

            while (DateTime.Now - startTime < timeout)
            {
                try
                {
                    element1 = driver.FindElement(By.CssSelector("#login\\:name"));
                    if (element1.Displayed)
                    {
                        break;
                    }
                }
                catch (NoSuchElementException){}
                catch (StaleElementReferenceException){}

                Thread.Sleep(pollInterval);
            }

            if (element1 != null)
            {
                element1.Click();
                element1.SendKeys("wagner.reis");
            }

            startTime = DateTime.Now;
            IWebElement element2 = null;

            while (DateTime.Now - startTime < timeout)
            {
                try
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
                element2.SendKeys("W4gn3r#08");
                element2.SendKeys(Keys.Enter);
            }
            Thread.Sleep(3000);

            driver.Navigate().GoToUrl("http://34.235.5.35:8080/los-dashboard/products/expirationDate.xhtml");
            wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            Thread.Sleep(3000);
            
            ((IJavaScriptExecutor)driver).ExecuteScript("console.clear();");

            string javascriptCommand = "mojarra.jsfcljs(document.getElementById('expirationForm'), {'expirationForm:expirationDataTableId:j_idt404':'expirationForm:expirationDataTableId:j_idt404'}, '');";

            ((IJavaScriptExecutor)driver).ExecuteScript(javascriptCommand);
            Thread.Sleep(5000);        

            string sourcePath = @"C:\Users\jpdeb\Downloads\";
            string targetPath = @"C:\Users\jpdeb\OneDrive\Área de Trabalho\Aviso Encapsulados";

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
