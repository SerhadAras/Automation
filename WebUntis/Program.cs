using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;

namespace WebUntis
{
    class Program
    {
        static void DownloadSixWeeks(string email, string pass, int classKey, string klas, string path)
        {
            var options = new ChromeOptions();
            options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
            IWebDriver driver = new ChromeDriver(options);
            driver.Navigate()
                .GoToUrl("https://arche.webuntis.com/WebUntis/?school=AP-Hogeschool-Antwerpen#/basic/login");

            driver.FindElement(By.XPath("//input[@type='text']")).SendKeys(email);
            driver.FindElement(By.XPath("//input[@type='password']")).SendKeys(pass);
            driver.FindElement(By.XPath("//button[@type='submit']")).Click();

            for (int i = 0; i < 6; i++)
            {
                DateTime week = DateTime.Today.AddDays(i * 7);
                string download = $"https://arche.webuntis.com/WebUntis/Ical.do?elemType=1&elemId={classKey}&rpt_sd=" +
                                  week.ToString("yyyy-MM-dd");
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(download);
            }

            Merger(driver, klas, path);
            EditIcs(path);
            Outlook(email, pass, driver);
        }

        static void Clean(string path)
        {
            var folder = new System.IO.DirectoryInfo(path);
            var files = folder.GetFiles($"*.ics");
            foreach (var item in files)
            {
                File.Delete(item.FullName);
            }
        }

        static void Merger(IWebDriver driver, string klas, string path)
        {
            driver.Navigate().GoToUrl("https://icsmerger.com/");

            //UPLOADING FILES FOR MERGER
            var inputFile = driver.FindElement(By.XPath("//*[@id='fileChooser']"));
            string filePath = path + $"{klas}.ics";
            for (int i = 0; i < 6; i++)
            {
                if (i != 0)
                    filePath = path + $"{klas} ({i}).ics";
                inputFile.SendKeys(filePath);
            }

            driver.FindElement(By.XPath("//*[@id='content']/button")).Click();
            Thread.Sleep(1000);

            //DELETING WEEKLY CALENDAR
            var folder = new System.IO.DirectoryInfo(path);
            var files = folder.GetFiles($"*.ics");

            foreach (var item in files)
            {
                if (item.Name != "calendar (6).ics")
                    File.Delete(item.FullName);
            }
        }

        static void Outlook(string email, string pass, IWebDriver driver)
        {
            driver.Navigate().GoToUrl("https://outlook.office365.com/calendar/addcalendar");
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("//*[@id='i0116']")).SendKeys(email);
            driver.FindElement(By.XPath("//*[@id='idSIButton9']")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("//*[@id='i0118']")).SendKeys(pass);
            driver.FindElement(By.XPath("//*[@id='idSIButton9']")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("//*[@id='idSIButton9']")).Click();

            Thread.Sleep(3000);
            driver.FindElement(By.XPath("//*[@id='ImportFromFile']/span")).Click();
            driver.FindElement(By.XPath("//span[contains(text(), 'Browse')]")).Click();
            Thread.Sleep(1000);
            Process.Start(Directory.GetCurrentDirectory()+@"\upload.exe"); 
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("//span[contains(text(), 'Select a calendar')]")).Click();
            driver.FindElement(By.XPath("//span[contains(text(), 'Calendar')]")).Click();
            driver.FindElement(By.XPath("//span[contains(text(), 'Import')]")).Click();
            Thread.Sleep(1000);
            driver.Close();
            driver.Quit();
        }

        static void EditIcs(string path)
        {
            string[] classes =
            {
                "1ITBUS1 ", "1ITSOF1 ", "1ITSOF2 ", "1ITSOF3 ", "1ITSOF4 ", "1ITCSC1 ", "1ITCSC2 ", "1EA_U2 ",
                "1I\n TSOF4 ", "1ITIOT1 ", "1TI_U2 "
            };

            string text = File.ReadAllText(path + "calendar (6).ics");

            foreach (var VARIABLE in classes)
            {
                text = text.Replace(VARIABLE, "");
            }

            File.WriteAllText(path + "calendar (6).ics", text);
        }

        static void Main(string[] args)
        {
            /*1ITBUS1: 8123   1ITCSC1: 8128    1ITCSC2: 8133      1ITIOT1: 8143  
              1ITSOF1: 8148   1ITSOF2: 8153    1ITSOF3: 8158      1ITSOF4: 9741*/


            string email = "snummer@ap.be"; //office365 account
            string pass = "pass"; //password
            string klas = "1ITSOF4"; //classname
            string path = @"D:\Downloads\"; // download folder path
            int classID = 9741; //classid

            Clean(path);
            DownloadSixWeeks(email, pass, classID, klas.ToUpper(), path);
        }
    }
}