using NUnit.Framework;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium;
using System.Drawing;
using System.Xml.Linq;

namespace AppiumArchive7ZipTest
{
    public class Tests7ZipArchiver
    {
        private const string AppiumUriString = "http://127.0.0.1:4723/wd/hub";
        private const string ZipLocation = @"C:\Program Files\7-Zip\7zFM.exe";
        private const string tempDirectory = @"C:\temp";
        private WindowsDriver<WindowsElement> driver;
        private WindowsDriver<WindowsElement> driverArchiveWindow;
        private AppiumOptions options;
        private AppiumOptions optionsArchiveWindow;

        [SetUp]
        public void OpenApp()
        {
            this.options = new AppiumOptions() { PlatformName = "Windows" };
            options.AddAdditionalCapability("app", ZipLocation);
            this.driver = new WindowsDriver<WindowsElement>(new Uri(AppiumUriString), options);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);


            this.optionsArchiveWindow = new AppiumOptions() { PlatformName = "Windows" };
            optionsArchiveWindow.AddAdditionalCapability("app", "Root");
            this.driverArchiveWindow = new WindowsDriver<WindowsElement>(new Uri(AppiumUriString), optionsArchiveWindow);
            driverArchiveWindow.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

            if (Directory.Exists(tempDirectory))
            {
                Directory.Delete(tempDirectory, true);
            }

            Directory.CreateDirectory(tempDirectory);
            Thread.Sleep(1000);
        }

        [TearDown]
        public void CloseApp()
        {
            driver.Quit();
        }

        [Test]
        public void Test_ArchiveFunctionality()
        {
            var inputFilePath = driver.FindElementByXPath("/Window/Pane/Pane/ComboBox/Edit");
            inputFilePath.SendKeys(@"C:\Program Files\7-Zip\" + Keys.Enter);

            var listFiles = driver.FindElementByClassName("SysListView32");
            listFiles.SendKeys(Keys.Control + "a");

            var buttonAdd = driver.FindElementByName("Add");
            buttonAdd.Click();

            var windowArchive = driverArchiveWindow.FindElementByName("Add to Archive");

            //var inputArchivePath = windowArchive.FindElementByXPath("//Edit[@Name='Archive:']");
            var inputArchivePath = windowArchive.FindElementByXPath("/Window/ComboBox/Edit[@Name='Archive:']");
            inputArchivePath.SendKeys(@"C:\temp\archive.7z");

            //var dropDownFieldArchiveFormat = windowArchive.FindElementByXPath("/Window/ComboBox[@Name='Archive format:']");
            var dropDownFieldArchiveFormat = windowArchive.FindElementByXPath("/Window/ComboBox[@Name='Archive format:']");
            dropDownFieldArchiveFormat.SendKeys("7z");

            //var dropDownFieldCompressionLevel = windowArchive.FindElementByName("Compression level:");
            var dropDownFieldCompressionLevel = windowArchive.FindElementByXPath("/Window/ComboBox[@Name='Compression level:']");
            dropDownFieldCompressionLevel.SendKeys("9 - Ultra");

            //var dropDownFieldCompressionMethod = windowArchive.FindElementByName("Compression method:");
            var dropDownFieldCompressionMethod = windowArchive.FindElementByXPath("/Window/ComboBox[@Name='Compression method:']");
            dropDownFieldCompressionMethod.SendKeys("*");

            //var buttonOk = windowArchive.FindElementByXPath("//Button[@Name='OK']");
            var buttonOk = windowArchive.FindElementByXPath("/Window/Button[@Name='OK']");
            buttonOk.Click();
            Thread.Sleep(1000);

            inputFilePath.SendKeys(tempDirectory + @"\archive.7z" + Keys.Enter);

            var buttonExtract = driver.FindElementByName("Extract");
            buttonExtract.Click();

            //var inputFieldCopyto = driver.FindElementByName("Copy to:");
            var inputFieldCopyto = driver.FindElementByXPath("/Window/Window/ComboBox/Edit[@Name='Copy to:']");
            inputFieldCopyto.SendKeys(tempDirectory + Keys.Enter);
            Thread.Sleep(1000);

            FileAssert.AreEqual(ZipLocation, @"C:\temp\7zFM.exe");
        }
    }
}