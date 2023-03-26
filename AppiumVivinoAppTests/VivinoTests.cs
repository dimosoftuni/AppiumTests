using NUnit.Framework;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Android;
using System.ComponentModel.DataAnnotations;
using System.Globalization;

namespace AppiumVivinoAppTests
{
    public class VivinoTests
    {
        private const string UriString = "http://127.0.0.1:4723/wd/hub";
        private const string VivinoAppLocation = @"C:\DemoApps\vivino_8.18.11-8181203.apk";
        private AndroidDriver<AndroidElement> driver;
        private AppiumOptions options;

        [SetUp]
        public void PrepareApp()
        {
            this.options = new AppiumOptions() { PlatformName = "Android" };
            options.AddAdditionalCapability("app", VivinoAppLocation);
            options.AddAdditionalCapability("appPackage", "vivino.web.app");
            options.AddAdditionalCapability("appActivity", "com.sphinx_solution.activities.SplashActivity");
            // This will execute tests on the currently running emulator
            options.AddAdditionalCapability("udid", "emulator-5554");
            // This will execute tests on the attached device
            //options.AddAdditionalCapability("udid", "R9YTB04CQ3Z");
            this.driver = new AndroidDriver<AndroidElement> (new Uri(UriString), options);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
        }

        [TearDown]
        public void CloseApp()
        {
            driver.Quit();
        }

        [Test]
        public void Test_SearchWine_VerifyNameAndRating()
        {

            var linkAccount = driver.FindElementById("vivino.web.app:id/txthaveaccount");
            linkAccount.Click();

            var inputUsername = driver.FindElementById("vivino.web.app:id/edtEmail");
            inputUsername.SendKeys("pesho@abv.bg");

            var inputPassword = driver.FindElementById("vivino.web.app:id/edtPassword");
            inputPassword.SendKeys("parola1parola1");

            var linkLogin = driver.FindElementById("vivino.web.app:id/action_signin");
            linkLogin.Click();

            var buttonSearch = driver.FindElementById("vivino.web.app:id/wine_explorer_tab");
            buttonSearch.Click();

            var searchArea = driver.FindElementById("vivino.web.app:id/search_header_text");
            searchArea.Click();

            var inputSearchButton = driver.FindElementById("vivino.web.app:id/editText_input");
            inputSearchButton.SendKeys("Katarzyna Reserve Red 2006");

            driver.PressKeyCode(66);

            var listWineResultElement = driver.FindElementById("vivino.web.app:id/winename_textView");
            listWineResultElement.Click();

            var labelWineName = driver.FindElementById("vivino.web.app:id/wine_name");
            var labelRatingText = driver.FindElementById("vivino.web.app:id/rating").Text;
            var rating = double.Parse(labelRatingText, CultureInfo.InvariantCulture);

            // var labelHighlights = driver.FindElementById("vivino.web.app:id/highlight_description");

            var labelHighlights = driver.FindElementByAndroidUIAutomator(
                "new UiScrollable(new UiSelector().scrollable(true))" +
                ".scrollIntoView(new UiSelector().resourceIdMatches(" +
                "\"vivino.web.app:id/highlight_description\"))"
            );



            Assert.That(labelWineName.Text, Is.EqualTo("Reserve Red 2006"));
            Assert.That(rating >= 1.00 && rating <= 5.00);
            Assert.That(labelHighlights.Text, Is.EqualTo("Among top 1% of all wines in the world"));

            var sectionSummary = driver.FindElementById("vivino.web.app:id/tabs");
            var linkFacts = sectionSummary.FindElementByXPath("//android.widget.TextView[2]");
            //var linkFacts = driver.FindElementByXPath("//android.widget.HorizontalScrollView/android.widget.LinearLayout/android.widget.TextView[2]");
            linkFacts.Click();

            var labelFactTitle = driver.FindElementById("vivino.web.app:id/wine_fact_title");
            Assert.That(labelFactTitle.Text, Is.EqualTo("Grapes"));

            var labelFactText = driver.FindElementById("vivino.web.app:id/wine_fact_text");

            Assert.That(labelFactText.Text, Is.EqualTo("Cabernet Sauvignon,Merlot"));

        }
    }
}