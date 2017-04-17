using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;


namespace FormTestOtoCare
{
    class WebCapabilities
    {
        public static string Browser;
        public static IWebDriver driver;
        private string iniFileLocation = "C:\\Company AAB\\LiteTestingTools\\Configuration\\Settings.INI";

        public WebCapabilities()
        {
            Ini ini = new Ini(iniFileLocation);
            ini.Load();
            Browser = ini.GetValue("Browser", "WebCapabilities");
        }

        public void CapabilitiesDriver()
        {
            switch (Browser.ToUpper())
            {
                case "CHROME": driver = new ChromeDriver();
                    break;
                case "FIREFOX": 
                    var optionsFireFox = new FirefoxOptions();
                    optionsFireFox.BrowserExecutableLocation = @"C:\Program Files (x86)\Mozilla Firefox\firefox.exe";
                    driver = new FirefoxDriver(optionsFireFox);
                    break;
                case "IE":
                    var options = new InternetExplorerOptions { RequireWindowFocus = true, EnablePersistentHover = false };
                    options.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
                    options.IgnoreZoomLevel = true;
                    //options.UnexpectedAlertBehavior = InternetExplorerUnexpectedAlertBehavior.Ignore;
                    driver = new InternetExplorerDriver(options);
                    break;
            }
        }
    }
}
