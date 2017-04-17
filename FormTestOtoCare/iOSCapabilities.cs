using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.iOS;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FormTestOtoCare
{
    class iOSCapabilities
    {
        public static DesiredCapabilities _capabilities;
        public static IOSDriver<IOSElement> driver, tempDriver1;
        private static string platformName;
        private static string automationName;
        private static string platform;
        private static string deviceName;
        private static string platformVersion;
        private static string app;
        private static string serverUri1;
        private string iniFileLocation = "C:\\Company AAB\\LiteTestingTools\\Configuration\\Settings.INI";

        public iOSCapabilities()
        {
            Ini ini = new Ini(iniFileLocation);
            ini.Load();

            platformName = ini.GetValue("platformName", "IOSCapabilities");
            platformVersion = ini.GetValue("platformVersion", "IOSCapabilities");
            platform = ini.GetValue("platform", "IOSCapabilities");
            deviceName = ini.GetValue("deviceName", "IOSCapabilities");
            app = ini.GetValue("app", "IOSCapabilities");
            serverUri1 = ini.GetValue("serverUri1", "IOSCapabilities");
        }

        public void CapabilitiesDevice()
        {
            _capabilities = new DesiredCapabilities();
            _capabilities.SetCapability("platformName", platformName.ToString());
            _capabilities.SetCapability("platformVersion", platformVersion.ToString());
            _capabilities.SetCapability("platform", platform.ToString());
            _capabilities.SetCapability("deviceName", deviceName.ToString());
            _capabilities.SetCapability("app", app.ToString());
            driver = new IOSDriver<IOSElement>(new Uri(serverUri1), _capabilities, TimeSpan.FromSeconds(1000));
            tempDriver1 = driver;
            //driver = new AppiumDriver<RemoteWebDriver>(new Uri(serverUri), _capabilities, TimeSpan.FromSeconds(180));
        }
    }

}
