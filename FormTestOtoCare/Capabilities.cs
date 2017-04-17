using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Appium.Android.Interfaces;
using OpenQA.Selenium.Appium.Enums;
using OpenQA.Selenium.Appium.Interfaces;
using OpenQA.Selenium.Appium.Service;
using OpenQA.Selenium.Appium;
using System.Threading;

namespace FormTestOtoCare
{
    class Capabilities
    {
        public static DesiredCapabilities _capabilities;
        public static AndroidDriver<AndroidElement> driver, tempDriver1;
        //public static IWebDriver Webdriver;
        public string lastElement;
        private static string deviceType;
        
        private static string deviceName;
        private static string appPlatform;
        private static string platformVersion;
        private static string appPackage;
        private static string serverUri1;
        public static string osVersion;
        public static string speedSwipe;
        private static string appActivity;
        private static string defaultBashPath;
        private string iniFileLocation = "C:\\Company AAB\\LiteTestingTools\\Configuration\\Settings.INI";

        public Capabilities()
        {
            Ini ini = new Ini(iniFileLocation);
            ini.Load();
            deviceType = ini.GetValue("deviceType", "AndroidCapabilities");
            deviceName = ini.GetValue("deviceName", "AndroidCapabilities");
            appPlatform = ini.GetValue("appPlatform", "AndroidCapabilities");
            platformVersion = ini.GetValue("platformVersion", "AndroidCapabilities");
            appPackage = ini.GetValue("appPackage", "AndroidCapabilities");
            osVersion = ini.GetValue("osVersion", "AndroidCapabilities");
            speedSwipe = ini.GetValue("speedSwipe", "AndroidCapabilities");

            //Tambahan untuk akomodir Multiple Devices
            serverUri1 = ini.GetValue("serverUri1", "AndroidCapabilities");
            appActivity = ini.GetValue("appActivity", "AndroidCapabilities");
            defaultBashPath = ini.GetValue("defaultBashPath", "AndroidCapabilities");
            System.Diagnostics.Process.Start(defaultBashPath);
            Thread.Sleep(5000);
        }

        public void CapabilitiesDevice()
        {
            _capabilities = new DesiredCapabilities();
            _capabilities.SetCapability("device", deviceType.ToString());
            _capabilities.SetCapability(CapabilityType.Platform, "Windows");
            _capabilities.SetCapability("deviceName", deviceName.ToString());
            _capabilities.SetCapability("platformName", appPlatform.ToString());
            _capabilities.SetCapability("platformVersion", platformVersion.ToString());
            _capabilities.SetCapability("appPackage", appPackage.ToString());
            _capabilities.SetCapability("appActivity", appActivity.ToString());
            //_capabilities.SetCapability("newCommandTimeout", 0);
            
            driver = new AndroidDriver<AndroidElement>(new Uri(serverUri1), _capabilities, TimeSpan.FromSeconds(180));
            //Webdriver = new AndroidDriver<AndroidElement>(new Uri(serverUri1), _capabilities, TimeSpan.FromSeconds(180));
            tempDriver1 = driver;
        }
    }
}
