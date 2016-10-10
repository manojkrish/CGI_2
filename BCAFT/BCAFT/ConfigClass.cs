using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
using System.Collections;
using System.IO;

namespace BCAFT
{    
 
    public class ConfigClass
    {        
        protected IWebDriver driver;
        protected HelperClass helperClass;
        protected string strScenario, strTestCase, strTestCaseDescription, strBrowserName, strResultsPath, strStartIteration, strEndIteration;
        public static string strResult="Fail";
        public static ArrayList arrListSheetsqueried = new ArrayList();
        public static Dictionary<string, object> dic_InputData = new Dictionary<string, object>();
        public static Boolean stopExecution = false;
         
        public ConfigClass(HelperClass helperClass)
        {
            this.helperClass = helperClass;
            this.driver = helperClass.getDriver();
            this.strScenario = helperClass.getScenario();
            this.strTestCase = helperClass.getTestCase();
            this.strTestCaseDescription = helperClass.getTestCaseDescription();
            this.strBrowserName = helperClass.getBrowserName();
            this.strResultsPath = helperClass.getResultsPath();
            this.strStartIteration = helperClass.getStartIteration();
            this.strEndIteration = helperClass.getEndIteration();
        }

        protected string GetProjectPath()
        {
            string temp = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            return temp;
        }

        protected string  GetTestDataPath() {
            //string temp = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            return GetProjectPath() + "\\Data\\" + strScenario + ".xlsx";
        }

        public static string GetPropertiesFilePath()
        {
            string temp = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName + "\\ObjectProperties";
            return temp;
        }

        //-------------------Object Property File------------------------------------//
        //protected Object initializeCustomerJourneyPropertyFile()
        //{
        //    string strResultsPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName + "\\ObjectProperties";
        //    Object CJProperties = new Properties(strResultsPath + "\\CustomerJourney.txt");
        //    return CJProperties;
        //}

        //protected Object initializeAdminPropertyFile()
        //{
        //    string strResultsPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName + "\\ObjectProperties";
        //    Object CJProperties = new Properties(strResultsPath + "\\Admin.txt");
        //    return CJProperties;
        //}
    }
}
