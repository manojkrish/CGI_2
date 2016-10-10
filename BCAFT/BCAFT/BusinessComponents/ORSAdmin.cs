using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using TDAPIOLELib;

namespace BCAFT.BusinessComponents
{
    class ORSAdmin : ConfigClass
    {

        CommonFunctions objComm;
       CommonLibraries objCommLib;
       public static string ProjectPath;

       public ORSAdmin(HelperClass helperClass)
           : base(helperClass)
       {
            objComm = new CommonFunctions(helperClass);
            objCommLib = new CommonLibraries(helperClass);
        }
        //Property file initialization
        public Properties ObjAdminProp = new Properties(GetPropertiesFilePath() + "\\Admin.txt");
        
        //Initialize Test data sheet
        public void Initialize(string strSheetName)
        {
            if (!arrListSheetsqueried.Contains(strSheetName + strScenario + strTestCase + strStartIteration))
            {
                objCommLib.readTestData(null, strSheetName, strStartIteration);
                arrListSheetsqueried.Add(strSheetName + strScenario + strTestCase + strStartIteration);
            }          
        }

        public void launchORSAdmin()
        {
            Properties obj = new Properties(strResultsPath + "\\ConfigFile.txt");
            driver.Navigate().GoToUrl(obj.get("Admin_URL"));
            objComm.wait(By.XPath("//div[@id='dialogNewVersionDetail_dialog']"), 20); ////select[contains(@id,'ddlCountry')]
            objCommLib.logInfo("Launch ORS Admin portal", "ORS Amin portal launched successfully", "Pass");
            driver.Manage().Window.Maximize();
            objComm.click(By.XPath("//div[@id='dialogNewVersionDetail_dialog']//input[@id='dialogNewVersionDetail_bOK']"),"OK - NewVersion details");
        }

        public void navigateToLevel()
        {
            Initialize("InputData");
            string strCountry = objCommLib.readTestData("Country");
            string strCentre = objCommLib.readTestData("Centre");
            string strVenue = objCommLib.readTestData("Venue");
            //Expand Global
            if (expand("Global"))
            {
                if (expand("British Council"))
                {
                    if (!(strCountry.Equals("")) && !(strCountry.Equals("NULL", StringComparison.InvariantCultureIgnoreCase)))
                    {
                        if (expand(strCountry))
                        {
                            if (!(strCentre.Equals("")) && !(strCentre.Equals("NULL", StringComparison.InvariantCultureIgnoreCase)))
                            {
                                if (expand(strCentre))
                                {
                                    if (!(strVenue.Equals("")) && !(strVenue.Equals("NULL", StringComparison.InvariantCultureIgnoreCase)))
                                    {
                                        objComm.click(By.XPath("//div[@id='Menu1_TreeViewDisplay1_dTree']//a[text()='" + strVenue + "']"), "Venue " + strVenue);
                                    } else
                                        objComm.click(By.XPath("//div[@id='Menu1_TreeViewDisplay1_dTree']//a[text()='" + strCentre + "']"), "Centre " + strCentre);
                                }
                            }else
                                objComm.click(By.XPath("//div[@id='Menu1_TreeViewDisplay1_dTree']//a[text()='" + strCountry + "']"), "Country " + strCountry);
                        }
                    }else
                        objComm.click(By.XPath("//div[@id='Menu1_TreeViewDisplay1_dTree']//a[text()='British Council']"), "Organisation British Council");
                }
            }
            
        }

        private Boolean expand(string strName)
        {
            IWebElement ele1 = objComm.find(By.XPath("//div[@id='Menu1_TreeViewDisplay1_dTree']//a[text()='" + strName + "']"), strName + " link");
            IWebElement parentEle1 = objComm.GetParent(ele1);
            if (parentEle1 != null)
            {
                if (((string)parentEle1.GetAttribute("class")).StartsWith("dynatree-node dynatree-has-children"))
                {                    
                    objComm.click(parentEle1, By.CssSelector(".dynatree-expander"), strName + " Expander");
                    objCommLib.logInfo("Expand level", strName + " - Selected", "Pass");
                    return true;
                }//Venue - Parent Span
                else
                    return true;
            }
            else
            {
                objCommLib.logInfo("Expand level", strName + " - Expander not found", "Fail");
                return false;
            }
        }

        public void chooseTab()
        {
            Initialize("InputData");
            string strTab = objCommLib.readTestData("TabName");
            Boolean flag = false;
            IWebElement parentEle = objComm.find(By.XPath("//ul[@id='uTabs']"),"Parent Tab");
            IList<IWebElement> wElement = objComm.findElements(parentEle, By.XPath("//a[text()='" + strTab + "']"), strTab + " tab");
            for (int i = 0; i < wElement.Count; i++)
            {
                if ((wElement.ElementAt(i)).Enabled && (wElement.ElementAt(i)).Displayed)
                {
                    Console.WriteLine((string)(wElement.ElementAt(i).GetCssValue("style")));
                    wElement.ElementAt(i).Click();
                    flag = true;
                    break;
                }
            }
            if(!flag)
                objCommLib.logInfo("Select Tab", strTab + " Tab not found", "Fail");
            else
                objCommLib.logInfo("Select Tab", strTab + " Tab selected", "Pass");
            
        }

        public void chooseShowEntries()
        {

        }

        public void QCInt()
        {
            //String qcUrl = "http://10.35.6.16:8080/qcbin";
            //String qcDomain = "LOB";
            //String qcProject = "MyPerformancePortfolio";
            //String qcLoginName = "vithyasankaran";
            //String qcPassword = "";
            //String testSetPath = @"ROOT\Unattached";
            //String testSetName = "default"; // "[1]Test the funcvtionality of login screen";
            //QCIntegration obj = new QCIntegration();
            //obj.Connect(qcUrl, qcDomain, qcProject, qcLoginName, qcPassword);
            //TestSet testSet = obj.GetTestSet(testSetPath, testSetName);
            //obj.RunTestSet(testSet);

            //QCI obj1 = new QCI();
            //obj1.sendRequest();

            QCInt obj1 = new QCInt();
            obj1.qcResultUpdate();
        }

    }
}
