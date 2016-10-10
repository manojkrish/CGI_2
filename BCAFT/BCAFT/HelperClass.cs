using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
using System.IO;

namespace BCAFT
{
  public class HelperClass
    {
        private string strScenario, strTestCase, strTestCaseDescription, strBrowserName, strResultsPath, strStartIteration, strEndIteration;
        private string strQCUpdateResultFlag;
        private IWebDriver driver; //{ get; set; }        


        public HelperClass(string strScenario, string strTestCase, string strTestCaseDescription, string strBrowserName, IWebDriver driver, string strResultsPath, string strStartIteration, string strEndIteration)
        {
            this.strScenario = strScenario;
            this.strTestCase = strTestCase;
            this.strTestCaseDescription = strTestCaseDescription;
            this.strBrowserName = strBrowserName;
            this.driver = driver;
            this.strResultsPath = strResultsPath;
            this.strStartIteration = strStartIteration;
            this.strEndIteration = strEndIteration;
        }

        //public HelperClass(string strResult)
        //{
        //    this.strResult = strResult;
        //}

        public HelperClass(IWebDriver driver)
        {
            this.driver = driver;
        }

        public IWebDriver getDriver()
        {
            return driver;
        }

        public void setDriver(IWebDriver driver)
        {
            this.driver = driver;
        }

        public string getScenario()
        {
            return strScenario;
        }

        public void setScenario(string strScenario)
        {
            this.strScenario = strScenario;
        }

        public string getTestCase()
        {
            return strTestCase;
        }

        public void setTestCase(string strTestCase)
        {
            this.strTestCase = strTestCase;
        }

        public string getTestCaseDescription()
        {
            return strTestCaseDescription;
        }

        public void setTestCaseDescription(string strTestCaseDescription)
        {
            this.strTestCaseDescription = strTestCaseDescription;
        }               

        public string getBrowserName()
        {
            return strBrowserName;
        }

        public void setBrowserName(string strBrowserName)
        {
            this.strBrowserName = strBrowserName;
        }

        public string getResultsPath()
        {
            return strResultsPath;
        }

        public void setResultsPath(string strResultsPath)
        {
            this.strResultsPath = strResultsPath;
        }

        public string getStartIteration()
        {
            return strStartIteration;
        }

        public void setStartIteration(string strStartIteration)
        {
            this.strStartIteration = strStartIteration;
        }

        public string getEndIteration()
        {
            return strEndIteration;
        }

        public void setEndIteration(string strEndIteration)
        {
            this.strEndIteration = strEndIteration;
        }

        public string getQCUpdateResultFlag()
        {
            return strQCUpdateResultFlag;
        }

        public void setQCUpdateResultFlag(string strQCUpdateResultFlag)
        {
            this.strQCUpdateResultFlag = strQCUpdateResultFlag;
        }

        //public string getResult()
        //{
        //    return this.strResult;
        //}

        //public void setResult(string strResult)
        //{
        //    this.strResult = strResult;
        //}

    }
}
