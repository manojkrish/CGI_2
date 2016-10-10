using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Globalization;

namespace BCAFT
{
    class CommonFunctions_Old : ConfigClass
    {
        CommonLibraries objCommLib;
        public CommonFunctions_Old(HelperClass helperClass)
            : base(helperClass)
        {
            objCommLib = new CommonLibraries(helperClass);
        }

        public Boolean verifyPageText(IWebDriver driver, string strExpText)
        {
            String bodyTxt = driver.FindElement(By.TagName("body")).Text;
            if (bodyTxt.Contains(strExpText))
                return true;
            else
                return false;
        }

        public IWebElement find(By by, string strElementNameInUI)
        {
            IWebElement wElement = null; int timeOut = 0;
            do
            {
                if (timeOut > 0)
                {
                    sleep(3);
                }
                try
                {
                    wElement = driver.FindElement(by);
                }
                catch (NoSuchElementException e) { timeOut++; }
            } while (wElement == null && timeOut < 5);
            if (wElement == null)
            {
                objCommLib.logInfo("Find Element", strElementNameInUI + " - could not found", "Fail");
            }
            return wElement;
        }

        public IWebElement findWebElement(IWebElement parentEle, By by, string strElementNameInUI)
        {
            IWebElement wElement = null; int timeOut = 0;
            do
            {
                if (timeOut > 0)
                {
                    sleep(3);
                }
                try
                {
                    wElement = parentEle.FindElement(by);
                }
                catch (NoSuchElementException e) { timeOut++; }
            } while (wElement == null && timeOut < 5);
            if (wElement == null)
            {
                objCommLib.logInfo("Find Element", strElementNameInUI + " - could not found", "Fail");
            }
            return wElement;
        }

        public Boolean ifExist(By by, string strElementNameInUI)
        {
            IWebElement wElement = null; int timeOut = 0;
            do
            {
                if (timeOut > 0)
                {
                    sleep(3);
                }
                try
                {
                    wElement = driver.FindElement(by);
                }
                catch (NoSuchElementException e) { timeOut++; }
            } while (wElement == null && timeOut < 2);
            if (wElement == null)
            {
                return false;
            }else
                return true;
        }

        public void select(By by, String elementName, string strValue)
        {
            IWebElement wElement = null; int timeOut = 0;
            if (strValue != null && !strValue.Equals(""))
            {
                do
                {
                    if (timeOut > 0)
                    {
                        sleep(2);
                    }
                    try
                    {
                        wElement = driver.FindElement(by);
                        if (wElement != null)
                        {
                            try
                            {
                                SelectElement select = new SelectElement(wElement);
                                select.SelectByText(strValue);
                            }
                            catch (Exception e)
                            {
                                objCommLib.logInfo("Select Option", "Exception occured while trying to select " + elementName + " [" + strValue + "]. Exception is [" + e + "]", "Fail");
                            }
                        }//Ele Not null
                    }
                    catch (NoSuchElementException e) { timeOut++; }
                } while (wElement == null && timeOut < 5);
                if (wElement == null)
                {
                    objCommLib.logInfo("Find Element", elementName + " - could not found", "Fail");
                }
            }
        }

        public void select(By by, String elementName)
        {
            string strValue = objCommLib.readTestData(elementName);
            IWebElement wElement = null; int timeOut = 0;
            if (strValue != null && !strValue.Equals(""))
            {
                do
                {
                    if (timeOut > 0)
                    {
                        sleep(2);
                    }
                    try
                    {
                        wElement = driver.FindElement(by);
                        if (wElement != null)
                        {
                            try
                            {
                                SelectElement select = new SelectElement(wElement);
                                select.SelectByText(strValue);
                                wElement = driver.FindElement(by);
                                select = new SelectElement(wElement);
                                objCommLib.logInfo("Choose " + elementName, "Selected " + select.SelectedOption.Text, "Done");
                            }
                            catch (Exception e)
                            {
                                objCommLib.logInfo("Select Option", "Exception occured while trying to select " + elementName + " [" + strValue + "]. Exception is [" + e + "]", "Fail");
                            }
                        }//Ele Not null
                    }
                    catch (NoSuchElementException e) { timeOut++; }
                } while (wElement == null && timeOut < 5);
                if (wElement == null)
                {
                    objCommLib.logInfo("Find Element", elementName + " - could not found", "Fail");
                }
            }
        }

        public void select(By by, String elementName, int iIndex)
        {
            IWebElement wElement = null; int timeOut = 0;
            if (!iIndex.Equals(""))
            {
                do
                {
                    if (timeOut > 0)
                    {
                        sleep(2);
                    }
                    try
                    {
                        wElement = driver.FindElement(by);
                        if (wElement != null)
                        {
                            try
                            {
                                SelectElement select = new SelectElement(wElement);
                                select.SelectByIndex(iIndex);
                                objCommLib.logInfo("Choose " + elementName, "Selected index " + iIndex, "Done");
                            }
                            catch (Exception e)
                            {
                                objCommLib.logInfo("Select Oprion by Index", "Exception occured while trying to select " + elementName + " by index [" + iIndex + "]. Exception is [" + e + "]", "Fail");
                            }
                        }//Ele Not null
                    }
                    catch (NoSuchElementException e) { timeOut++; }
                } while (wElement == null && timeOut < 5);
                if (wElement == null)
                {
                    objCommLib.logInfo("Find Element", elementName + " - could not found", "Fail");
                }
            }
        }

        //public void findAndEnter(IWebDriver driver, By by, string strValue, String elementName)
        //{
        //    IWebElement wElement = null; int timeOut = 0;
        //    if (strValue != null && !strValue.Equals(""))
        //    {
        //        do
        //        {
        //            if (timeOut > 0)
        //            {
        //                sleep( 2);
        //            }
        //            try
        //            {
        //                wElement = driver.FindElement(by);
        //                if (wElement != null)
        //                {
        //                    try
        //                    {
        //                        wElement.Clear();
        //                        wElement.SendKeys(strValue);
        //                    }
        //                    catch (Exception e)
        //                    {
        //                        Console.WriteLine("Exception occured while trying to enter " + elementName + " [" + strValue + "]. Exception is [" + e + "]");
        //                    }
        //                }//Ele Not null
        //            }
        //            catch (NoSuchElementException e) { timeOut++; }
        //        } while (wElement == null && timeOut < 5);
        //        if (wElement == null)
        //        {
        //            Console.WriteLine(elementName + " - could not found");
        //        }
        //    }
        //}

        public void enter(By by, String elementName, string strValue)
        {
            IWebElement wElement = null; int timeOut = 0;
            if (strValue != null && !strValue.Equals(""))
            {
                do
                {
                    if (timeOut > 0)
                    {
                        sleep(2);
                    }
                    try
                    {
                        wElement = driver.FindElement(by);
                        if (wElement != null)
                        {
                            try
                            {
                                wElement.Clear();
                                wElement.SendKeys(strValue);
                            }
                            catch (Exception e)
                            {
                                objCommLib.logInfo("Enter Value", "Exception occured while trying to enter " + elementName + " [" + strValue + "]. Exception is [" + e + "]", "Fail");
                            }
                        }//Ele Not null
                    }
                    catch (NoSuchElementException e) { timeOut++; }
                } while (wElement == null && timeOut < 5);
                if (wElement == null)
                {
                    objCommLib.logInfo("Find Element", elementName + " - could not found", "Fail");
                }
            }
        }

        public void enter(By by, String elementName)
        {
            string strValue = objCommLib.readTestData(elementName);
            if (strValue.Equals("Random", StringComparison.InvariantCultureIgnoreCase))
            {
                strValue = generateRandomChars(4);
            }
            IWebElement wElement = null; int timeOut = 0;
            if (strValue != null && !strValue.Equals(""))
            {
                do
                {
                    if (timeOut > 0)
                    {
                        sleep(2);
                    }
                    try
                    {
                        wElement = driver.FindElement(by);
                        if (wElement != null)
                        {
                            try
                            {
                                wElement.Clear();
                                wElement.SendKeys(strValue);
                                objCommLib.logInfo("Enter " + elementName, "Entered " + strValue, "Done");
                            }
                            catch (Exception e)
                            {
                                objCommLib.logInfo("Enter Value", "Exception occured while trying to enter " + elementName + " [" + strValue + "]. Exception is [" + e + "]", "Fail");
                            }
                        }//Ele Not null
                    }
                    catch (NoSuchElementException e) { timeOut++; }
                } while (wElement == null && timeOut < 5);
                if (wElement == null)
                {
                    objCommLib.logInfo("Find Element", elementName + " - could not found", "Fail");
                }
            }
        }

        public void click(String txt)
        {
            IWebElement wElement = null; int timeOut = 0;
            do
            {
                if (timeOut > 0)
                {
                    sleep(6);
                }
                try
                {
                    wElement = driver.FindElement(By.PartialLinkText(txt));
                    if (wElement != null)
                        wElement.Click();
                }
                catch (NoSuchElementException e) { timeOut++; }
            } while (wElement == null && timeOut < 5);
            if (wElement == null)
            {
                txt = txt.Replace("'", "");
                objCommLib.logInfo("Find Link", txt + " - not found", "Fail");
            }
        }

        public IWebElement click(By by, string elementName)
        {
            IWebElement wElement = null; int timeOut = 0;
            do
            {
                if (timeOut > 0)
                {
                    sleep(5);
                }
                try
                {
                    wElement = driver.FindElement(by);
                    if (wElement != null)
                    {
                        waitTillPageload_UKVI(driver, 10);
                        //((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", wElement);
                        wElement.Click();
                    }
                }
                catch (NoSuchElementException e)
                {
                    try
                    {
                        if (e.StackTrace.ToString().ToLower().Replace(" ", "").Contains("elementnotvisible"))
                        {
                            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", wElement);
                            wElement.Click();
                        }
                    }
                    catch (Exception e1)
                    {
                        // TODO Auto-generated catch block
                        objCommLib.logInfo("Click on Element", e1.Message, "Fail");
                    }
                    sleep(5);
                    timeOut++;
                }
            } while (wElement == null && timeOut < 4);

            if (wElement == null)
            {
                objCommLib.logInfo("Find Element", elementName + " not found", "Fail");
            }
            return wElement;
        }

        public void waitTillPageload_UKVI(IWebDriver driver, int timeOut)
        {
            do
            {
                try
                {
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@id='overlay-background']")));
                    timeOut--;
                }
                catch (Exception e)
                {
                    sleep(3);
                    timeOut = 0;
                }
            } while (timeOut > 0);
        }

        public Boolean click(IWebElement parentEle, By by, string elementName)
        {
            IWebElement wElement = null; int timeOut = 0;
            do
            {
                if (timeOut > 0)
                {
                    sleep(2);
                }
                try
                {
                    wElement = parentEle.FindElement(by);
                    if (wElement != null)
                    {
                        try
                        {
                            wElement.Click();
                            sleep(1);
                        }
                        catch (Exception e)
                        {
                            objCommLib.logInfo("Find Element", "Exception occured while trying to click on " + elementName + ". Exception is [" + e + "]", "Fail");
                        }
                    }
                }
                catch (NoSuchElementException e) { timeOut++; }
            } while (wElement == null && timeOut < 5);
            if (wElement == null)
            {
                objCommLib.logInfo("Find Element", "Element -" + elementName + "- not found", "Fail");
                return false;
            }
            else
                return true;
        }

        public void wait(By by, int timeoutInSeconds)
        {
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(drv => drv.FindElement(by));
            }
            catch (OpenQA.Selenium.NoSuchElementException e)
            {
                objCommLib.logInfo("Find element", "Unable to find element -" + by + ". Exception:" + e.Message, "Fail");
                stopExecution = true;
            }
            catch (Exception e)
            {
                objCommLib.logInfo("Find element", "Exception while finding element -" + by + "." + e.Message, "Fail");
                stopExecution = true;
            }
        }

        public void waitForElementToBeClickable(IWebDriver driver, IWebElement ele, int timeoutInSeconds, String strEleName)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.ElementToBeClickable(ele));
            }
            catch (Exception e)
            {
                // TODO Auto-generated catch block
                Console.WriteLine("Element " + strEleName + ", not clickable");
            }
        }

        public void sleep(int seconds)
        {
            try
            {
                //		 int mSec = seconds * 100;
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(seconds));
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public IList<IWebElement> findElements(By by, String elementName)
        {
            IList<IWebElement> wElement = null; int timeOut = 0;
            do
            {
                if (timeOut > 0)
                {
                    sleep(5);
                }
                try
                {
                    wElement = driver.FindElements(by);
                }
                catch (NoSuchElementException e)
                {
                    sleep(5);
                }
                timeOut++;
            } while (wElement.Equals(null) && timeOut < 4);

            if (wElement == null)
            {
                Console.WriteLine(elementName + " not found");
            }
            return wElement;
        }

        public IList<IWebElement> findElements(IWebElement parentEle, By by, String elementName)
        {
            IList<IWebElement> wElement = null; int timeOut = 0;
            do
            {
                if (timeOut > 0)
                {
                    sleep(5);
                }
                try
                {
                    wElement = parentEle.FindElements(by);
                }
                catch (NoSuchElementException e)
                {
                    sleep(5);
                }
                timeOut++;
            } while (wElement.Equals(null) && timeOut < 4);

            if (wElement == null)
            {
                Console.WriteLine(elementName + " not found");
            }
            return wElement;
        }

        public string generatePIN()
        {
            int _min = 1000;
            int _max = 9999;
            Random _rdm = new Random();
            return _rdm.Next(_min, _max).ToString();
        }

        private static Random random = new Random();
        public static string generateRandomChars(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        public string formatToDate(string strDate)
        {
            DateTime dt = DateTime.ParseExact(strDate.ToString(), "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            return dt.ToString("dd/M/yyyy", CultureInfo.InvariantCulture);
        }
    }
}
