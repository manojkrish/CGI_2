using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Reflection;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using System.Data.OleDb;

namespace BCAFT
{
    public class MainClass
    {
        private HelperClass helperClass;
        private IWebDriver driver;
        private string strCtScenario, strCtTestCase, strCtTestCaseDescription, strBrowserName, strStartIteration, strEndIteration;
        private string strResultsPath;

        public static void run()
        {
            MainClass objMainClass = new MainClass();
            DataTable dtRunInfo = objMainClass.getRunInfo();
            objMainClass.strResultsPath = objMainClass.initializeResultLog();
            objMainClass.executeTestCase(dtRunInfo);            
        }

        private void initializeWebdriver(string strBrowserName)
        {
            string temp = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName + "\\Libraries";
            switch (strBrowserName)
            {
                case "Chrome":
                    driver = new ChromeDriver(@""+ temp);
                    break;
                case "IE":
                    driver = new InternetExplorerDriver(@"" + temp);
                    break;
                case "Firefox":
                    driver = new FirefoxDriver();
                    break;
            }
        }

        private string initializeResultLog()
        {
            string strResultsPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName + "\\Results";
            strResultsPath = strResultsPath +"\\"+ DateTime.Now.ToString().Replace('/', '-').Replace(':', '.');
            System.IO.Directory.CreateDirectory(strResultsPath);
            //Screenshots folder
            string ssPath = strResultsPath + "\\Screenshots";
            System.IO.Directory.CreateDirectory(ssPath);
            return strResultsPath;
        }        

        private DataTable getRunInfo()
        {            
            DataTable dtRunInfo = new DataTable();
            String strQuery = "SELECT * FROM [RunSettings$] WHERE Execute='Yes'";
            string temp = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            dtRunInfo = readDataFromExcel(temp + "\\RunSettings.xlsx", strQuery);           
            return dtRunInfo;
        }

        private void executeTestCase(DataTable dtRunInfo)
        {            
            for (int i = 0; i < dtRunInfo.Rows.Count; i++)
            {
                strCtScenario = (string)dtRunInfo.Rows[i]["ModuleName"];
                strCtTestCase = (string)dtRunInfo.Rows[i]["TestCaseNo"];
                strCtTestCaseDescription = (string)dtRunInfo.Rows[i]["Description"];
                strBrowserName = (string)dtRunInfo.Rows[i]["Browser"];
                strStartIteration = Convert.ToString(dtRunInfo.Rows[i]["Start Iteration"]);
                strEndIteration = Convert.ToString(dtRunInfo.Rows[i]["End Iteration"]);
                string strQCResultUpdateFlag = Convert.ToString(dtRunInfo.Rows[i]["QCResultUpdate"]);
                initializeWebdriver(strBrowserName);
                helperClass = new HelperClass(strCtScenario, strCtTestCase, strCtTestCaseDescription, strBrowserName, driver, strResultsPath, strStartIteration, strEndIteration);
                helperClass.setQCUpdateResultFlag(strQCResultUpdateFlag);
                CommonLibraries commLib = new CommonLibraries(helperClass);                
                if (i == 0)
                {
                    commLib.initializeSummaryResult();
                    commLib.initializeEnvironmentSetting();
                    if (strQCResultUpdateFlag.Equals("Yes", StringComparison.InvariantCultureIgnoreCase))
                        commLib.loadQCVariables();
                }
                commLib.createTCResultLog();
                executeIterations(strCtScenario, strCtTestCase);
                commLib.logSummaryResult();
                commLib.consolidateAllScreenshots();
                if (strQCResultUpdateFlag.Equals("Yes", StringComparison.InvariantCultureIgnoreCase))
                    commLib.qcResultUpdate();
                if (i == dtRunInfo.Rows.Count - 1)
                    commLib.initializeSummaryResultFooter();
                tearDown();
                
            }//FOR each Row in Data Table
        }

        private void executeIterations(string strModule, string strTC)
        {
            Boolean flag = false; ;
            CommonLibraries commLib = new CommonLibraries(helperClass);
            
            //Get All Class files
            Type type = this.GetType();
            Type[] typelist = GetTypesInNamespace(Assembly.GetExecutingAssembly(), type.Namespace + ".BusinessComponents");

            //Get Keywords from Data table
            String strQuery = "SELECT * FROM [Keyword$] WHERE TestCaseNo='" + strTC.Trim() + "'";
            string temp = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            DataTable dtKeyword = commLib.readDataFromExcel(temp + "\\Data\\" + strModule + ".xlsx", strQuery);
            //Iteration Execution
            for (int ctIteration = Convert.ToInt16(strStartIteration); ctIteration <= Convert.ToInt16(strEndIteration); ctIteration++)
            {
                helperClass.setStartIteration(Convert.ToString(ctIteration));
                commLib.logInfo("Iteration:" + Convert.ToString(ctIteration));//Log Iteration No
                flag = executeKeywords(commLib, typelist, dtKeyword);
                if (flag)
                {
                    break;
                }
            }//Each Iteration
            commLib.initializeTCResultFooter();
        }

        private Boolean executeKeywords(CommonLibraries commLib, Type[] typelist, DataTable dtKeyword)
        {
            Boolean isMethodFound = false, flag = false;
            for (int r = 1; r < dtKeyword.Columns.Count; r++) //Read Method Names
            {
                if ((!object.ReferenceEquals(dtKeyword.Rows[0][r], DBNull.Value)))
                {
                    isMethodFound = false;
                    for (int i = 0; i < typelist.Length; i++)//Search in all Classes
                    {
                        MethodInfo ctMethod = typelist[i].GetMethod((string)dtKeyword.Rows[0][r]);
                        if (ctMethod != null)
                        {
                            commLib.logInfo((string)dtKeyword.Rows[0][r]);//Log Keyword name
                            //Call Method
                            isMethodFound = true;
                            if (ctMethod.IsStatic)
                                ctMethod.Invoke(null, null);
                            else
                            {
                                try
                                {
                                    object instance = Activator.CreateInstance(typelist[i], helperClass);
                                    ctMethod.Invoke(instance, null);
                                }
                                catch (Exception e)
                                {
                                    ConfigClass.stopExecution = true;
                                    commLib.logInfo("Execute Keyword", "Keyword " + (string)dtKeyword.Rows[0][r] + ". Exception " + e.Message, "Exception");
                                    //throw;
                                }
                                if (ConfigClass.stopExecution)
                                {
                                    flag = true; break;
                                }
                            }
                            break;
                        }
                    }//Each Class
                    if (!isMethodFound)
                    {
                        commLib.logInfo("Execute Keyword", "Keyword " + (string)dtKeyword.Rows[0][r] + " not found", "Exception");
                        flag = true; break;
                    }
                }// IF Keyword not null                
            }//Each Method
            return flag;
        }

        private void ABC(string strModule, string strTC)
        {
            Boolean isMethodFound = false, flag = false; ;
            CommonLibraries commLib = new CommonLibraries(helperClass);
            commLib.createTCResultLog();     
            //Get All Class files
            Type type = this.GetType();
            Type[] typelist = GetTypesInNamespace(Assembly.GetExecutingAssembly(), type.Namespace + ".BusinessComponents");

            //Get Keywords from Data table
            String strQuery = "SELECT * FROM [Keyword$] WHERE TestCaseNo='"+ strTC.Trim() +"'";
            string temp = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            DataTable dtKeyword = commLib.readDataFromExcel(temp + "\\Data\\" + strModule + ".xlsx", strQuery);
            //Iteration Execution
            for (int ctIteration = Convert.ToInt16(strStartIteration); ctIteration <= Convert.ToInt16(strEndIteration); ctIteration++)
            {
                helperClass.setStartIteration(Convert.ToString(ctIteration));
                commLib.logInfo(Convert.ToString(ctIteration));//Log Iteration No
                for (int r = 1; r < dtKeyword.Columns.Count; r++) //Read Method Names
                {                    
                    if ((!object.ReferenceEquals(dtKeyword.Rows[0][r], DBNull.Value)))
                    {
                        isMethodFound = false;
                        for (int i = 0; i < typelist.Length; i++)//Search in all Classes
                        {
                            MethodInfo ctMethod = typelist[i].GetMethod((string)dtKeyword.Rows[0][r]);
                            if (ctMethod != null)
                            {
                                commLib.logInfo((string)dtKeyword.Rows[0][r]);//Log Keyword name
                                //Call Method
                                isMethodFound = true;
                                if (ctMethod.IsStatic)
                                    ctMethod.Invoke(null, null);
                                else
                                {
                                    try
                                    {
                                        object instance = Activator.CreateInstance(typelist[i], helperClass);
                                        ctMethod.Invoke(instance, null);
                                    }
                                    catch (Exception e)
                                    {
                                        ConfigClass.stopExecution = true;
                                        commLib.logInfo("Execute Keyword", "Keyword " + (string)dtKeyword.Rows[0][r] + ". Exception "+e.Message, "Fail");
                                        //throw;
                                    }
                                    if (ConfigClass.stopExecution)
                                    {
                                        flag = true; break;
                                    }
                                }
                                break;
                            }
                        }//Each Class
                        if (!isMethodFound)
                        {
                            commLib.logInfo("Execute Keyword", "Keyword " + (string)dtKeyword.Rows[0][r] + " not found", "Fail");
                            flag = true; break;
                        }
                    }// IF Keyword not null                
                }//Each Method
                if (flag)
                {                    
                    break;
                }
            }//Each Iteration
            commLib.initializeTCResultFooter();      
        }

        public void tearDown()
        {
            driver.Quit();
        }

        //-------------------------------------------------------------------------------//

        public DataTable readDataFromExcel(string strFilePath, string strQuery)
        {
            OleDbCommand cmd = new OleDbCommand();
            OleDbDataAdapter da = new OleDbDataAdapter();
            //DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            String query = strQuery; // "SELECT * FROM [" + strSheetName + "$]";
            OleDbConnection conn = new OleDbConnection(connString);
            if (conn.State == ConnectionState.Closed) conn.Open();
            try
            {
                cmd = new OleDbCommand(query, conn);
                da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
            }
            catch
            {
                // Exception Msg 

            }
            finally
            {
                da.Dispose();
                conn.Close();
            }
            return dt;
        }
        private Type[] GetTypesInNamespace(Assembly assembly, string nameSpace)
        {
            return assembly.GetTypes().Where(t => String.Equals(t.Namespace, nameSpace, StringComparison.Ordinal)).ToArray();
        }

        

        

    }
}
