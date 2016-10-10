using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using OpenQA.Selenium;
using System.Drawing.Imaging;
using System.IO;
using TDAPIOLELib; 

namespace BCAFT
{
    public class CommonLibraries : ConfigClass
    { 
        public CommonLibraries(HelperClass helperClass) : base(helperClass) { }
        public static int iStepNo, iPassedStepCount, iFailedStepCount = 0, iPassedTCCount, iFailedTCCount = 0;
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

        public string readTestData(string strColumnName)
        {
            return Convert.ToString(dic_InputData[strColumnName]);       

        }

        public void readTestData(string[] strColumns, string strSheetName, string strIteration)
        {
            string ctKey = null; object ctValue = null;
            string strQuery = createQueryforTestData(strColumns, strSheetName, strTestCase, strIteration);
            DataTable dtResult = readDataFromExcel(GetTestDataPath(), strQuery);
            try
            {
                if (dtResult.Rows.Count > 0)
                {
                    for (int d = 1; d < dtResult.Columns.Count; d++)
                    {
                        ctKey = (string)dtResult.Columns[d].ColumnName;
                        ctValue = dtResult.Rows[0][d];
                        if (dic_InputData.ContainsKey(ctKey))
                        {
                            dic_InputData[ctKey] = ctValue;
                        }
                        else
                        {
                            dic_InputData.Add(ctKey, ctValue);
                        }
                    }
                }
                else
                {
                    logInfo("Read Datasheet", "No records found for the Test Case " + strTestCase + ", Iteration " + strIteration, "Fail");
                    dic_InputData.Clear();
                    stopExecution = true;
                }
            }catch(Exception e){
                //Error while reading data from Data sheet
                logInfo("Read Test data", "Exception :" + e.Message, "Fail");
            }
        }

        public string createQueryforTestData(string[] strArrColumns, string strSheetname, string strTC, string strIteration)
        {
            string strQuery, strColumns = null; int c = 0;
            if (strColumns == null)
            {
                strColumns = "*";
            }
            else
            {
                while (c < strArrColumns.Length)
                {
                    if (c == 0)
                        strColumns = strArrColumns[0];
                    else
                    {
                        strColumns = strColumns + "," + strArrColumns[c];
                    }
                    c++;
                }
            }
            strQuery = "select " + strColumns + " from [" + strSheetname + "$] where TestCaseNo='" + strTC + "' and Iteration=" + strIteration + "";
            return strQuery;
        }

        public void initializeEnvironmentSetting()
        {
            DataTable dtRunInfo = new DataTable();
            string strQuery = "SELECT * FROM [EnvironmentSettings$] WHERE Execute='Yes'";
            string temp = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            dtRunInfo = readDataFromExcel(temp + "\\RunSettings.xlsx", strQuery);
            if (dtRunInfo.Rows.Count > 0)
            {
                string line, key, value;
                int index;
                createFile(strResultsPath + "\\ConfigFile.txt");
                System.IO.StreamReader file = new System.IO.StreamReader(temp + "\\EnvironmentDetails.txt");
                try
                {
                    while ((line = file.ReadLine()) != null)
                    {
                        if ((!string.IsNullOrEmpty(line)) && (!line.StartsWith("*")))
                        {
                            index = line.IndexOf('=');
                            key = line.Substring(0, index).Trim();
                            if (key.Equals("ENVIRONMENT", StringComparison.InvariantCultureIgnoreCase))
                            {
                                value = line.Substring(index + 1).Trim();
                                if (value.Equals((string)dtRunInfo.Rows[0]["Environment"], StringComparison.InvariantCultureIgnoreCase))
                                {                                    
                                    //string[] configLines = {"Environment="+value};
                                    IList<string> configLines = new List<string>();
                                    line = file.ReadLine();
                                    line = file.ReadLine();
                                    while (line.Trim().IndexOf("*") == -1)
                                    {
                                        index = line.IndexOf('=');
                                        key = line.Substring(0, index).Trim();
                                        value = line.Substring(index + 1).Trim();
                                        //dic_EnvironmentData.Add(key, value);
                                        configLines.Add(line);
                                        line = file.ReadLine();
                                    }//While Loop  
                                    file.Close();
                                    File.AppendAllLines(strResultsPath + "\\ConfigFile.txt", configLines);
                                    break;
                                }//IF Loop
                            }//IF Loop
                        }//IF Loop
                    }//While Loop
                }
                catch (Exception e)
                {
                    file.Close();
                    throw;
                }
                file.Close();
            }
        }

        private void createFile(string filename)
        {
            try
            {
                //if (!File.Exists(filename))
                //{
                    //File.Create(filename);
                    using (FileStream fs = new FileStream(filename, FileMode.OpenOrCreate))
                    using (StreamWriter str = new StreamWriter(fs))
                    {
                        str.BaseStream.Seek(0, SeekOrigin.End);
                        str.Write("");
                        str.Flush();
                    }

                    //TextWriter tw = new StreamWriter(filename);
                    //tw.WriteLine("File Created!");
                    //tw.Close();
                //}
                //else if (File.Exists(filename))
                //{
                //    TextWriter tw = new StreamWriter(filename, true);
                //    tw.WriteLine("//Update here...");
                //    tw.Close();
                //}
            }
            catch (Exception)
            {
                throw;
            }
        }

        //----------------------------------------------UpdateLog--------------------------------------------------//
        private string getConnectionString(string strFileName)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            props["Extended Properties"] = "Excel 12.0 XML";
            props["Data Source"] = strResultsPath + "\\"+ strFileName +".xlsx";

            StringBuilder sb = new StringBuilder();
            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }
            return sb.ToString();
        }

        public void createTCResultLog()
        {
            string connectionString = getConnectionString(strScenario + "_" + strTestCase + "_" + strBrowserName);
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;
                cmd.CommandText = "CREATE TABLE [Results] (SNo INT, StepName VARCHAR, ActualResult VARCHAR, Status VARCHAR, UpdatedOn VARCHAR );";
                cmd.ExecuteNonQuery();
                //cmd.CommandText = "INSERT INTO [table1](id,name,datecol) VALUES(1,'AAAA','2014-01-01');";
                //cmd.ExecuteNonQuery();
                conn.Close();
            }
            initializeTCResultHeader();
        }

        public void initializeSummaryResult()
        {
            iPassedTCCount = 0; iFailedTCCount = 0;
            string connectionString = getConnectionString("Summary");
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;
                cmd.CommandText = "CREATE TABLE [Summary] (SNo INT, ModuleName VARCHAR, TestCase VARCHAR, BrowserName VARCHAR, Status VARCHAR);";
                cmd.ExecuteNonQuery();
                conn.Close();
            }
            initializeSummaryHeader();
        }

        public void initializeSummaryHeader()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            object misValue = System.Reflection.Missing.Value;
            string strFilePath = strResultsPath + "\\Summary.xlsx";
            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(strFilePath);            
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Summary");

            range = xlWorkSheet.get_Range("A1", "F1");
            //range.Merge(true);
            //range.FormulaR1C1 = "Summary Result";
            //range.HorizontalAlignment = 3;
            //range.VerticalAlignment = 3;
            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            range.Font.Size = 18;

            xlApp.UserControl = false;
            xlWorkBook.Save();

            xlWorkBook.Close();
            xlWorkSheet = null;
            xlWorkBook = null;
            xlApp.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public void logSummaryResult()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            object misValue = System.Reflection.Missing.Value;
            string strFilePath = strResultsPath + "\\Summary.xlsx";
            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(strFilePath);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Summary");
            range = xlWorkSheet.UsedRange;
            int iRowCount = range.Rows.Count + 1;
            //add data 
            xlWorkSheet.Cells[iRowCount, 1] = iRowCount - 1;
            xlWorkSheet.Cells[iRowCount, 2] = strScenario;
            xlWorkSheet.Cells[iRowCount, 3] = strTestCase;
            xlWorkSheet.Cells[iRowCount, 4] = strBrowserName;
            xlWorkSheet.Cells[iRowCount, 5] = strResult;
            //if (iFailedStepCount > 0)
            //{
            //    xlWorkSheet.Cells[iRowCount, 5] = "Fail";
            //}
            //else
            //{
            //    xlWorkSheet.Cells[iRowCount, 5] = "Pass";
            //}
            //if (iFailedStepCount > 0)
            if (strResult.Equals("Fail", StringComparison.InvariantCultureIgnoreCase))
            {
                iFailedTCCount++;
                xlWorkSheet.Cells[iRowCount, 5].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);                
            }
            else
            {
                iPassedTCCount++;
                xlWorkSheet.Cells[iRowCount, 5].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
            }

            //Create Hyperlink
            strFilePath = strResultsPath + "\\" + strScenario + "_" + strTestCase + "_" + strBrowserName + ".xlsx";
            //xlWorkSheet.Hyperlinks.Add(xlWorkSheet.Cells[iRowCount, 3], strFilePath, Type.Missing, "View Testcase", strTestCase);
            xlWorkSheet.Hyperlinks.Add(xlWorkSheet.Cells[iRowCount, 3], strScenario + "_" + strTestCase + "_" + strBrowserName + ".xlsx", Type.Missing, "View Testcase", strTestCase);

            xlApp.UserControl = false;
            xlWorkBook.Save();

            xlWorkBook.Close();
            xlWorkSheet = null;
            xlWorkBook = null;
            xlApp.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public void initializeSummaryResultFooter()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            object misValue = System.Reflection.Missing.Value;
            string strFilePath = strResultsPath + "\\Summary.xlsx";
            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(strFilePath); //.Add(misValue);
            //xlWorkBook = xlApp.Workbooks.Open(strFilePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Summary");
            range = xlWorkSheet.UsedRange;
            int iRowCount = range.Rows.Count + 2;

            xlWorkSheet.Cells[iRowCount, 1] = "Passed:";
            xlWorkSheet.Cells[iRowCount, 2] = iPassedTCCount;
            xlWorkSheet.Cells[iRowCount, 3] = "Failed:";
            xlWorkSheet.Cells[iRowCount, 4] = iFailedTCCount;

            range = xlWorkSheet.get_Range("A" + iRowCount, "B" + iRowCount);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
            range = xlWorkSheet.get_Range("C" + iRowCount, "D" + iRowCount);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

            xlApp.UserControl = false;
            xlWorkBook.Save();

            xlWorkBook.Close();
            xlWorkSheet = null;
            xlWorkBook = null;
            xlApp.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

        }

        public void initializeTCResultHeader()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            object misValue = System.Reflection.Missing.Value;
            iPassedStepCount = 0; iFailedStepCount = 0; iStepNo = 0;
            string strFilePath = strResultsPath + "\\" + strScenario + "_" + strTestCase + "_" + strBrowserName + ".xlsx";
            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(strFilePath); //.Add(misValue);
            //xlWorkBook = xlApp.Workbooks.Open(strFilePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Results");

            range = xlWorkSheet.get_Range("A1", "E1");
            //range.FormulaR1C1 = "Execution summary of "+ strTestCase + " in " + strBrowserName + " Browser";
            //range.HorizontalAlignment = 3;
            //range.VerticalAlignment = 3;
            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            range.Font.Size = 14;

            xlApp.UserControl = false;
            xlWorkBook.Save();

            xlWorkBook.Close();
            xlWorkSheet = null;
            xlWorkBook = null;
            xlApp.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            
        }

        public void initializeTCResultFooter()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            object misValue = System.Reflection.Missing.Value;
            string strFilePath = strResultsPath + "\\" + strScenario + "_" + strTestCase + "_" + strBrowserName + ".xlsx";
            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(strFilePath); //.Add(misValue);
            //xlWorkBook = xlApp.Workbooks.Open(strFilePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Results");
            range = xlWorkSheet.UsedRange;
            int iRowCount = range.Rows.Count + 2;

            xlWorkSheet.Cells[iRowCount, 1] = "Passed:";
            xlWorkSheet.Cells[iRowCount, 2] = iPassedStepCount;
            xlWorkSheet.Cells[iRowCount, 3] = "Failed:";
            xlWorkSheet.Cells[iRowCount, 4] = iFailedStepCount;
            if (iFailedStepCount > 0)
            {
                strResult = "Fail";
            }
            else strResult = "Pass";
            range = xlWorkSheet.get_Range("A" + iRowCount, "B" + iRowCount);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
            range = xlWorkSheet.get_Range("C" + iRowCount, "D" + iRowCount);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

            xlApp.UserControl = false;
            xlWorkBook.Save();

            xlWorkBook.Close();
            xlWorkSheet = null;
            xlWorkBook = null;
            xlApp.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

        }

        public void logInfo(String strExpectedStep, string strActualStep, string strStatus)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            object misValue = System.Reflection.Missing.Value;
            string strFilePath = strResultsPath + "\\" + strScenario + "_" + strTestCase + "_" + strBrowserName + ".xlsx";
            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(strFilePath); //.Add(misValue);
            //xlWorkBook = xlApp.Workbooks.Open(strFilePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Results");
            range = xlWorkSheet.UsedRange;
            int iRowCount = range.Rows.Count + 1;
            iStepNo++;
            //add data 
            xlWorkSheet.Cells[iRowCount, 1] = iStepNo;
            xlWorkSheet.Cells[iRowCount, 2] = strExpectedStep;
            xlWorkSheet.Cells[iRowCount, 3] = strActualStep;
            xlWorkSheet.Cells[iRowCount, 4] = strStatus;
            xlWorkSheet.Cells[iRowCount, 5] = DateTime.Now.ToString().Replace('/', '-').Replace(':', '.');

            //Screenshot path
            string screensshotName = DateTime.Now.ToString().Replace('/', '-').Replace(':', '.').Replace(" ", "_");
            screensshotName = strScenario + "_" + strTestCase + "_" + strBrowserName + "_" + screensshotName + ".png";
            strFilePath = strResultsPath + "\\Screenshots\\" + screensshotName; // strScenario + "_" + strTestCase + "_" + strBrowserName + "_" + temp + ".png";
            if (strStatus.Equals("Pass", StringComparison.InvariantCultureIgnoreCase))
            {
                iPassedStepCount++;
                xlWorkSheet.Cells[iRowCount, 4].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
                //Take Screenshot                
                captureScreen(strFilePath);
                //Create Hyperlink
                xlWorkSheet.Hyperlinks.Add(xlWorkSheet.Cells[iRowCount, 4], "Screenshots\\" + screensshotName, Type.Missing, "View screenshot", "Pass");
                //xlWorkSheet.Cells[iRowCount, 4].Formula = "=HYPERLINK(" + strFilePath + " & b1)";
                //xlWorkSheet.Hyperlinks.Add(xlWorkSheet.Cells[iRowCount, 4], strFilePath, Type.Missing, Type.Missing, Type.Missing);
            }
            else if (strStatus.Equals("Fail", StringComparison.InvariantCultureIgnoreCase) || strStatus.Equals("Exception", StringComparison.InvariantCultureIgnoreCase))
            {
                iFailedStepCount++;
                xlWorkSheet.Cells[iRowCount, 4].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                captureScreen(strFilePath);
                xlWorkSheet.Hyperlinks.Add(xlWorkSheet.Cells[iRowCount, 4], strFilePath, Type.Missing, "View screenshot", "Fail");
            }

            xlApp.UserControl = false;
            xlWorkBook.Save();

            xlWorkBook.Close();
            xlWorkSheet = null;
            xlWorkBook = null;
            xlApp.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public void logInfo(String strMethodName)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            object misValue = System.Reflection.Missing.Value;
            string strFilePath = strResultsPath + "\\" + strScenario + "_" + strTestCase + "_" + strBrowserName + ".xlsx";
            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(strFilePath); //.Add(misValue);
            //xlWorkBook = xlApp.Workbooks.Open(strFilePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Results");
            range = xlWorkSheet.UsedRange;
            int iRowCount = range.Rows.Count + 1;
            //add data 
            xlWorkSheet.Cells[iRowCount, 1] = strMethodName;
            xlWorkSheet.Cells[iRowCount, 1].Font.Bold = true;
            xlWorkSheet.Cells[iRowCount, 1].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
            xlWorkSheet.get_Range("A" + iRowCount, "E" + iRowCount).Merge(true);
            xlApp.UserControl = false;
            xlWorkBook.Save();

            xlWorkBook.Close();
            xlWorkSheet = null;
            xlWorkBook = null;
            xlApp.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public void captureScreen(string strFileName)
        {
            try
            {
                Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
                string screenshot = ss.AsBase64EncodedString;
                byte[] screenshotAsByteArray = ss.AsByteArray;
                ss.SaveAsFile(strFileName, ImageFormat.Png);
                ss.ToString();
            }
            catch (Exception e)
            {
                logInfo("Capture Screen",e.Message, "Exception");
            }
        }

        public void consolidateAllScreenshots()
        {
            Word.Application wordApp = new Word.Application();//Create an instance for word app
            wordApp.Visible = false;//Set animation status for word application
            Word.Document doc = wordApp.Documents.Add();

            string screensshotName = DateTime.Now.ToString().Replace('/', '-').Replace(':', '.').Replace(" ", "_");
            screensshotName = strScenario + "_" + strTestCase + "_" + strBrowserName + "_" + screensshotName + ".png";
            string strImageFilesPath = strResultsPath + "\\Screenshots"; 
            foreach (string filePath in Directory.GetFiles(strImageFilesPath, "*.png").Reverse())
            {
                if (System.IO.Path.GetFileName(filePath).StartsWith(strScenario + "_" + strTestCase + "_" + strBrowserName + "_"))
                {
                    doc.InlineShapes.AddPicture(filePath);
                }                
            }
            doc.SaveAs(strImageFilesPath + @"\" + strScenario + "_" + strTestCase + "_" + strBrowserName + ".doc");
            doc.Close();
            wordApp.Quit();

        }

        ////////////////////////////////////-QC Integration-/////////////////////////////////////////
        private static String qcUrl;
        private static String qcDomain;
        private static String qcProject;
        private static String qcLoginName;
        private static String qcPassword;
        private static String testSetPath;
        private static String testSetName;

        public void loadQCVariables()
        {
            Properties ObjQCProp = new Properties(GetProjectPath() + "\\QCDetails.txt");
            qcUrl = ObjQCProp.get("QCURL");
            qcDomain = ObjQCProp.get("QCDomain");
            qcProject = ObjQCProp.get("QCProject");
            qcLoginName = ObjQCProp.get("QCLoginName");
            qcPassword = ObjQCProp.get("QCPassword");
            testSetPath = ObjQCProp.get("QCTestLab");            
        }

        public void qcResultUpdate()
        {     
            testSetName = strTestCaseDescription;
            TDConnection tdConn = new TDConnection();
            try
            {
                tdConn.InitConnectionEx(qcUrl);
                tdConn.ConnectProjectEx(qcDomain, qcProject, qcLoginName, qcPassword);

                TestSetFactory tsFactory = (TestSetFactory)tdConn.TestSetFactory;
                TestSetTreeManager tsTreeMgr = (TestSetTreeManager)tdConn.TestSetTreeManager;
                TestSetFolder tsFolder = (TestSetFolder)tsTreeMgr.get_NodeByPath(testSetPath);
                List tsList = tsFolder.FindTestSets("");
                foreach (TestSet testSet in tsList)
                {
                    TestSetFolder tsFolder1 = (TestSetFolder)testSet.TestSetFolder;
                    TSTestFactory tsTestFactory1 = (TSTestFactory)testSet.TSTestFactory;
                    TDFilter testCaseFilter = tsTestFactory1.Filter as TDFilter;
                    List tsTestList = tsTestFactory1.NewList("");
                    foreach (TSTest tsTest in tsTestList)
                    {
                        //Run lastRun = (Run)tsTest.LastRun;
                        if ((tsTest.Name).Equals(testSetName))
                        {
                            string result= null;
                            if(strResult.Equals("Pass", StringComparison.InvariantCultureIgnoreCase))
                                result = "Passed";
                            else if (strResult.Equals("Fail", StringComparison.InvariantCultureIgnoreCase))
                                result = "Failed";
                            //Run Test
                            String date = DateTime.Now.ToString("yyyy-MM-dd hh.mm.ss");
                            RunFactory runFactory = (RunFactory)tsTest.RunFactory;                            
                            Run run = (Run)runFactory.AddItem("Run" + date);
                            run.Status = result;
                            run.Post();
                            //Add Attachment
                            string filePath = strResultsPath + @"\Screenshots" + @"\";
                            string fileName = strScenario + "_" + strTestCase + "_" + strBrowserName + ".doc";
                            var attachmentFactory = tsTest.Attachments;
                            var attachment = attachmentFactory.AddItem(fileName);
                            attachment.Description = "Auto Upload";
                            attachment.Post();
                            var attachmentStorage = attachment.AttachmentStorage;
                            attachmentStorage.ClientPath = filePath;
                            attachmentStorage.Save(fileName, true);
                            break;
                        }
                    } // end loop of test cases
                }//FOR
            }
            catch (Exception e)
            {
                logInfo("Update QC Results->Exception:"+e.Message);
            }finally{
                tdConn.Logout();
                tdConn.Disconnect();
                tdConn.ReleaseConnection();
            }

        }
    }
}