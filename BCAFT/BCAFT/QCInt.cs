using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TDAPIOLELib;

namespace BCAFT
{
    class QCInt
    {
        public static String qcUrl = "http://10.35.6.16:8080/qcbin";
        public static String qcDomain = "LOB";
        public static String qcProject = "MyPerformancePortfolio";
        public static String qcLoginName = "vithyasankaran";
        public static String qcPassword = "";

        public static String testSetPath = @"Root\test";
        public static String testSetName = "[1]Test the funcvtionality of login screen";//Test the funcvtionality of login screen
        public void qcResultUpdate()
        {
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
                            RunFactory runFactory = (RunFactory)tsTest.RunFactory;
                            String date = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                            Run run = (Run)runFactory.AddItem("Run" + date);
                            run.Status = "Passed";
                            run.Post();
                            break;
                        }
                    } // end loop of test cases
                }//FOR
            }
            catch (Exception)
            {
                
                throw;
            }

        }
    }
}
