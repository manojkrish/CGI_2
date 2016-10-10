using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TDAPIOLELib;

namespace BCAFT
{
    class QCIntegration
    {
        public static String qcUrl = "http://10.35.6.16:8080/qcbin";
        public static String qcDomain = "LOB";
        public static String qcProject = "MyPerformancePortfolio";
        public static String qcLoginName = "vithyasankaran";
        public static String qcPassword = "";

        public static String testSetPath = @"ROOT\Unattached\default";
        public static String testSetName = "[1]Test the funcvtionality of login screen";

        private TDConnection connection;
        //[1]Test the funcvtionality of login screen

        public static void mail()
        {
            QCIntegration obj = new QCIntegration();
            obj.Connect(qcUrl, qcDomain, qcProject, qcLoginName, qcPassword);
            TestSet testSet = obj.GetTestSet(testSetPath, testSetName);
            obj.RunTestSet(testSet);
        }

        public TDConnection Connect(string qcUrl, string qcDomain, string qcProject, string qcLoginName, string qcPassword)
        {
            connection = new TDConnection();
            connection.InitConnectionEx(qcUrl);
            connection.ConnectProjectEx(qcDomain, qcProject, qcLoginName, qcPassword);
            return connection;
        }

        public TestSet GetTestSet(String path, String testSetName)
        {
            TestSetFactory testSetFactory = connection.TestSetFactory;
            TestSetTreeManager testSetTreeManager = connection.TestSetTreeManager;

            TestSetFolder testSetFolder = (TestSetFolder)testSetTreeManager.NodeByPath[path];
            List testSetList = testSetFolder.FindTestSets(testSetName);
            TestSet testSet = testSetList[0];

            return testSet;
        }

        public void RunTestSet(TestSet testSet)
        {
            TSScheduler scheduler = testSet.StartExecution("");
            scheduler.RunAllLocally = true;
            scheduler.Run();
        }

    }
}