using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TDAPIOLELib;

namespace BCAFT
{
    class QCI
    {
//        public void sendRequest(String strTestCaseId, String strStatus) {
        public void sendRequest()
        {
        ITDConnection4 connection=null;
        
        //QC url
        String url = "http://10.35.6.16:8080/qcbin";
        //username for login
        String username = "vithyasankaran";
        //password for login
        String password = "";
        //domain
        String domain = "LOB";
        
        //project
        String project = "MyPerformancePortfolio";
        String strTestLabPath  = "Root";
        String strTestSetName = "Test the funcvtionality of login screen";
        
        try{
           
            //QC Connection
            connection = new TDConnection();
            connection.InitConnectionEx(url);
            connection.Login(username, password);
                
            ////To get all projects name
            //for (Com4jObject obj : connection.getAllVisibleProjectDescriptors()) {
            //    IProjectDescriptor pd = obj.queryInterface(IProjectDescriptor.class);
                      
            //}
            
            connection.Connect(domain, project);
            
            //To get the Test Set folder in Test Lab        
            TestSetTreeManager objTestSetTreeManager = (connection.TestSetTreeManager); //  TestSetTreeManager);
            //TestSetFolder objTestSetFolder = (TestSetFolder)objTestSetTreeManager.NodeByPath(strTestLabPath);
            TestSetFolder objTestSetFolder = (TestSetFolder)objTestSetTreeManager.NodeByPath[strTestLabPath];        
            //IList tsTestList = objTestSetFolder.FindTestSets(null, true, null);
            TestFactory tstF = connection.TestFactory as TestFactory;
            TDFilter testCaseFilter = tstF.Filter as TDFilter;
            testCaseFilter["TS_TEST_NAME"] = strTestSetName;
            List testsList = tstF.NewList(testCaseFilter.Text);
            if (testsList != null && testsList.Count == 1)
            {
                //log.log("Test case " + testCaseName + " was found ");
                //Test tmp = testsList[0] as Test;
                RunFactory runFactory = (RunFactory)testsList[0].RunFactory;
                String date = DateTime.Now.ToString("yyyyMMddhhmmss");
                Run run = (Run)runFactory.AddItem("Run" + date);
                run.Status = "Pass";
                run.Post();
                //return (int)tmp.ID;
            }
            //for (int i=1;i<=tsTestList.Count;i++) {
            //    Com4jObject comObj = (Com4jObject) tsTestList.item(i);
            //    ITestSet tst = comObj.queryInterface(ITestSet); 
                        
            //    if(tst.name().equalsIgnoreCase(strTestSetName)){
                            
            //        IBaseFactory testFactory = tst.tsTestFactory().queryInterface(IBaseFactory.class);
              
            //        IList testInstances = testFactory.newList("");
                                
            //        //To get Test Case ID instances
            //        for (Com4jObject testInstanceObj : testInstances){  
            //            ITSTest testInstance = testInstanceObj.queryInterface(ITSTest.class);  
                                    
            //            if(testInstance.testName().equalsIgnoreCase(strTestCaseId)){
            //                IRunFactory runfactory = testInstance.runFactory().queryInterface();
                        
            //                IRun run= runfactory.addItem("Selenium").queryInterface(IRun.class);
            //                run.status(strStatus);
            //                run.post();
            //                break;
            //            }
            //        } 
            //    }
            //}
        }catch(Exception e){
            //System.out.println(e.getMessage());
        }
        //finally{
        //    connection.logout();
        //    connection.disconnect();
        //    connection.releaseConnection();
        //}
    }


    }
}
