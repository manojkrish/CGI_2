using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using BCAFT;
using System.IO;
namespace BCAFT.BusinessComponents
{
    public class FunctionalComponents : ConfigClass
    {
       
       CommonFunctions objComm;
       CommonLibraries objCommLib;
       public static string ProjectPath;

        public FunctionalComponents(HelperClass helperClass) : base(helperClass) {
            objComm = new CommonFunctions(helperClass);
            objCommLib = new CommonLibraries(helperClass);
        }
        //Property file initialization
        public Properties ObjCJProp = new Properties(GetPropertiesFilePath() + "\\CustomerJourney.txt");
        public Properties ObjB2BProp = new Properties(GetPropertiesFilePath() + "\\B2B.txt");
        
        //Initialize Test data sheet
        public void Initialize(string strSheetName)
        {
            if (!arrListSheetsqueried.Contains(strSheetName + strScenario + strTestCase + strStartIteration))
            {
                objCommLib.readTestData(null, strSheetName, strStartIteration);
                arrListSheetsqueried.Add(strSheetName + strScenario + strTestCase + strStartIteration);
            }          
        }

        public void launchORSCJ()
        {
            Properties obj = new Properties(strResultsPath + "\\ConfigFile.txt");
            driver.Navigate().GoToUrl(obj.get("CJ_URL"));//"http://g1gsn0bms057.britishcouncil.org:12223/"
            objComm.wait(By.XPath(ObjCJProp.get("Country.xpath")), 20); ////select[contains(@id,'ddlCountry')]
            objCommLib.logInfo("Launch ORS Customer Registration portal","ORS CJ launched successfully", "Pass");
            driver.Manage().Window.Maximize();
        }

        public void addNewRegistration_CJ()
        {
            Initialize("InputData");
            //Select Country
            objComm.select(By.XPath(ObjCJProp.get("Country.xpath")), "Country");
            //Click on Continue button
            objComm.click(By.XPath("//input[contains(@id,'imgbRegisterBtn')]"), "Continue button"); ////input[contains(@id,'imgbRegisterBtn')]
            objComm.wait(By.XPath("//select[contains(@id,'ddlDateMonthYear')]"), 10);
            //Select Month Year
            objComm.select(By.XPath("//select[contains(@id,'ddlDateMonthYear')]"), "MonthYear Dropdown", 0);
            //Select City
            objComm.select(By.XPath("//select[contains(@id,'ddlTownCityVenue')]"), "City");
            //Select Module
            objComm.select(By.XPath("//select[contains(@id,'ddlModule')]"), "ExamType");
            //Click on Find button
            objComm.click(By.XPath("//input[contains(@id,'imgbSearch')]"), "Find button");
            objComm.wait(By.XPath("//input[contains(@id,'btnBook')]"), 10);
            //Click on Apply button
            objComm.click(By.XPath("//input[contains(@id,'btnBook')]"), "Apply button");
            objComm.wait(By.XPath("//input[contains(@id,'chkAccept')]"), 10);
            //Select Agree checkbox
            IWebElement eLe = objComm.find(By.XPath("//input[contains(@id,'chkAccept')]"), "Agree checkbox");
            if (eLe != null)
                eLe.Click();
            //Click on Continue button
            objComm.click(By.XPath("//input[contains(@id,'imgbContinue')]"), "Continue button");
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));

            //Enter Mandate Data
            enterTestTakerDetails_CJ();

            //Click on Continue button
            objComm.click(By.XPath("//input[contains(@id,'ibtnContinue')]"), "Continue button");
            objComm.wait(By.XPath("//input[contains(@id,'ibtnBookNow')]"), 10);

            //Choose your Payment method
	        eLe = objComm.find(By.XPath("//select[contains(@id,'ddlPaymentMethod')]"),"Choose your Payment method Dropdown");
	        if(eLe != null){
                objComm.select(By.XPath("//select[contains(@id,'ddlPaymentMethod')]"), "Payment Option","Pay Later");		        
		        objComm.sleep(5);
	        }
	        //Click on Apply Now button
            objComm.click(By.XPath(ObjCJProp.get("ApplyNowBtn.Xpath")), "Apply Now button"); //"//input[contains(@id,'ibtnBookNow')]"
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
        }

        public void enterTestTakerDetails_CJ()
        {
            Initialize("InputData");
            string value = "";
            //Select Title
            objComm.select(By.XPath("//select[contains(@id,'ddlTitle')]"), "Title");
            //Enter FirstName
            objComm.enter(By.XPath("//input[contains(@id,'txtOtherNames')]"), "FirstName");
            //Enter LastName
            objComm.enter(By.XPath("//input[contains(@id,'txtFamilyName')]"), "LastName");
            //Select First Language
            objComm.select(By.XPath("//select[contains(@id,'ddlFirstLanguage')]"), "FirstLanguage Dropdown", 2);
            //Select Country of Nationality
            objComm.select(By.XPath("//select[contains(@id,'ddlCountryRegion')]"), "Country of Nationality Dropdown", 2);
            //Email Address
            value = objComm.generatePIN();
            //value = "AutoTest." + value + "@gmail.com";
            objComm.enter(By.XPath("//input[contains(@id,'txtEmail')]"), "Email");
            //Confirm Mail
            objComm.enter(By.XPath("//input[contains(@id,'txtEmailConfirm')]"), "Email");
            //Select DOB
            objComm.select(By.XPath("//select[contains(@id,'ddlDays')]"), "DOB-Date");
            objComm.select(By.XPath("//select[contains(@id,'ddlMonths')]"), "DOB-Month");
            objComm.select(By.XPath("//select[contains(@id,'ddlYears')]"), "DOB-Year");
            //Select Identification Document
            objComm.select(By.XPath("//select[contains(@id,'ddlIdDocument')]"), "Identification Document Dropdown", "Passport");
           
            //Enter Address
            objComm.enter(By.XPath("//input[contains(@id,'txtAddr1')]"), "Address1", "12 West Str");
            //Enter Addr2
            objComm.enter(By.XPath("//input[contains(@id,'txtAddr2')]"), "Address2", "Church Cross Lane");
            //Select Country
            objComm.select(By.XPath("//select[contains(@id,'ddlAddrCountry')]"), "Country Dropdown", 3);
            //Select Occupation Sector
            objComm.select(By.XPath("//select[contains(@id,'ddlOccupationSector')]"), "Occupation Sector Dropdown", 2);
            //Select Occupation Level
            objComm.select(By.XPath("//select[contains(@id,'ddlOccupationStatus')]"), "Occupation Level Dropdown", 2);
            //Y are you Tacking Test
            objComm.select(By.XPath("//select[contains(@id,'ddlReasonForTest')]"), "Reason for taking Test Dropdown", 2);
            //Dest Country
            objComm.select(By.XPath("//select[contains(@id,'ddlDestinationCountry')]"), "Dest Country Dropdown", 2);
            //Edu Completed
            objComm.select(By.XPath("//select[contains(@id,'ddlEducationLevel')]"), "Edu Level Dropdown", 2);
            //year Eng
            objComm.select(By.XPath("//select[contains(@id,'ddlEnglishStudyInYears')]"), "EnglishStudyInYears Dropdown", 2);
            //Identification Document No
            value = objComm.generatePIN();
            value = "UU7234" + value;
            objComm.enter(By.XPath("//input[contains(@id,'txtIdDocumentNumber')]"), "DocumentNo", value);
            //Select ID Doc Exp Date
            objComm.select(By.XPath("//select[contains(@id,'ddlDocIdDay')]"), "IDExp-Date");
            objComm.select(By.XPath("//select[contains(@id,'ddlDocIdMonth')]"), "IDExp-Month");
            objComm.select(By.XPath("//select[contains(@id,'ddlDocIdYear')]"), "IDExp-Year");
        }

        //-------------------------------------------------ORS B2B--------------------------------------------////
        public void loginORSB2B()
        {
            Properties obj = new Properties(strResultsPath + "\\ConfigFile.txt");
            driver.Navigate().GoToUrl(obj.get("B2B_URL"));           
            objCommLib.logInfo("Launch ORS Partnership Programme portal", "ORS B2B portal launched successfully", "Pass");
            driver.Manage().Window.Maximize();
            //Login
            objComm.enter(By.XPath(ObjB2BProp.get("Username.Xpath")), "UserName", obj.get("B2B_UserName"));
            objComm.enter(By.XPath(ObjB2BProp.get("Password.Xpath")), "UserName", obj.get("B2B_Password"));
            objComm.click(By.XPath(ObjB2BProp.get("Login.Xpath")), "Login button");
        }

        public void addNewRegistration_B2B(){
            Initialize("InputData");
    	    //Navigate to Candidates tab
             objComm.click(By.XPath("//a[contains(@href,'MyCandidates')]"), "Candidates tab");
            
    	    //navigateToTab(By.XPath("//div[contains(@id,'mbrMaster_dTopMenuBar')]//ul"),"Candidates");   
    	    objComm.sleep(15);
    	    //Click on AddNew
            objComm.click(By.XPath("//div[contains(@id,'ContentPlaceHolder1_mbrNewCandidates_dTopMenuBar')]//span[contains(text(),'Add New')]"), "AddNew button");
    	    //navigateToTab(By.XPath("//div[contains(@id,'ContentPlaceHolder1_mbrNewCandidates_dTopMenuBar')]//ul"),"Add New...");  
    	    objComm.sleep(10);
    	
    	    //Accept Terms And Condition
            objComm.click(By.XPath("//input[contains(@id,'ucTermsAndConditions_chkAccept')]"), "Accept Terms And Condition");
	        objComm.sleep(3);
	        //Accept Confirmation dialog
	        IWebElement parentDiv = objComm.find(By.XPath("//div[contains(@class,'ui-dialog-buttonpane ui-widget-content')]"), "Parent div");
	        if(parentDiv != null){
                objComm.click(parentDiv, By.TagName("button"), "OK button on Confirmation dialog");
	        }
    	    //Enter Candidate & Test details
            enterTestTakerDetails_B2B();
            //Accept Confirmation dialog
            parentDiv = objComm.find(By.XPath(ObjB2BProp.get("RegistrationConfirmationDialog.Xpath")), "Parent div");
            if (parentDiv != null)
            {
                objComm.click(parentDiv, By.XPath("//button[contains(@class,'ui-button ui-widget')]//span[contains(text(),'Close')]"), "Cancel button on Confirmation dialog");
            }
        }

        public void navigateToTab_B2B(By by, String strTabName){
        Boolean linkFound = false;
        IWebElement parentEle = objComm.find(by, "TopMenu bar");
        if(parentEle != null){
            try{
    	        IList<IWebElement> liList = parentEle.FindElements(By.TagName("li"));
                    for (int i = 0; i < liList.Count; i++)
                {
                        if((liList.ElementAt(i).Text).Equals(strTabName, StringComparison.InvariantCultureIgnoreCase)){
    			        liList.ElementAt(i).Click();
    			        linkFound = true;
    			        break;
    		        }

                    }
            }catch(Exception e){
    	        objCommLib.logInfo("Navigate to Tab","Exception:"+e.Message,"Fail");
            }
        }
        if(!linkFound)
            objCommLib.logInfo("Navigate to Tab", strTabName + " Tab Not Found", "Fail");
        }

        public void enterTestTakerDetails_B2B()
        {
            objComm.select(By.XPath("//select[contains(@id,'ddlExamVenue_CDNew')]"), "City"); 
            objComm.select(By.XPath("//select[contains(@id,'ddlExamModule_CDNew')]"), "ExamType");
            objComm.select(By.XPath("//select[contains(@id,'ddlExamDate_CDNew')]"), "ExamDate", 1);
            objComm.enter(By.XPath("//input[contains(@id,'txtFamilyName_CDNew_txtEntryControl')]"), "FirstName");
            objComm.select(By.XPath("//select[contains(@id,'ddlTitle_CDNew_ddlEntryControl')]"), "Title", 1);
            objComm.enter(By.XPath("//input[contains(@id,'txtOtherName_CDNew_txtEntryControl')]"), "LastName");
            objComm.enter(By.XPath("//input[contains(@id,'txtAddress1_CDNew_txtEntryControl')]"), "Address1", "12 West Street");
            objComm.select(By.XPath("//select[contains(@id,'ddlCountry_CDNew_ddlEntryControl')]"), "Country");
            objComm.enter(By.XPath("//input[contains(@id,'txtEmailAddress_CDNew_txtEntryControl')]"), "Email");
            string strDate = objComm.formatToDate(objCommLib.readTestData("DOB"));
            objComm.enter(By.XPath("//input[contains(@id,'txtDOB_CDNew_txtEntryControl')]"), "DOB", strDate);
            objComm.select(By.XPath("//select[contains(@id,'ddlIdDocument_CDNew_ddlEntryControl')]"), "Identification Document", "Passport");
            objComm.enter(By.XPath("//input[contains(@id,'txtIDNumber_CDNew_txtEntryControl')]"), "Document No", "PA" + objComm.generatePIN());
            strDate = objComm.formatToDate(objCommLib.readTestData("DocExpDate"));
            objComm.enter(By.XPath("//input[contains(@id,'txtIDExpiryDate_CDNew_txtEntryControl')]"), "IDExp Date", strDate);
            objComm.select(By.XPath("//select[contains(@id,'ddlCountryOrigin_CDNew_ddlEntryControl')]"), "Country origin", 1);
            objComm.select(By.XPath("//select[contains(@id,'ddlFirstLanguage_CDNew_ddlEntryControl')]"), "FirstLang", 1);
            objComm.select(By.XPath("//select[contains(@id,'ddlOccupationSector_CDNew_ddlEntryControl')]"), "OcupationSectior", 1);
            objComm.select(By.XPath("//select[contains(@id,'ddlOccupationStatus_CDNew_ddlEntryControl')]"), "OcupationStatus", 1);
            objComm.select(By.XPath("//select[contains(@id,'ddlTestReason_CDNew_ddlEntryControl')]"), "Test Reason", 1);
            objComm.select(By.XPath("//select[contains(@id,'ddlCountryApplyingTo_CDNew_ddlEntryControl')]"), "ApplyingTo", 1);
            objComm.select(By.XPath("//select[contains(@id,'ddlEducationLevel_CDNew_ddlEntryControl')]"), "EducationLevel", 1);
            objComm.select(By.XPath("//select[contains(@id,'ddlYearsEnglishStudy_CDNew_ddlEntryControl')]"), "Year English Studied", 2);
            objComm.click(By.XPath("//input[contains(@id,'ContentPlaceHolder1_dialogCandidateNew_bSave')]"), "Save button");
        }


        //-------------------------------------------------UKVI--------------------------------------------////       
        public void login_UKVI()
        {
            Properties obj = new Properties(strResultsPath + "\\ConfigFile.txt");
            driver.Navigate().GoToUrl(obj.get("UKVICJ_URL")); //"http://ieltsukvisas-uat.britishcouncil.org/"
            driver.Manage().Window.Maximize();
          
            objComm.wait(By.LinkText("Log In"), 15);
            objComm.click("Log In");
            objComm.wait(By.XPath("//form//input[contains(@ng-model,'vm.loginUserModel.email')]"), 10);
            objComm.enter(By.XPath("//form//input[contains(@ng-model,'vm.loginUserModel.email')]"), obj.get("Username"));
            objComm.enter(By.XPath("//form//input[@type='password']"), obj.get("Username"));
            objComm.click(By.XPath("//input[@value='Log in']"), "Log in");
        }

        //Add New registration - UKVI_CJ
        public void chooseTest()
        {
            Initialize("InputData");
            //Change Country
            objComm.waitTillPageload_UKVI(20);
            objComm.wait(By.LinkText("Change country"), 10);
            objComm.click("Change country");
            objComm.sleep(5);
            objComm.waitTillPageload_UKVI(40);
            objComm.wait(By.XPath("//select[@ng-model='vm.selectedCountry']"), 10);
            objComm.select(By.XPath("//select[@ng-model='vm.selectedCountry']"), "Country");
            objComm.click(By.XPath("//a[@class='cc-link'][text()='Close']"), "Notification Close");
            objComm.sleep(5);
            clickButton(driver, By.XPath("//div[starts-with(@class,'btn-group')]//button[@class='btn btn-primary btn-block']"), "Continue");
            objComm.sleep(20);
            objComm.wait(By.XPath("//div[@ng-bind-html='vm.countryInformation']"), 70);
            objComm.waitTillPageload_UKVI(20);
            clickButton(driver, By.XPath("//div[starts-with(@class,'btn-group')]//button[@class='btn btn-primary btn-block']"), "Start now");
            //Select Exam Type
            string strExamType = "IELTS for UKVI (General Training)";
            if (!selectExamType(strExamType))
                Console.WriteLine("Exam type " + strExamType + "not available");
            else
            {
                objComm.sleep(3);
                clickButton(driver, By.XPath("//div[starts-with(@class,'btn-group')]//button[@class='btn btn-primary']"), "Next");
                objComm.wait(By.XPath("//select[@ng-model='vm.selectedLocation']"), 10);
                objComm.select(By.XPath("//select[@ng-model='vm.selectedLocation']"), "Select Location", "Any");
                objComm.select(By.XPath("//select[@ng-model='vm.selectedMonth']"), "Select Month", 1);
                objComm.sleep(5);
                clickButton(driver, By.XPath("//div[starts-with(@class,'btn-group')]//button[@class='btn btn-primary']"), "Next");
                //objComm.wait(By.XPath("//*[starts-with(@class,'page-header')][@text='Choose your test date:']"), 20);
                objComm.sleep(15);
                clickButton(driver, By.XPath("//table[starts-with(@class,'table')]//button[@class='btn btn-primary btn-sm']"), "Select date");
                objComm.sleep(5);
                clickButton(driver, By.XPath("//table[starts-with(@class,'table')]//button[@class='btn btn-primary btn-sm']"), "Book Now");
                objComm.waitTillPageload_UKVI(20);//The Below two steps are inconsistant
                objComm.click(By.XPath("//input[@type='checkbox'][contains(@name,'userHasAccepted')]"), "Ack Checkbox");
                objComm.sleep(1);
                clickButton(driver, By.XPath("//div[starts-with(@class,'btn-group')]//button[@class='btn btn-primary']"), "Next");
            }
        }

        private Boolean clickButton(IWebDriver driver, By by, string strButtonText)
        {
            Boolean btnFound = false; string strActualBtnText = "";
            IList<IWebElement> btnList = objComm.findElements(by, strButtonText + " button");
            if (btnList != null && btnList.Count > 0)
            {
                objComm.waitTillPageload_UKVI(20);
            }
            btnList = objComm.findElements(by, strButtonText + " button");
            for (int i = 0; i < btnList.Count; i++)
            {
                try
                {
                    strActualBtnText = ((btnList.ElementAt(i).Text).Replace(" ", "")).Replace("\n", "");
                    Console.WriteLine("Btn Txt:" + strActualBtnText);
                    if ((btnList.ElementAt(i).Displayed) && (strActualBtnText.Equals(strButtonText.Replace(" ", ""), StringComparison.InvariantCultureIgnoreCase)))
                    {
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", btnList.ElementAt(i));
                        objComm.waitForElementToBeClickable( btnList.ElementAt(i), 20, strButtonText + " button");
                        btnList.ElementAt(i).Click();
                        btnFound = true;
                        break;
                    }
                }
                catch (Exception e)
                {
                    // TODO Auto-generated catch block
                    objCommLib.logInfo("Find Button", "Exception when try to find button -" + strButtonText + ".Exception:"+e.Message, "Fail");
                    //Console.WriteLine(e.Message);
                }
            }//FOR Loop
            if (!btnFound)
            {
                //tearDown();
                objCommLib.logInfo("Find Button","Button -" + strButtonText + "- not found","Fail");
            }
            return btnFound;
        }

        private Boolean selectExamType(string strExamType)
        {
            Boolean isExamSelected = false;
            objComm.waitTillPageload_UKVI(25);
            IList<IWebElement> eleExamLbl = objComm.findElements(By.XPath("//div[contains(@ng-repeat,'examType')]//label[contains(@class,'control-label')]"), "Examtype label");
            for (int i = 0; i < eleExamLbl.Count; i++)
            {
                try
                {
                    if (((eleExamLbl.ElementAt(i).Text).Replace(" ", "")).Replace("\n", "").Equals(strExamType.Replace(" ", ""), StringComparison.InvariantCultureIgnoreCase))
                    {
                        objComm.click(eleExamLbl.ElementAt(i), By.XPath("//input[@type='radio'][@name='examType']"), strExamType + " radio button");
                        isExamSelected = true;
                        break;
                    }
                }
                catch (Exception e)
                {
                    // TODO Auto-generated catch block
                    objCommLib.logInfo("Choose ExamType", "Exception when try to find ExamType. Exception:" + e.Message, "Fail");
                }
            }//FOR Loop
            return isExamSelected;
        }

    }
}
