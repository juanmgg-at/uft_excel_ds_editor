using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HP.LFT.SDK;
using HP.LFT.SDK.StdWin;
using System.Net;

//Libraries of objects
using static ExcelDatasheetEditorUFTProject.UIAProUIObjects;
using static ExcelDatasheetEditorUFTProject.StdWinUIObjects;

// Library with constant values 
using static ExcelDatasheetEditorUFTProject.Constants;

using System.Threading;
using HP.LFT.Verifications;

namespace ExcelDatasheetEditorUFTProject
{
    [TestClass]
    public class UftDeveloperTest1 : UnitTestClassBase<UftDeveloperTest1>
    {
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            GlobalSetup(context);
        }

        [TestInitialize]
        public void TestInitialize()
        {

        }


        #region Launch Excel Datasheet Editor

        [TestMethod]
        public void _0001_OpenWorkspace()
        {

            string exePath = @"C:\Program Files\AspenTech\Basic Engineering V15.0\UserServices\bin\AZExplorer.exe";
            Utilis.launchExe(exePath);

            OpenAspenBasicEngineeringV15WorkspaceDialog.Activate();


            cONNECTIONButton.Click();


            BrowseButton.Click();


            SysTreeView32TreeViewOpenWorkspace.Click();


            SysTreeView32TreeViewOpenWorkspace.Select($"Aspen Basic Engineering Workspaces on ({Dns.GetHostEntry(Environment.MachineName).HostName});{EnvironmentSetting.workspace} on {Dns.GetHostEntry(Environment.MachineName).HostName}");


            OKButtonOpenWorkspace.Click();


            OpenButtonOpenWorkspace.Click();

        }

        [TestMethod]
        public void _0002_CreateCentrifugalPump()
        {

            Thread.Sleep(1000);

            AspenBasicEngineeringV15Window.Activate();


            SysTreeView32TreeView.Select($@"\\{Dns.GetHostEntry(Environment.MachineName).HostName}\{EnvironmentSetting.workspace}");

            SysTreeView32TreeView.SendKeys("b", HP.LFT.SDK.KeyModifier.Alt);


            Thread.Sleep(1000);

            MenuBarUiObject.SendKeys("c");




            SelectClassDialogCreateEquipment.Activate();


            ListBoxCreateEquipment.Click();

            ListBoxCreateEquipment.Select("Centrifugal Pump");


            OKButtonCreateEquipment.Click();

        }



        [TestMethod]
        public void _001_CreateDatasheetsFolder()
        {

            Thread.Sleep(1000);


            AspenBasicEngineeringV15Window.Activate();


            SysTreeView32TreeView.Select($@"\\{Dns.GetHostEntry(Environment.MachineName).HostName}\{EnvironmentSetting.workspace}");

            SysTreeView32TreeView.SendKeys("o", HP.LFT.SDK.KeyModifier.Alt);


            Thread.Sleep(1000);

            MenuBarUiObject.SendKeys("nf");


            EditEditField.SetText("Datasheets");

            EditEditField.SendKeys(HP.LFT.SDK.Keys.Return);



        }


        [TestMethod]
        public void _002_CreateDatasheet()
        {
            Thread.Sleep(1000);

            AspenBasicEngineeringV15Window.Activate();


            SysTreeView32TreeView.Select($@"\\{Dns.GetHostEntry(Environment.MachineName).HostName}\{EnvironmentSetting.workspace}");

            SysTreeView32TreeView.SendKeys("b", HP.LFT.SDK.KeyModifier.Alt);


            Thread.Sleep(1000);

            MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

            MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

            MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

            MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

            MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

            MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Return);



            ListBox.Select("AZ Centrifugal Pump 1");

            var SelectDatasheetTypeOKButton = SelectDatasheetByTypeDialog.Describe<IButton>(new ButtonDescription
            {
                NativeClass = @"Button",
                Text = @"OK"
            });

            SelectDatasheetTypeOKButton.Click();


            DatasheetDialog.Activate();


            SysTabControl32TabControl.Select("New");


            SysListView32ListView.Select(1);


            OpenButton.Click();


        }



        [TestMethod]
        public void _003_ModifyCapacityNormal()
        {

            //Thread.Sleep(1000);

            //DesignDatasheetXlsmUserNameCWindow.WaitUntilEnabled();

            //selectFieldComboBox.Click();

            //Thread.Sleep(1000);

            //editEditField.SetText("field90");

            //editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            //Thread.Sleep(1000);

            //AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys(InitialNormalCapacityUnits);

            //eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);


            //selectFieldComboBox.Click();

            //Thread.Sleep(1000);

            //editEditField.SetText("field78");

            //editEditField.SendKeys(HP.LFT.SDK.Keys.Return);


            //Thread.Sleep(1000);

            //AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys(InitialNormalCapacityValue);


            //eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);


            //selectFieldComboBox.Click();

            //editEditField.SetText("field90");

            //editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            //AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("m");

            //eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

         

                Thread.Sleep(1000);

                DesignDatasheetXlsmUserNameCWindow.WaitUntilEnabled();

                editEditField.Select(0, 9);

                Thread.Sleep(5000);

                editEditField.SetText("field90");


                editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

                Thread.Sleep(1000);

                AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("l/h");



                eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);



                editEditField.Select(0, 9);

                Thread.Sleep(1000);

                editEditField.SetText("field78");

                editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

                AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("1000000");


                eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);


                editEditField.Select(0, 9);

                editEditField.SetText("field90");

                Thread.Sleep(1000);

                editEditField.SendKeys(HP.LFT.SDK.Keys.Return);


                AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("m");


                eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);


        }

        [TestMethod]
        public void _004_CheckPumpTypeRadioButton()
        {

            Thread.Sleep(1000);

            selectFieldComboBox.Click();

            editEditField.SetText("field260");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            Thread.Sleep(1000);

            field260CImage.Click();
            field264CImage.Click();


        }

        [TestMethod]
        public void _005_CheckGlanTapsCheckBox()
        {

            Thread.Sleep(1000);

            selectFieldComboBox.Click();

            editEditField.SetText("field433");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);


            Thread.Sleep(1000);

            Field433CImage.Click();
            Field434CImage.Click();
            Field434CImage.Click();

        }

        [TestMethod]
        public void _0061_EnterLongRemark()
        {

            Thread.Sleep(1000);

            selectFieldComboBox.Click();

            editEditField.SetText("field561");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

        }

        [TestMethod]
        public void _0062_EnterLongRemark()
        {
            Thread.Sleep(1000);

            selectFieldComboBox.Click();

            editEditField.SetText("field561");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys(ExpectedLongRemark);

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

        }

        [TestMethod]
        public void _007_ClickNonWitCheckbox()
        {

            Thread.Sleep(1000);

            selectFieldComboBox.Click();

            editEditField.SetText("field699");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            Field669CImage.Click();

        }


        [TestMethod]
        public void _008_CloseDatasheetEditor()
        {
            Thread.Sleep(1000);
            DesignDatasheetXlsmUserNameCWindow.Close();
        }


        [TestMethod]
        public void _009_AddDatasheetToFolder()
        {

            Thread.Sleep(1000);

            AspenBasicEngineeringV15Window.Activate();


            SysTreeView32TreeView.Select($@"\\{Dns.GetHostEntry(Environment.MachineName).HostName}\{EnvironmentSetting.workspace};Datasheets");


            Thread.Sleep(1000);

            SysTreeView32TreeView.SendKeys("o", HP.LFT.SDK.KeyModifier.Alt);

            MenuBarUiObject.SendKeys("t");


            Thread.Sleep(1000);

            SelectDatasheetDialog.Activate();

            ListBoxAddDatasheet.Select(1);

            var SelectDatasheetOKButton = SelectDatasheetDialog.Describe<IButton>(new ButtonDescription
            {
                NativeClass = @"Button",
                Text = @"OK"
            });

            SelectDatasheetOKButton.Click();

        }


        [TestMethod]
        public void _010_OpenDatasheet()
        {

            Thread.Sleep(1000);

            AspenBasicEngineeringV15Window.Activate();

            SysTreeView32TreeViewOpenDatasheet.Select($@"\\{Dns.GetHostEntry(Environment.MachineName).HostName}\{EnvironmentSetting.workspace};All Datasheets");

            Thread.Sleep(1000);


            //Select  method never works, only if you use Click method first

            SysListView32ListViewOpenDatasheet.Click();

            SysListView32ListViewOpenDatasheet.Select(1);

            Thread.Sleep(1000);

            SysListView32ListViewOpenDatasheet.DoubleClick();


            DesignDatasheetXlsmUserNameCWindow.WaitUntilEnabled();

            DesignDatasheetXlsmUserNameCWindow.Activate();

        }

        [TestMethod]
        public void _011_VerifyValuesInDatasheet()
        {
           
            Thread.Sleep(1000);

            Reporter.StartReportingContext("Normal Capacity");

            selectFieldComboBox.Click();

            editEditField.SetText("field78");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            // Get normal flow value and parse it to float
            float normalCapacityValue = float.Parse(i23GridItem.TextPattern.VisibleText);


            // Verify normal capacity value
            Verify.AreEqual(normalCapacityValue, ExpectedNormalCapacityValue, $"Normal Capacity Value is {ExpectedNormalCapacityValue}");

            Thread.Sleep(1000);

            // Get normal capacity units
            string normalCapacityUnits = k23GridItem.TextPattern.VisibleText;

            // Verify units
            Verify.AreEqual(normalCapacityUnits, ExpectedNormalCapacityUnits, $"Normal Capacity Units are {ExpectedNormalCapacityUnits}");

            Reporter.EndReportingContext();

            Thread.Sleep(1000);

            Reporter.StartReportingContext("Pump Type");

            selectFieldComboBox.Click();

            editEditField.SetText("field260");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            // Get Pump Type BB1 visible text
            string pumpTypeBB1 = field260CImage.GetVisibleText();

            // Verify pump type BB1 is unselected
            Verify.AreEqual(pumpTypeBB1, PumpTypeUnselected, "Pump Type BB1 is unselected");


            Thread.Sleep(1000);

            selectFieldComboBox.Click();

            editEditField.SetText("field264");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            // Get pump type BB5 visible text
            string pumpTypeBB5 = field264CImage.GetVisibleText();

            // Verify pump type BB5 is selected
            Verify.AreEqual(pumpTypeBB5, PumpTypeSelected, "Pump Type BB5 is selected");

            Reporter.EndReportingContext();

            Thread.Sleep(1000);

            Reporter.StartReportingContext("Gran Taps");

            selectFieldComboBox.Click();

            editEditField.SetText("field434");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            // Get gran taps cooling
            string granTapsCooling = Field434CImage.GetVisibleText();


            // Verify gran taps cooling is unselected
            Verify.AreEqual(granTapsCooling, GranTapUnselectedString, "Cooling Gran Taps is Unselected");

            Thread.Sleep(1000);

            selectFieldComboBox.Click();

            editEditField.SetText("field433");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            // Get gran taps drain
            string granTapsDrain = Field433CImage.GetVisibleText();

            // Verify gran taps drain  is selected
            Verify.AreEqual(granTapsDrain, GranTapsSelected, "Gran Taps Drain is Selected");


            Reporter.EndReportingContext();


            Thread.Sleep(1000);

            Reporter.StartReportingContext("Long Remark");

            selectFieldComboBox.Click();

            editEditField.SetText("field561");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);


            // Get remarks 
            var longRemark = e16GridItem.TextPattern.VisibleText;


            // Verify remarks are Correct
            Verify.AreEqual(longRemark, ExpectedLongRemark, "Long remak is correct");

            Reporter.EndReportingContext();

            Thread.Sleep(1000);

            Reporter.StartReportingContext("Non Wit");

            selectFieldComboBox.Click();

            editEditField.SetText("field669");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            // Get non wit
            string actualNonWit = Field669CImage.GetVisibleText();

            // Verify non wit is slected
            Verify.AreEqual(actualNonWit, GranTapsSelected, "Verify Non Wit is Selected");

            Reporter.EndReportingContext();

        }


        #endregion


        [TestCleanup]
        public void TestCleanup()
        {
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            GlobalTearDown();
        }
    }



}


