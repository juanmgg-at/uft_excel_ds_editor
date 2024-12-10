using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HP.LFT.SDK;
using HP.LFT.SDK.StdWin;
using System.Net;

//Libraries of objects
using static ExcelDatasheetEditorUFTProject.UIAProUIObjects;
using static ExcelDatasheetEditorUFTProject.StdWinUIObjects;

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

        //[TestMethod]
        //public void _0001_OpenWorkspace()
        //{

        //    string exePath = @"C:\Program Files\AspenTech\Basic Engineering V15.0\UserServices\bin\AZExplorer.exe";
        //    Utilis.launchExe(exePath);

        //    OpenAspenBasicEngineeringV15WorkspaceDialog.Activate();


        //    cONNECTIONButton.Click();


        //    BrowseButton.Click();


        //    SysTreeView32TreeViewOpenWorkspace.Click();


        //    SysTreeView32TreeViewOpenWorkspace.Select($"Aspen Basic Engineering Workspaces on ({Dns.GetHostEntry(Environment.MachineName).HostName});{EnvironmentSetting.workspace} on {Dns.GetHostEntry(Environment.MachineName).HostName}");


        //    OKButtonOpenWorkspace.Click();


        //    OpenButtonOpenWorkspace.Click();

        //}

        //[TestMethod]
        //public void _0002_CreateCentrifugalPump()
        //{

        //    Thread.Sleep(1000);

        //    AspenBasicEngineeringV15Window.Activate();


        //    SysTreeView32TreeView.Select($@"\\{Dns.GetHostEntry(Environment.MachineName).HostName}\{EnvironmentSetting.workspace}");

        //    SysTreeView32TreeView.SendKeys("b", HP.LFT.SDK.KeyModifier.Alt);


        //    Thread.Sleep(1000);

        //    MenuBarUiObject.SendKeys("c");




        //    SelectClassDialogCreateEquipment.Activate();


        //    ListBoxCreateEquipment.Click();

        //    ListBoxCreateEquipment.Select("Centrifugal Pump");


        //    OKButtonCreateEquipment.Click();

        //}



        //[TestMethod]
        //public void _001_CreateDatasheetsFolder()
        //{

        //    Thread.Sleep(1000);


        //    AspenBasicEngineeringV15Window.Activate();


        //    SysTreeView32TreeView.Select($@"\\{Dns.GetHostEntry(Environment.MachineName).HostName}\{EnvironmentSetting.workspace}");

        //    SysTreeView32TreeView.SendKeys("o", HP.LFT.SDK.KeyModifier.Alt);


        //    Thread.Sleep(1000);

        //    MenuBarUiObject.SendKeys("nf");


        //    EditEditField.SetText("Datasheets");

        //    EditEditField.SendKeys(HP.LFT.SDK.Keys.Return);



        //}


        //[TestMethod]
        //public void _002_CreateDatasheet()
        //{
        //    Thread.Sleep(1000);

        //    AspenBasicEngineeringV15Window.Activate();


        //    SysTreeView32TreeView.Select($@"\\{Dns.GetHostEntry(Environment.MachineName).HostName}\{EnvironmentSetting.workspace}");

        //    SysTreeView32TreeView.SendKeys("b", HP.LFT.SDK.KeyModifier.Alt);


        //    Thread.Sleep(1000);

        //    MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

        //    MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

        //    MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

        //    MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

        //    MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

        //    MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Return);



        //    ListBox.Select("AZ Centrifugal Pump 1");

        //    var OKButton = SelectDatasheetByTypeDialog.Describe<IButton>(new ButtonDescription
        //    {
        //        NativeClass = @"Button",
        //        Text = @"OK"
        //    });

        //    OKButton.Click();


        //    DatasheetDialog.Activate();


        //    SysTabControl32TabControl.Select("New");


        //    SysListView32ListView.Select(1);


        //    OpenButton.Click();


        //}



        [TestMethod]
        public void _003_ModifyCapacityNormal()
        {

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

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("100");


            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);


            editEditField.Select(0, 9);

            editEditField.SetText("field90");

            Thread.Sleep(1000);

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);


            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("g");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Down);
            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Down);
            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Down);
            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Down);


            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);



        }

        [TestMethod]
        public void _004_CheckPumpTypeRadioButton()
        {

            Thread.Sleep(1000);

            editEditField.Select(0, 9);

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

            var comboBox = DesignDatasheetXlsmUserNameCWindow
           .Describe<IComboBox>(new ComboBoxDescription
           {
               NativeClass = @"ComboBox"
           });

            comboBox.Click();

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
            var comboBox = DesignDatasheetXlsmUserNameCWindow
          .Describe<IComboBox>(new ComboBoxDescription
          {
              NativeClass = @"ComboBox"
          });


            comboBox.Click();

            editEditField.SetText("field561");


            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);


        }
        
        [TestMethod]
        public void _0062_EnterLongRemark()
        {

            Thread.Sleep(1000);

            var comboBox = DesignDatasheetXlsmUserNameCWindow
          .Describe<IComboBox>(new ComboBoxDescription
          {
              NativeClass = @"ComboBox"
          });


            comboBox.Click();

            editEditField.SetText("field561");


            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);


            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("This is a comprehensive and detailed remark intended for testing purposes. It includes various aspects of information, ranging from general observations to specific points that could be used in real-world scenarios");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

        }

        [TestMethod]
        public void _007_ClickNonWitCheckbox()
        {

            Thread.Sleep(1000);

            editEditField.Select(0, 3);


            Thread.Sleep(1000);

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

            var OKButton = SelectDatasheetDialog.Describe<IButton>(new ButtonDescription
            {
                NativeClass = @"Button",
                Text = @"OK"
            });

            OKButton.Click();

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
        public void _012()
        {
            //Thread.Sleep(1000);

            //editEditField.Select(0, 3);


            //Thread.Sleep(1000);

            //editEditField.SetText("field78");

            //editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            //float normalCapacity = float.Parse(i23GridItem.TextPattern.VisibleText);


            //Verify.AreEqual(normalCapacity, 100, $"Normal Capacity is {100}");

            //editEditField.Select(0, 3);


            //Thread.Sleep(1000);

            //editEditField.SetText("field260");

            //editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            //string visibleTextUnselected = "r";

            //string visisbleTextSelected = "";

            //string pumpTypeBB1 = field260CImage.GetVisibleText();



            //Verify.AreEqual(pumpTypeBB1, visibleTextUnselected, "Verify Pump Type BB1 is unselected");

            //editEditField.Select(0, 3);


            //Thread.Sleep(1000);

            //editEditField.SetText("field264");

            //editEditField.SendKeys(HP.LFT.SDK.Keys.Return);


            //string pumpTypeBB5 = field264CImage.GetVisibleText();

            //Verify.AreEqual(pumpTypeBB5, visisbleTextSelected, "Verify Pump Type BB5 is selected");

            //Thread.Sleep(1000);


            //editEditField.Select(0, 3);

            //editEditField.SetText("field433");

            //editEditField.SendKeys(HP.LFT.SDK.Keys.Return);


            string selectedString = "";
            //string unselectedString = "□";

            //string granTapsDrain = Field433CImage.GetVisibleText(); // selected

            //Verify.AreEqual(granTapsDrain, selectedString, "Drain Gran Taps is Selected");


            //Thread.Sleep(1000);


            //editEditField.Select(0, 3);

            //editEditField.SetText("field434");

            //editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            //string granTapsCooling = Field434CImage.GetVisibleText();


            //Verify.AreEqual(granTapsCooling, unselectedString, "Cooling Gran Taps is Unselected");

            //Thread.Sleep(1000);


            //editEditField.Select(0, 3);

            //editEditField.SetText("field561");

            //editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            //string expectedLongRemark = "This is a comprehensive and detailed remark intended for testing purposes. It includes various aspects of information, ranging from general observations to specific points that could be used in real-world scenarios";

            //var longRemark = e16GridItem.TextPattern.VisibleText;

            //Verify.AreEqual(longRemark, expectedLongRemark, "Verify long remak is correct");

            editEditField.Select(0, 3);

            editEditField.SetText("field669");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            string actualNonWit = Field669CImage.GetVisibleText();

            Verify.AreEqual(actualNonWit, selectedString, "Verify Non Wit is Selected");

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


