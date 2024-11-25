using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HP.LFT.SDK;
using HP.LFT.SDK.StdWin;
using System.Net;

//Libraries of objects
using static UftDeveloperTestProject1.UIAProUIObjects;
using static UftDeveloperTestProject1.StdWinUIObjects;

using System.Threading;

namespace UftDeveloperTestProject1
{
    [TestClass]
    public class UftDeveloperTest : UnitTestClassBase<UftDeveloperTest>
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

        #region Launch Excel Datasheet Editor


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

            var OKButton = SelectDatasheetByTypeDialog.Describe<IButton>(new ButtonDescription
            {
                NativeClass = @"Button",
                Text = @"OK"
            });

            OKButton.Click();

            
            DatasheetDialog.Activate();

            
            SysTabControl32TabControl.Select("New");

           
            SysListView32ListView.Select(1);

            
            OpenButton.Click();


        }

        

        [TestMethod]
        public void _003_ModifyCapacityNormal()
        {

            Thread.Sleep(1000);

            AZCentrifugalPump1DesignDatasheetXlsmUserNameCWindow.WaitUntilEnabled();

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
            field264CImage.Click();


        }

        [TestMethod]
        public void _005_CheckGlanTapsCheckBox()
        {

            Thread.Sleep(1000);
           

            editEditField.Select(0, 9);

            editEditField.SetText("field433");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);


            Thread.Sleep(1000);

            Field433CImage.Click();
            Field434CImage.Click();
            Field434CImage.Click();

        }


        //[TestMethod] 
        public void _006_EnterLongRemark()
        {

            Thread.Sleep(5000);

            AZCentrifugalPump1DesignDatasheetXlsmUserNameCWindow.WaitUntilEnabled();

            editEditField.Select(0, 9);

         

            editEditField.SetText("field561");


            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("Lorem ipsum dolor sit amet, consectetur adipiscing elit.Suspendisse quam enim, faucibus non aliquam aliquet, tempor vitae libero.Vivamus faucibus in risus ac semper.Suspendisse et erat in lorem dapibus facilisis.Vestibulum lobortis a erat eu varius.Nulla placerat viverra fringilla.Nam auctor vehicula sem.Duis laoreet quam nec sem consectetur, aliquam interdum arcu bibendum.Donec vel tincidunt metus, nec sagittis purus.In id eros ac nulla eleifend varius.Generados 1 párrafos, 69 palabras, 467 bytes de Lorem");


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

            AZCentrifugalPump1DesignDatasheetXlsmUserNameCWindow.Close();


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

            SysListView32ListViewOpenDatasheet.DoubleClick();


        }



        #endregion


        #region Continuous List Datasheet

        [TestMethod]
        public void _011_CreateShellAndTubeHeatExchangers()
        {
            AspenBasicEngineeringV15Window.Activate();
            
            for (int i = 0; i < 4; i++)
            {
                SysTreeView32TreeView.Select($@"\\{Dns.GetHostEntry(Environment.MachineName).HostName}\{EnvironmentSetting.workspace}");

                Thread.Sleep(1000);

                SysTreeView32TreeView.SendKeys("b", HP.LFT.SDK.KeyModifier.Alt);


                Thread.Sleep(1000);

                MenuBarUiObject.SendKeys("c");
                
                SelectClassDialogCreateEquipment.Activate();
                
                ListBoxCreateEquipment.Select("Shell And Tube Heat Exchanger");
                
                OKButtonCreateEquipment.Click();
            }
        }



        [TestMethod]
        public void _012_CreateATContinuousListShellAndTubeExchangers()
        {


            Thread.Sleep(1000);

            AspenBasicEngineeringV15Window.Activate();


            SysTreeView32TreeView.Select($@"\\{Dns.GetHostEntry(Environment.MachineName).HostName}\{EnvironmentSetting.workspace}");

            Thread.Sleep(1000);

            SysTreeView32TreeView.SendKeys("b", HP.LFT.SDK.KeyModifier.Alt);


            Thread.Sleep(1000);

            MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

            MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

            MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

            MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

            MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);

            MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            
            SelectDatasheetByTypeDialog.Activate();


            ListBox.Click();

            ListBox.Select("AT Continuous List Shell And Tube Heat Exchangers");


            var OKButton = SelectDatasheetByTypeDialog.Describe<IButton>(new ButtonDescription
            {
                NativeClass = @"Button",
                Text = @"OK"
            });

            OKButton.Click();

           
            DatasheetDialog.Activate();
            

            SysTabControl32TabControl.Select("New");



            SysListView32ListView.Click();

            SysListView32ListView.Select("AT Continuous List Shell And Tube Heat Exchangers");

            
            OpenButton.Click();

        }


        [TestMethod]
        public void _013_EditDatasheet()
        {
            
            editEditField.Select(0, 3);

            editEditField.SetText("Field1177");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            
            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("1");

            
            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("1");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("2");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("2");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("3");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("3");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("4");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("4");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            editEditField.Select(0, 3);

            editEditField.SetText("Field1169");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("c");

            eXCEL6UiObject.SendKeys("m");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            editEditField.Select(0, 9);

            editEditField.SetText("Field1185");

            editEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("r");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("t");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("d");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("e");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

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


