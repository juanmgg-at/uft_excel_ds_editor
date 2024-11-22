using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HP.LFT.SDK;
using HP.LFT.SDK.StdWin;
using HP.LFT.Verifications;
using System.Net;

//Librariews of objects
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
            
           
            OpenAspenBasicEngineeringV15WorkspaceDialog.Activate();

            
            cONNECTIONButton.Click();

            
            BrowseButton.Click();

            

            SysTreeView32TreeViewOpenWorkspace.Click();

            //sysTreeView32TreeView.Select("Aspen Basic Engineering Workspaces on (gutierrj-S22.qae.aspentech.com);A1 on gutierrj-S22.qae.aspentech.com");

            //gutierrj-S22.qae.aspentech.com

            SysTreeView32TreeViewOpenWorkspace.Select($"Aspen Basic Engineering Workspaces on ({Dns.GetHostEntry(Environment.MachineName).HostName});{EnvironmentSetting.workspace} on {Dns.GetHostEntry(Environment.MachineName).HostName}");

            
            OKButtonOpenWorkspace.Click();

            
            OpenButtonOpenWorkspace.Click();

        }

        [TestMethod]
        public void _0002_CreateCentrifugalPump()
        {
            
            AspenBasicEngineeringV15Window.Activate();


            SysTreeView32TreeView.Select("\\\\gutierrj-S22.qae.aspentech.com\\A1");

            SysTreeView32TreeView.SendKeys("b", HP.LFT.SDK.KeyModifier.Alt);


            MenuBarUiObject.SendKeys("c");

            
            SelectClassDialogCreatePump.Activate();

            
            ListBoxCreatePump.Click();

            ListBoxCreatePump.Select("Centrifugal Pump");

            
            OKButtonCreatePump.Click();

        }


        [TestMethod]
        public void _001_CreateDatasheetsFolder()
        {

            Thread.Sleep(1000);

            
            AspenBasicEngineeringV15Window.Activate();

            
            SysTreeView32TreeView.Select($@"\\{Dns.GetHostEntry(Environment.MachineName).HostName}\{EnvironmentSetting.workspace}");

            SysTreeView32TreeView.SendKeys("o", HP.LFT.SDK.KeyModifier.Alt);

           
            MenuBarUiObject.SendKeys("nf");

           
            EditEditField.SetText("Datasheets");

            EditEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            //SysTreeView32TreeView.Select($@"\\{Dns.GetHostEntry(Environment.MachineName).HostName}\{EnvironmentSetting.workspace};Datasheets");


        }


        [TestMethod]
        public void _002_CreateDatasheet()
        {
            Thread.Sleep(1000);

            AspenBasicEngineeringV15Window.Activate();

            
            SysTreeView32TreeView.Select($@"\\{Dns.GetHostEntry(Environment.MachineName).HostName}\{EnvironmentSetting.workspace}");

            SysTreeView32TreeView.SendKeys("b", HP.LFT.SDK.KeyModifier.Alt);

            
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


            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("m");

            eXCEL6UiObject.SendKeys(HP.LFT.SDK.Keys.Down);
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


            //Select  method never works 

            SysListView32ListViewOpenDatasheet.Click();

            SysListView32ListViewOpenDatasheet.Select(1);

            SysListView32ListViewOpenDatasheet.SendKeys(HP.LFT.SDK.Keys.Return);




            //Thread.Sleep(1000);

            //AspenBasicEngineeringV15Window.Activate();

            //SysTreeView32TreeViewOpenDatasheet.Select("\\\\gutierrj-S22.qae.aspentech.com\\A1;All Datasheets");

            //Thread.Sleep(1000);


            //SysListView32ListViewOpenDatasheet.Click();
            //for (int index = 0; index < SysListView32ListViewOpenDatasheet.Items.Count; index++)
            //{
            //    SysListView32ListViewOpenDatasheet.Select(1);
            //}




        }




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


