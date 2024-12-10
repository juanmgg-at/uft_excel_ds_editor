using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HP.LFT.SDK;
using HP.LFT.SDK.StdWin;
using System.Net;

//Libraries of objects
using static ExcelDatasheetEditorUFTProject.UIAProUIObjects;
using static ExcelDatasheetEditorUFTProject.StdWinUIObjects;

using System.Threading;

namespace ExcelDatasheetEditorUFTProject
{
    [TestClass]
    public class UftDeveloperTest2 : UnitTestClassBase<UftDeveloperTest2>
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

        #region Continuous List Datasheet

        [TestMethod]
        public void _011_CreateShellAndTubeHeatExchangers()
        {

            Thread.Sleep(1000);
            //AZCentrifugalPump1DesignDatasheetXlsmUserNameCWindow.Close();

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

            Thread.Sleep(4000);

            //editEditField.Select(0, 8);

            var comboBox = DesignDatasheetXlsmUserNameCWindow
            .Describe<IComboBox>(new ComboBoxDescription
            {
                NativeClass = @"ComboBox"
            });

            comboBox.Click();

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
