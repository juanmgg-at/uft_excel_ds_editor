using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HP.LFT.SDK;
using HP.LFT.SDK.StdWin;
using System.Net;

//Libraries of objects
using static ExcelDatasheetEditorUFTProject.UIAProUIObjects;
using static ExcelDatasheetEditorUFTProject.StdWinUIObjects;


// Library with constants 
using static ExcelDatasheetEditorUFTProject.Constants;

using System.Threading;
using HP.LFT.Verifications;

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
        public void _012_CreateShellAndTubeHeatExchangers()
        {

            Thread.Sleep(1000);

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
        public void _013_CreateATContinuousListShellAndTubeExchangersDatasheet()
        {


            Thread.Sleep(1000);

            AspenBasicEngineeringV15Window.Activate();


            SysTreeView32TreeView.Select($@"\\{Dns.GetHostEntry(Environment.MachineName).HostName}\{EnvironmentSetting.workspace}");

            Thread.Sleep(1000);

            SysTreeView32TreeView.SendKeys("b", HP.LFT.SDK.KeyModifier.Alt);


            Thread.Sleep(1000);

            for (int i = 0; i < 5; i++)
            {
                MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Down);
            }

            MenuBarUiObject.SendKeys(HP.LFT.SDK.Keys.Return);


            SelectDatasheetByTypeDialog.Activate();


            SelectDatasheetTypeListBox.Click();

            SelectDatasheetTypeListBox.Select("AT Continuous List Shell And Tube Heat Exchangers");


            //var OKButton = SelectDatasheetByTypeDialog.Describe<IButton>(new ButtonDescription
            //{
            //    NativeClass = @"Button",
            //    Text = @"OK"
            //});

            SelectDatasheetTypeOKButton.Click();


            DatasheetDialog.Activate();


            SysTabControl32TabControl.Select("New");



            SysListView32ListView.Click();

            SysListView32ListView.Select("AT Continuous List Shell And Tube Heat Exchangers");


            CreateDatasheetOpenButton.Click();

        }


        [TestMethod]
        public void _014_EditCreateATContinuousListShellAndTubeExchangersDatasheet()
        {

            Thread.Sleep(4000);

            Reporter.StartReportingContext("Shell Diameter Values");

            ExcelEditEditField.Select(0, 9);

            ExcelEditEditField.SetText("Field1177");

            ExcelEditEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("1");


            Excel6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("1");

            Excel6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("2");

            Excel6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("2");

            Excel6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("3");

            Excel6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("3");

            Excel6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("4");

            Excel6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("4");

            Excel6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            Reporter.EndReportingContext();


            Thread.Sleep(1000);

            Reporter.StartReportingContext("Shell Diameter Units");


            ExcelEditEditField.Select(0, 3);

            ExcelEditEditField.SetText("Field1169");

            ExcelEditEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys(ExpectedShellDiamUnits);

            //Excel6UiObject.SendKeys("m");

            Excel6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            Reporter.EndReportingContext();


            Thread.Sleep(1000);

            Reporter.StartReportingContext("Heat Exchanger Status");

            ExcelEditEditField.Select(0, 9);

            ExcelEditEditField.SetText("Field1185");

            ExcelEditEditField.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("r");

            Excel6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("t");

            Excel6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("d");

            Excel6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            AZCentrifugalPump1DesignDatasheetXlsmUiObject.SendKeys("e");

            Excel6UiObject.SendKeys(HP.LFT.SDK.Keys.Return);

            Reporter.EndReportingContext();

        }

        [TestMethod]
        public void _015_VerifyCreateATContinuousListShellAndTubeExchangersDatasheetValues()
        {

            Reporter.StartReportingContext("Shell Diameter Values");

            string shellDiamD10 = d10GridItem.TextPattern.VisibleText;
            Verify.AreEqual(shellDiamD10, ExpectedShellDiamD10, $"Shell Diameter Value is {ExpectedShellDiamD10}");

            string shellDiamD11 = d11GridItem.TextPattern.VisibleText;
            Verify.AreEqual(shellDiamD11, ExpectedShellDiamD11, $"Shell Diameter Value is {ExpectedShellDiamD11}");

            string shellDiamD12 = d12GridItem.TextPattern.VisibleText;
            Verify.AreEqual(shellDiamD12, ExpectedShellDiamD12, $"Shell Diameter Value is {ExpectedShellDiamD12}");

            string shellDiamD13 = d13GridItem.TextPattern.VisibleText;
            Verify.AreEqual(shellDiamD13, ExpectedShellDiamD13, $"Shell Diameter Value is {ExpectedShellDiamD13}");

            string shellDiamD14 = d14GridItem.TextPattern.VisibleText;
            Verify.AreEqual(shellDiamD14, ExpectedShellDiamD14, $"Shell Diameter Value is {ExpectedShellDiamD14}");

            string shellDiamD15 = d15GridItem.TextPattern.VisibleText;
            Verify.AreEqual(shellDiamD15, ExpectedShellDiamD15, $"Shell Diameter Value is {ExpectedShellDiamD15}");

            string shellDiamD16 = d16GridItem.TextPattern.VisibleText;
            Verify.AreEqual(shellDiamD16, ExpectedShellDiamD16, $"Shell Diameter Value is {ExpectedShellDiamD16}");

            string shellDiamD17 = d17GridItem.TextPattern.VisibleText;
            Verify.AreEqual(shellDiamD17, ExpectedShellDiamD17, $"Shell Diameter Value is {ExpectedShellDiamD17}");


            Reporter.EndReportingContext();

            Reporter.StartReportingContext("Shell Diameter Units");
            string shallDiamUnits = d9GridItem.TextPattern.VisibleText;

            Verify.AreEqual(shallDiamUnits, ExpectedShellDiamUnits, $"Shall Diameter Units are {ExpectedShellDiamUnits}");

            Reporter.EndReportingContext();

            Reporter.StartReportingContext("Heat Exchanger Status");

            string exchangerStatus1 = m10GridItem.TextPattern.VisibleText;
            Verify.AreEqual(exchangerStatus1, ExpectedExchangerStatus1, $"Exchanger Status is {ExpectedExchangerStatus1}");

            string exchangerStatus2 = m12GridItem.TextPattern.VisibleText;
            Verify.AreEqual(exchangerStatus2, ExpectedExchangerStatus2, $"Exchanger Status is {ExpectedExchangerStatus2}");

           string exchangerStatus3 = m14GridItem.TextPattern.VisibleText;
            Verify.AreEqual(exchangerStatus3, ExpectedExchangerStatus3, $"Exchanger Status is {ExpectedExchangerStatus3}");

           string exchangerStatus4 = m16GridItem.TextPattern.VisibleText;
            Verify.AreEqual(exchangerStatus4, ExpectedExchangerStatus4, $"Exchanger Status is {ExpectedExchangerStatus4}");


            DesignDatasheetXlsmUserNameCWindow.Close();
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
