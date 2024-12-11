using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.LFT.SDK;
using HP.LFT.SDK.UIAPro;


namespace ExcelDatasheetEditorUFTProject
{
    public static class UIAProUIObjects
    {

        #region General UI Objects (Apply to both types of datasheets)

        public static IWindow ExcelWindow = Desktop.Describe<IWindow>(new WindowDescription
        {
            ProcessName = @"EXCEL",
            SupportedPatterns = new string[] { @"LegacyIAccessible", @"Transform", @"Window" },
            FrameworkId = @"Win32",
            ControlType = @"Window",
            AutomationId = string.Empty
        });

        
        public static IPane MainPane = ExcelWindow.Describe<IPane>(new PaneDescription
        {
            ProcessName = @"EXCEL",
            Name = string.Empty,
            Path = @"Window;Pane",
            SupportedPatterns = new string[] { @"LegacyIAccessible" },
            FrameworkId = @"Win32",
            ControlType = @"Pane",
            AutomationId = string.Empty,
            Index = 3
        });

        


        public static IUiObject CustomPanel = MainPane
            .Describe<IUiObject>(new UiObjectDescription
            {
                ProcessName = @"EXCEL",
                Path = @"Window;Pane;Custom",
                SupportedPatterns = new string[] { @"LegacyIAccessible" },
                FrameworkId = @"Win32",
                ControlType = @"Custom",
                AutomationId = string.Empty
            });


        

        public static IPane NestedPane = CustomPanel
            .Describe<IPane>(new PaneDescription
            {
                ProcessName = @"EXCEL",
                Path = @"Window;Pane;Custom;Pane",
                SupportedPatterns = new string[] { @"LegacyIAccessible", @"Selection" },
                FrameworkId = string.Empty,
                ControlType = @"Pane",
            });

       

        #endregion


        // Entering values to datasheet

        #region  AZ Centrifugal Pump1 Page 2 


        public static IPane PanePage2 = NestedPane
            .Describe<IPane>(new PaneDescription
            {
                ProcessName = @"EXCEL",
                Name = @"Sheet Page 2",
                Path = @"Window;Pane;Custom;Pane;Pane",
                SupportedPatterns = new string[] { @"LegacyIAccessible", @"Selection", @"Selection2", @"SelectionItem" },
                FrameworkId = string.Empty,
                ControlType = @"Pane",
                AutomationId = @"Page 2"
            });


        public static IImage field260CImage = PanePage2
            .Describe<IImage>(new ImageDescription
            {
                ProcessName = @"EXCEL",
                Name = @"Field260-C",
                Path = @"Window;Pane;Custom;Pane;Pane;Image",
                SupportedPatterns = new string[] { @"Invoke", @"LegacyIAccessible", @"SelectionItem" },
                FrameworkId = string.Empty,
                ControlType = @"Image",
                AutomationId = string.Empty
            });

        public static IImage field264CImage = PanePage2
            .Describe<IImage>(new ImageDescription
            {
                ProcessName = @"EXCEL",
                Name = @"Field264-C",
                Path = @"Window;Pane;Custom;Pane;Pane;Image",
                SupportedPatterns = new string[] { @"Invoke", @"LegacyIAccessible", @"SelectionItem" },
                FrameworkId = string.Empty,
                ControlType = @"Image",
                AutomationId = string.Empty
            });

        #endregion
        #region AZ Centrifugal Pump1  Page 3

        public static IPane PanePage3 = NestedPane.Describe<IPane>(new PaneDescription
        {
            ProcessName = @"EXCEL",
            Name = @"Sheet Page 3",
            Path = @"Window;Pane;Custom;Pane;Pane",
            SupportedPatterns = new string[] { @"LegacyIAccessible", @"Selection", @"Selection2", @"SelectionItem" },
            FrameworkId = string.Empty,
            ControlType = @"Pane",
            AutomationId = @"Page 3"
        });

        public static IImage Field433CImage = PanePage3.Describe<IImage>(new ImageDescription
        {
            ProcessName = @"EXCEL",
            Name = @"Field433-C",
            Path = @"Window;Pane;Custom;Pane;Pane;Image",
            SupportedPatterns = new string[] { @"Invoke", @"LegacyIAccessible", @"SelectionItem" },
            FrameworkId = string.Empty,
            ControlType = @"Image",
            AutomationId = string.Empty
        });

        public static IImage Field434CImage = PanePage3.Describe<IImage>(new ImageDescription
        {
            ProcessName = @"EXCEL",
            Name = @"Field434-C",
            Path = @"Window;Pane;Custom;Pane;Pane;Image",
            SupportedPatterns = new string[] { @"Invoke", @"LegacyIAccessible", @"SelectionItem" },
            FrameworkId = string.Empty,
            ControlType = @"Image",
            AutomationId = string.Empty
        });

        #endregion
        #region AZ Centrifugal Pump1 Page 5 

        public static IPane PanePage5 = NestedPane.Describe<IPane>(new PaneDescription
        {
            ProcessName = @"EXCEL",
            Name = @"Sheet Page 5",
            Path = @"Window;Pane;Custom;Pane;Pane",
            SupportedPatterns = new string[] { @"LegacyIAccessible", @"Selection", @"Selection2", @"SelectionItem" },
            FrameworkId = string.Empty,
            ControlType = @"Pane",
            AutomationId = @"Page 5"
        });

        public static IImage Field669CImage = PanePage5.Describe<IImage>(new ImageDescription
        {
            ProcessName = @"EXCEL",
            Name = @"Field669-C",
            Path = @"Window;Pane;Custom;Pane;Pane;Image",
            SupportedPatterns = new string[] { @"Invoke", @"LegacyIAccessible", @"SelectionItem" },
            FrameworkId = string.Empty,
            ControlType = @"Image",
            AutomationId = string.Empty
        });

        #endregion


        // Verification
        #region AZ Centrifugal Pump1 Page 1 

        public static IPane PanePage1 = NestedPane
            .Describe<IPane>(new PaneDescription
            {
                ProcessName = @"EXCEL",
                Name = @"Sheet Page 1",
                Path = @"Window;Pane;Custom;Pane;Pane",
                SupportedPatterns = new string[] { @"LegacyIAccessible", @"Selection", @"Selection2", @"SelectionItem" },
                FrameworkId = string.Empty,
                ControlType = @"Pane",
                AutomationId = @"Page 1"
            });


        public static IGrid GridPage1 = PanePage1
                    .Describe<IGrid>(new GridDescription
                    {
                        ProcessName = @"EXCEL",
                        Name = @"Grid",
                        Path = @"Window;Pane;Custom;Pane;Pane;DataGrid",
                        SupportedPatterns = new string[] { @"Grid", @"LegacyIAccessible" },
                        FrameworkId = string.Empty,
                        ControlType = @"DataGrid",
                        AutomationId = @"Grid"
                    });


        public static IGridItem i23GridItem = GridPage1
                    .Describe<IGridItem>(new GridItemDescription
                    {
                        ProcessName = @"EXCEL",
                        Name = @"I23",
                        Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
                        SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
                        FrameworkId = string.Empty,
                        ControlType = @"DataItem",
                        AutomationId = @"I23"
                    });

        public static IGridItem k23GridItem = GridPage1
            .Describe<IGridItem>(new GridItemDescription
            {
                ProcessName = @"EXCEL",
                Name = @"K23",
                Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
                SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
                FrameworkId = string.Empty,
                ControlType = @"DataItem",
                AutomationId = @"K23"
            });


        #endregion

        #region AZ Centrifugal Pump1 Page 4 

        public static IPane PanePage4 = NestedPane
           .Describe<IPane>(new PaneDescription
           {
               ProcessName = @"EXCEL",
               Name = @"Sheet Page 4",
               Path = @"Window;Pane;Custom;Pane;Pane",
               SupportedPatterns = new string[] { @"LegacyIAccessible", @"Selection", @"Selection2", @"SelectionItem" },
               FrameworkId = string.Empty,
               ControlType = @"Pane",
               AutomationId = @"Page 4"
           });


        public static IGrid GridPage4 = PanePage4
            .Describe<IGrid>(new GridDescription
            {
                ProcessName = @"EXCEL",
                Name = @"Grid",
                Path = @"Window;Pane;Custom;Pane;Pane;DataGrid",
                SupportedPatterns = new string[] { @"Grid", @"LegacyIAccessible" },
                FrameworkId = string.Empty,
                ControlType = @"DataGrid",
                AutomationId = @"Grid"
            });

        public static IGridItem e16GridItem = GridPage4
            .Describe<IGridItem>(new GridItemDescription
            {
                ProcessName = @"EXCEL",
                Name = @"E16",
                Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
                SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
                FrameworkId = string.Empty,
                ControlType = @"DataItem",
                AutomationId = @"E16"
            });

        #endregion

        // Verification
        #region Heat Exchanger List Datasheet Page 1 


        public static IPane PaneSandTHTExs = NestedPane
             .Describe<IPane>(new PaneDescription
             {
                 ProcessName = @"EXCEL",
                 Name = @"Sheet S&T HT EXs",
                 Path = @"Window;Pane;Custom;Pane;Pane",
                 SupportedPatterns = new string[] { @"LegacyIAccessible", @"Selection", @"Selection2", @"SelectionItem" },
                 FrameworkId = string.Empty,
                 ControlType = @"Pane",
                 AutomationId = @"S&T HT EXs"
             });


        public static IGrid GridSandTHTExs = PaneSandTHTExs
             .Describe<IGrid>(new GridDescription
             {
                 ProcessName = @"EXCEL",
                 Name = @"Grid",
                 Path = @"Window;Pane;Custom;Pane;Pane;DataGrid",
                 SupportedPatterns = new string[] { @"Grid", @"LegacyIAccessible" },
                 FrameworkId = string.Empty,
                 ControlType = @"DataGrid",
                 AutomationId = @"Grid"
             });


        public static IGridItem d9GridItem = GridSandTHTExs
             .Describe<IGridItem>(new GridItemDescription
             {
                 ProcessName = @"EXCEL",
                 Name = @"D9",
                 Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
                 SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
                 FrameworkId = string.Empty,
                 ControlType = @"DataItem",
                 AutomationId = @"D9"
             });

		public static IGridItem d10GridItem = GridSandTHTExs
			.Describe<IGridItem>(new GridItemDescription
			{
				ProcessName = @"EXCEL",
				Name = @"D10",
				Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
				SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
				FrameworkId = string.Empty,
				ControlType = @"DataItem",
				AutomationId = @"D10"
			});

		public static IGridItem d11GridItem = GridSandTHTExs
					.Describe<IGridItem>(new GridItemDescription
					{
						ProcessName = @"EXCEL",
						Name = @"D11",
						Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
						SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
						FrameworkId = string.Empty,
						ControlType = @"DataItem",
						AutomationId = @"D11"
					});

		public static IGridItem d12GridItem = GridSandTHTExs
					.Describe<IGridItem>(new GridItemDescription
					{
						ProcessName = @"EXCEL",
						Name = @"D12",
						Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
						SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
						FrameworkId = string.Empty,
						ControlType = @"DataItem",
						AutomationId = @"D12"
					});

		public static IGridItem d13GridItem = GridSandTHTExs
					.Describe<IGridItem>(new GridItemDescription
					{
						ProcessName = @"EXCEL",
						Name = @"D13",
						Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
						SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
						FrameworkId = string.Empty,
						ControlType = @"DataItem",
						AutomationId = @"D13"
					});

		public static IGridItem d14GridItem = GridSandTHTExs
					.Describe<IGridItem>(new GridItemDescription
					{
						ProcessName = @"EXCEL",
						Name = @"D14",
						Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
						SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
						FrameworkId = string.Empty,
						ControlType = @"DataItem",
						AutomationId = @"D14"
					});

		public static IGridItem d15GridItem = GridSandTHTExs
                    .Describe<IGridItem>(new GridItemDescription
					{
						ProcessName = @"EXCEL",
						Name = @"D15",
						Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
						SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
						FrameworkId = string.Empty,
						ControlType = @"DataItem",
						AutomationId = @"D15"
					});

		public static IGridItem d16GridItem = GridSandTHTExs
                    .Describe<IGridItem>(new GridItemDescription
					{
						ProcessName = @"EXCEL",
						Name = @"D16",
						Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
						SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
						FrameworkId = string.Empty,
						ControlType = @"DataItem",
						AutomationId = @"D16"
					});

		public static IGridItem d17GridItem = GridSandTHTExs
                    .Describe<IGridItem>(new GridItemDescription
					{
						ProcessName = @"EXCEL",
						Name = @"D17",
						Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
						SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
						FrameworkId = string.Empty,
						ControlType = @"DataItem",
						AutomationId = @"D17"
					});

		public static IGridItem m10GridItem = GridSandTHTExs
			        .Describe<IGridItem>(new GridItemDescription
			{
				ProcessName = @"EXCEL",
				Name = @"M10",
				Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
				SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
				FrameworkId = string.Empty,
				ControlType = @"DataItem",
				AutomationId = @"M10"
			});

		public static IGridItem m12GridItem = GridSandTHTExs
                    .Describe<IGridItem>(new GridItemDescription
					{
						ProcessName = @"EXCEL",
						Name = @"M12",
						Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
						SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
						FrameworkId = string.Empty,
						ControlType = @"DataItem",
						AutomationId = @"M12"
					});

		public static IGridItem m14GridItem = GridSandTHTExs
                    .Describe<IGridItem>(new GridItemDescription
					{
						ProcessName = @"EXCEL",
						Name = @"M14",
						Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
						SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
						FrameworkId = string.Empty,
						ControlType = @"DataItem",
						AutomationId = @"M14"
					});

		public static IGridItem m16GridItem = GridSandTHTExs
                    .Describe<IGridItem>(new GridItemDescription
					{
						ProcessName = @"EXCEL",
						Name = @"M16",
						Path = @"Window;Pane;Custom;Pane;Pane;DataGrid;DataItem",
						SupportedPatterns = new string[] { @"GridItem", @"Invoke", @"LegacyIAccessible", @"SelectionItem", @"Text", @"Value" },
						FrameworkId = string.Empty,
						ControlType = @"DataItem",
						AutomationId = @"M16"
					});
		#endregion
	}
}
