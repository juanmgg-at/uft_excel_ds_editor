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

        #region General UI Objects 

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

        #region Page 2


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


        #region Page 3

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


        #region Page 4 

        public static IGridItem e16GridItem = Desktop.Describe<IWindow>(new WindowDescription
        {
            ProcessName = @"EXCEL",
           Path = @"Window",
            SupportedPatterns = new string[] { @"LegacyIAccessible", @"Transform", @"Window" },
            FrameworkId = @"Win32",
            ControlType = @"Window",
            AutomationId = string.Empty
        })
            .Describe<IPane>(new PaneDescription
            {
                ProcessName = @"EXCEL",
                Name = string.Empty,
                Path = @"Window;Pane",
                SupportedPatterns = new string[] { @"LegacyIAccessible" },
                FrameworkId = @"Win32",
                ControlType = @"Pane",
                AutomationId = string.Empty,
                Index = 3
            })
            .Describe<IUiObject>(new UiObjectDescription
            {
                ProcessName = @"EXCEL",
                Path = @"Window;Pane;Custom",
                SupportedPatterns = new string[] { @"LegacyIAccessible" },
                FrameworkId = @"Win32",
                ControlType = @"Custom",
                AutomationId = string.Empty
            })
            .Describe<IPane>(new PaneDescription
            {
                ProcessName = @"EXCEL",
                Path = @"Window;Pane;Custom;Pane",
                SupportedPatterns = new string[] { @"LegacyIAccessible", @"Selection" },
                FrameworkId = string.Empty,
                ControlType = @"Pane",
            })
            .Describe<IPane>(new PaneDescription
            {
                ProcessName = @"EXCEL",
                Name = @"Sheet Page 4",
                Path = @"Window;Pane;Custom;Pane;Pane",
                SupportedPatterns = new string[] { @"LegacyIAccessible", @"Selection", @"Selection2", @"SelectionItem" },
                FrameworkId = string.Empty,
                ControlType = @"Pane",
                AutomationId = @"Page 4"
            })
            .Describe<IGrid>(new GridDescription
            {
                ProcessName = @"EXCEL",
                Name = @"Grid",
                Path = @"Window;Pane;Custom;Pane;Pane;DataGrid",
                SupportedPatterns = new string[] { @"Grid", @"LegacyIAccessible" },
                FrameworkId = string.Empty,
                ControlType = @"DataGrid",
                AutomationId = @"Grid"
            })
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

        #region Page 5 

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
        #region Page 1 
        public static IGridItem i23GridItem = Desktop.Describe<IWindow>(new WindowDescription
        {
            ProcessName = @"EXCEL",
            
            Path = @"Window",
            SupportedPatterns = new string[] { @"LegacyIAccessible", @"Transform", @"Window" },
            FrameworkId = @"Win32",
            ControlType = @"Window",
            AutomationId = string.Empty
        })
                    .Describe<IPane>(new PaneDescription
                    {
                        ProcessName = @"EXCEL",
                        Name = string.Empty,
                        Path = @"Window;Pane",
                        SupportedPatterns = new string[] { @"LegacyIAccessible" },
                        FrameworkId = @"Win32",
                        ControlType = @"Pane",
                        AutomationId = string.Empty,
                        Index = 3
                    })
                    .Describe<IUiObject>(new UiObjectDescription
                    {
                        ProcessName = @"EXCEL",
                        Path = @"Window;Pane;Custom",
                        SupportedPatterns = new string[] { @"LegacyIAccessible" },
                        FrameworkId = @"Win32",
                        ControlType = @"Custom",
                        AutomationId = string.Empty
                    })
                    .Describe<IPane>(new PaneDescription
                    {
                        ProcessName = @"EXCEL",
                        Path = @"Window;Pane;Custom;Pane",
                        SupportedPatterns = new string[] { @"LegacyIAccessible", @"Selection" },
                        FrameworkId = string.Empty,
                        ControlType = @"Pane",
                    })
                    .Describe<IPane>(new PaneDescription
                    {
                        ProcessName = @"EXCEL",
                        Name = @"Sheet Page 1",
                        Path = @"Window;Pane;Custom;Pane;Pane",
                        SupportedPatterns = new string[] { @"LegacyIAccessible", @"Selection", @"Selection2", @"SelectionItem" },
                        FrameworkId = string.Empty,
                        ControlType = @"Pane",
                        AutomationId = @"Page 1"
                    })
                    .Describe<IGrid>(new GridDescription
                    {
                        ProcessName = @"EXCEL",
                        Name = @"Grid",
                        Path = @"Window;Pane;Custom;Pane;Pane;DataGrid",
                        SupportedPatterns = new string[] { @"Grid", @"LegacyIAccessible" },
                        FrameworkId = string.Empty,
                        ControlType = @"DataGrid",
                        AutomationId = @"Grid"
                    })
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

        public static IGridItem k23GridItem = Desktop.Describe<IWindow>(new WindowDescription
        {
            ProcessName = @"EXCEL",
            Path = @"Window",
            SupportedPatterns = new string[] { @"LegacyIAccessible", @"Transform", @"Window" },
            FrameworkId = @"Win32",
            ControlType = @"Window",
            AutomationId = string.Empty
        })
            .Describe<IPane>(new PaneDescription
            {
                ProcessName = @"EXCEL",
                Name = string.Empty,
                Path = @"Window;Pane",
                SupportedPatterns = new string[] { @"LegacyIAccessible" },
                FrameworkId = @"Win32",
                ControlType = @"Pane",
                AutomationId = string.Empty,
                Index = 3
            })
            .Describe<IUiObject>(new UiObjectDescription
            {
                ProcessName = @"EXCEL",
                Path = @"Window;Pane;Custom",
                SupportedPatterns = new string[] { @"LegacyIAccessible" },
                FrameworkId = @"Win32",
                ControlType = @"Custom",
                AutomationId = string.Empty
            })
            .Describe<IPane>(new PaneDescription
            {
                ProcessName = @"EXCEL",
                Path = @"Window;Pane;Custom;Pane",
                SupportedPatterns = new string[] { @"LegacyIAccessible", @"Selection" },
                FrameworkId = string.Empty,
                ControlType = @"Pane",
            })
            .Describe<IPane>(new PaneDescription
            {
                ProcessName = @"EXCEL",
                Name = @"Sheet Page 1",
                Path = @"Window;Pane;Custom;Pane;Pane",
                SupportedPatterns = new string[] { @"LegacyIAccessible", @"Selection", @"Selection2", @"SelectionItem" },
                FrameworkId = string.Empty,
                ControlType = @"Pane",
                AutomationId = @"Page 1"
            })
            .Describe<IGrid>(new GridDescription
            {
                ProcessName = @"EXCEL",
                Name = @"Grid",
                Path = @"Window;Pane;Custom;Pane;Pane;DataGrid",
                SupportedPatterns = new string[] { @"Grid", @"LegacyIAccessible" },
                FrameworkId = string.Empty,
                ControlType = @"DataGrid",
                AutomationId = @"Grid"
            })
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

        #region Page 2 



        #endregion

    }
}
