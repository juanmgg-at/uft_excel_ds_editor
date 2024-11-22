using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.LFT.SDK;
using HP.LFT.SDK.UIAPro;


namespace UftDeveloperTestProject1
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
        
        public static IPane PanePage3= NestedPane.Describe<IPane>(new PaneDescription
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

        
        #region Page 5 
        
        public static IPane PanePage5= NestedPane.Describe<IPane>(new PaneDescription
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
    }
}
