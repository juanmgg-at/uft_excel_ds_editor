using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HP.LFT.SDK;
using HP.LFT.SDK.StdWin;

namespace UftDeveloperTestProject1
{
    public static class StdWinUIObjects
    {

        #region ABE UI elements

        //General UI Elements 
        public static IWindow AspenBasicEngineeringV15Window = Desktop.Describe<IWindow>(new WindowDescription
        {
            IsChildWindow = false,
            IsOwnedWindow = false,
            WindowClassRegExp = @"Afx:",
            WindowTitleRegExp = @"Aspen Basic Engineering " + EnvironmentSetting.version
        });

        public static ITreeView SysTreeView32TreeView = AspenBasicEngineeringV15Window.Describe<ITreeView>(new TreeViewDescription
        {
            NativeClass = @"SysTreeView32",
            WindowId = 2,
            Index = 1
        });

        public static IUiObject MenuBarUiObject = AspenBasicEngineeringV15Window.Describe<IUiObject>(new UiObjectDescription
        {
            WindowClassRegExp = @"Afx:",
            WindowId = 59398
        });

        

        public static IEditField EditEditField = SysTreeView32TreeView.Describe<IEditField>(new EditFieldDescription
        {
            NativeClass = @"Edit"
        });

        // Open Workspace 

        public static IDialog OpenAspenBasicEngineeringV15WorkspaceDialog = AspenBasicEngineeringV15Window.Describe<IDialog>(new DialogDescription
        {
            IsChildWindow = false,
            IsOwnedWindow = true,
            NativeClass = @"#32770",
            Text = @"Open Aspen Basic Engineering " + EnvironmentSetting.version + " Workspace"
        });

        public static IButton cONNECTIONButton = OpenAspenBasicEngineeringV15WorkspaceDialog.Describe<IButton>(new ButtonDescription
            {
                NativeClass = @"Button",
                Text = @"CONNECTION",
                WindowId = 337
            });

        public static IButton BrowseButton = OpenAspenBasicEngineeringV15WorkspaceDialog.Describe<IButton>(new ButtonDescription
        {
            NativeClass = @"Button",
            Text = @"&Browse..."
        });

        public static IDialog SelectWorkspaceOrProjectDialog = Desktop.Describe<IDialog>(new DialogDescription
        {
            IsChildWindow = false,
            IsOwnedWindow = true,
            NativeClass = @"#32770",
            Text = @"Select Workspace or Project"
        });
        public static ITreeView SysTreeView32TreeViewOpenWorkspace = SelectWorkspaceOrProjectDialog.Describe<ITreeView>(new TreeViewDescription
        {
            NativeClass = @"SysTreeView32"
        });

        public static IButton OKButtonOpenWorkspace = SelectWorkspaceOrProjectDialog.Describe<IButton>(new ButtonDescription
        {
            NativeClass = @"Button",
            Text = @"OK"
        });

        public static IButton OpenButtonOpenWorkspace = OpenAspenBasicEngineeringV15WorkspaceDialog.Describe<IButton>(new ButtonDescription
            {
                NativeClass = @"Button",
                Text = @"&Open"
            });

        // Create Centrifugal Pump 
        public static IDialog SelectClassDialogCreateEquipment = AspenBasicEngineeringV15Window.Describe<IDialog>(new DialogDescription
        {
            IsChildWindow = false,
            IsOwnedWindow = true,
            NativeClass = @"#32770",
            Text = @"Select Class"
        });

        public static IListBox ListBoxCreateEquipment = SelectClassDialogCreateEquipment.Describe<IListBox>(new ListBoxDescription
        {
            NativeClass = @"ListBox"
        });

        public static IButton OKButtonCreateEquipment 
            = SelectClassDialogCreateEquipment.Describe<IButton>(new ButtonDescription
            {
                NativeClass = @"Button",
                Text = @"OK"
            });

        //Create datasheet


        public static IDialog SelectDatasheetByTypeDialog = AspenBasicEngineeringV15Window.Describe<IDialog>(new DialogDescription
        {
            IsChildWindow = false,
            IsOwnedWindow = true,
            NativeClass = @"#32770",
            Text = @"Select datasheet by type"
        });

        public static IListBox ListBox = SelectDatasheetByTypeDialog.Describe<IListBox>(new ListBoxDescription
        {
            NativeClass = @"ListBox"
        });

        public static IDialog DatasheetDialog = AspenBasicEngineeringV15Window.Describe<IDialog>(new DialogDescription
        {
            IsChildWindow = false,
            IsOwnedWindow = true,
            NativeClass = @"#32770",
            Text = @"Datasheet"
        });

        public static ITabControl SysTabControl32TabControl = DatasheetDialog.Describe<ITabControl>(new TabControlDescription
        {
            NativeClass = @"SysTabControl32"
        });

        public static IListView SysListView32ListView = DatasheetDialog.Describe<IListView>(new ListViewDescription
        {
            NativeClass = @"SysListView32"
        });

        public static IButton OpenButton = DatasheetDialog.Describe<IButton>(new ButtonDescription
        {
            NativeClass = @"Button",
            Text = @"&Open"
        });


        //Add Datsheet to folder

        public static IDialog SelectDatasheetDialog = AspenBasicEngineeringV15Window.Describe<IDialog>(new DialogDescription
        {
            IsChildWindow = false,
            IsOwnedWindow = true,
            NativeClass = @"#32770",
            Text = @"Select Datasheet"
        });

        public static IListBox ListBoxAddDatasheet = SelectDatasheetDialog.Describe<IListBox>(new ListBoxDescription
        {
            NativeClass = @"ListBox"
        });

        //Open Datasheet 

        public static ITreeView SysTreeView32TreeViewOpenDatasheet
            = AspenBasicEngineeringV15Window.Describe<ITreeView>(new TreeViewDescription
        {
            NativeClass = @"SysTreeView32",
            WindowId = 2,
            Index = 0
        });

        public static IWindow FilterViewAllDatasheetsWindow = AspenBasicEngineeringV15Window.Describe<IWindow>(new WindowDescription
        {
            IsChildWindow = true,
            IsOwnedWindow = false,
            WindowClassRegExp = @"Afx:",
            WindowTitleRegExp = @"Filter View: AllDatasheets"
        });
        public static IListView SysListView32ListViewOpenDatasheet = FilterViewAllDatasheetsWindow.Describe<IListView>(new ListViewDescription
        {
            NativeClass = @"SysListView32"
        });


        #endregion


        #region  Excel UI elements

        public static IWindow DesignDatasheetXlsmUserNameCWindow = Desktop.Describe<IWindow>(new WindowDescription
        {
            IsChildWindow = false,
            IsOwnedWindow = false,
            WindowClassRegExp = @"XLMAIN",

        });

        public static IEditField editEditField = DesignDatasheetXlsmUserNameCWindow.Describe<IEditField>(new EditFieldDescription
        {
            NativeClass = @"Edit"
        });


        public static IUiObject AZCentrifugalPump1DesignDatasheetXlsmUiObject = DesignDatasheetXlsmUserNameCWindow.Describe<IUiObject>(new UiObjectDescription
        {
            WindowClassRegExp = @"EXCEL7"
        });


        public static IUiObject eXCEL6UiObject = DesignDatasheetXlsmUserNameCWindow.Describe<IUiObject>(new UiObjectDescription
        {
            WindowClassRegExp = @"EXCEL6"
        });

        #endregion
    }
}
