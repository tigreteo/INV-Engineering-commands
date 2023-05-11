using Inventor;
using Microsoft.Win32;
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;



namespace MGBW_ENG_Commands
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [GuidAttribute("ed06946a-9e76-4747-9038-6439ffeeb2ee")]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {
        // Inventor application object.
        private Inventor.Application m_inventorApplication;

        public StandardAddInServer()
        {
        }

        #region ApplicationAddInServer Members

        // Declaration of the object for the UserInterfaceEvents to be able to handle
        // if the user resets the ribbon so the button can be added back in.
        private UserInterfaceEvents _m_uiEvents;

        public UserInterfaceEvents m_uiEvents
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _m_uiEvents;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_m_uiEvents != null)
                {
                    _m_uiEvents.OnResetRibbonInterface -= m_uiEvents_OnResetRibbonInterface;
                }

                _m_uiEvents = value;
                if (_m_uiEvents != null)
                {
                    _m_uiEvents.OnResetRibbonInterface += m_uiEvents_OnResetRibbonInterface;
                }
            }
        }

        #region setting up for buttons
        public class UI_Button
        {
            private ButtonDefinition _bd;

            public ButtonDefinition bd
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return this._bd;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (this._bd != null)
                    {
                        this._bd.OnExecute -= bd_OnExecute;
                    }

                    this._bd = value;
                    if (this._bd != null)
                    {
                        this._bd.OnExecute += bd_OnExecute;
                    }
                }
            }

            private void bd_OnExecute(NameValueMap Context)
            {
                // Link button clicks to their respective commands.
                switch (bd.InternalName)
                {
                    case "my_first_button":
                        CommandFunctions.RunAnExe();
                        return;
                    case "my_second_button":
                        CommandFunctions.PopupMessage();
                        return;
                    case "close_doc_button":
                        CommandFunctions.CloseDocument();
                        return;
                    case "export_dxf_button":
                        CommandFunctions.ExportDxf();
                        return;
                    case "Create_Overview_Drawing":
                        CommandFunctions.CreateOverViewDrawing();
                        return;
                    case "Create_Cutlist_Sheet":
                        CommandFunctions.CreateCutlistDrawing();
                        //CommandFunctions.PopupMessage();
                        return;
                    case "Prep_Files_for_Vendor":
                        CommandFunctions.PrepForVendors();
                        return;
                    case "Print_fitted_button":
                        CommandFunctions.printFitted();
                        return;
                    case "export_partslist":
                        CommandFunctions.exportPartsList();
                        return;
                    case "Prep_Files_for_Formal":
                        CommandFunctions.PrepForRelease();
                        return;
                    default:
                        return;
                }
            }
        }

        public delegate ButtonDefinition CreateButton(string display_text, string internal_name, string icon_path);
        public ButtonDefinition button_template(string display_text, string internal_name, string icon_path)
        {
            UI_Button MyButton = new UI_Button();
            MyButton.bd = Utilities.CreateButtonDefinition(display_text, internal_name, "", icon_path);
            return MyButton.bd;
        }

        // Declare all buttons here
        ButtonDefinition MyFirstButton;
        ButtonDefinition MySecondButton;
        ButtonDefinition CloseDocButton;
        ButtonDefinition ExportDxfButton;
        ButtonDefinition OverviewDrawingButton;
        ButtonDefinition CutlistButton;
        ButtonDefinition PrepToVendorButton;
        ButtonDefinition PrintFittedButton;
        ButtonDefinition ExportPartsList;
        ButtonDefinition PrepToReleaseButton;

        #endregion
        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            try
            { 
                // Initialize AddIn members.
                Globals.invApp = addInSiteObject.Application;
                m_inventorApplication = addInSiteObject.Application;

                // Connect to the user-interface events to handle a ribbon reset.
                m_uiEvents = Globals.invApp.UserInterfaceManager.UserInterfaceEvents;

                // This method is called by Inventor when it loads the addin.
                // The AddInSiteObject provides access to the Inventor Application object.
                // The FirstTime flag indicates if the addin is loaded for the first time.

                #region createbuttons
                // ButtonName = create_button(display_text, internal_name, icon_path)
                CreateButton create_button = new CreateButton(button_template);
                //MyFirstButton = create_button           ("    My First              \n    Command    ",     "my_first_button",          @"ButtonResources\MyIcon1");
                //MySecondButton = create_button          ("    My Second             \n    Command    ",     "my_second_button",         @"ButtonResources\MyIcon1");       
                CloseDocButton = create_button("    Close                 \n    Document    ", "close_doc_button", @"ButtonResources\MyIcon3");
                //ExportDxfButton = create_button         ("    Export                \n    DXF         ",    "export_dxf_button",        @"ButtonResources\MyIcon4");
                OverviewDrawingButton = create_button("Drawing Overview", "Create_Overview_Drawing", @"ButtonResources\MyIcon5");
                CutlistButton = create_button("    Create Cutlist        \n    Command    ", "Create_Cutlist_Sheet", @"ButtonResources\MyIcon6");
                PrepToVendorButton = create_button("    Prep for Vendor       \n    Command    ", "Prep_Files_for_Vendor", @"ButtonResources\MyIcon7");
                PrintFittedButton = create_button("Print Scaled to     \n    D1    ", "Print_fitted_button", @"ButtonResources\MyIcon8");
                ExportPartsList = create_button("Export PartsList \n    as Excel Doc ", "export_partslist", @"ButtonResources\MyIcon9");
                PrepToReleaseButton = create_button("Prep for Formal  \n    Command ", "Prep_Files_for_Formal", @"ButtonResources\MyIcon10");
                #endregion

                // TODO: Add ApplicationAddInServer.Activate implementation.
                // e.g. event initialization, command creation etc.
                if (firstTime)
                {
                    AddToUserInterface();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unexpected failure in the activation of the add-in \"Inventor_MGBW_Commans\"" + System.Environment.NewLine + System.Environment.NewLine + ex.Message + System.Environment.NewLine + ex.ToString());
            }
        }

        public void Deactivate()
        {
            // This method is called by Inventor when the AddIn is unloaded.
            // The AddIn will be unloaded either manually by the user or
            // when the Inventor session is terminated

            // TODO: Add ApplicationAddInServer.Deactivate implementation

            // Release objects.
            // Release objects.
            MyFirstButton = null;
            MySecondButton = null;
            CloseDocButton = null;
            ExportDxfButton = null;
            OverviewDrawingButton = null;
            CutlistButton = null;
            PrepToVendorButton = null;
            PrintFittedButton = null;
            ExportPartsList = null;
            PrepToReleaseButton = null;

            m_uiEvents = null;
            Globals.invApp = null;

            m_inventorApplication = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        // Note:this method is now obsolete, you should use the 
        // ControlDefinition functionality for implementing commands.
        public void ExecuteCommand(int commandID)
        {
            
        }

        public object Automation
        {
            // This property is provided to allow the AddIn to expose an API 
            // of its own to other programs. Typically, this  would be done by
            // implementing the AddIn's API interface in a class and returning 
            // that class object through this property.

            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }


        // Adds whatever is needed by this add-in to the user-interface.  This is 
        // called when the add-in loaded and also if the user interface is reset.
        private void AddToUserInterface()
        {

            // Get the ribbon. (more buttons can be added to various ribbons within this single addin)
            // Ribbons:
            // ZeroDoc
            // Part
            // Assembly
            // Drawing
            // Presentation
            // iFeatures
            // UnknownDocument
            Ribbon asmRibbon = Globals.invApp.UserInterfaceManager.Ribbons["Assembly"];
            Ribbon prtRibbon = Globals.invApp.UserInterfaceManager.Ribbons["Part"];
            Ribbon dwgRibbon = Globals.invApp.UserInterfaceManager.Ribbons["Drawing"];


            // Set up Tabs.
            // tab = setup_panel(display_name, internal_name, inv_ribbon)
            RibbonTab MyTab_asm;
            MyTab_asm = setup_tab("PD Commands", "pd_tab_asm", asmRibbon);

            RibbonTab MyTab_prt;
            MyTab_prt = setup_tab("PD Commands", "pd_tab_prt", prtRibbon);

            RibbonTab MyTab_dwg;
            MyTab_dwg = setup_tab("PD Commands", "pd_tab_dwg", dwgRibbon);


            // Set up Panels.
            // panel = setup_panel(display_name, internal_name, ribbon_tab)
            RibbonPanel MyPanel_prt;
            MyPanel_prt = setup_panel("My Panel", "my_panel_prt", MyTab_prt);

            RibbonPanel ExportPanel_prt;
            ExportPanel_prt = setup_panel("Export", "export_panel_prt", MyTab_prt);

            RibbonPanel MyPanel_dwg;
            MyPanel_dwg = setup_panel("My Panel", "my_panel_dwg", MyTab_dwg);

            RibbonPanel MyPanel_asm;
            MyPanel_asm = setup_panel("My Panel", "my_panel_asm", MyTab_asm);

            //MessageBox.Show("all good before we set up buttons");

            // Set up Buttons.
            if (!(MyFirstButton == null))
            {
                MyPanel_asm.CommandControls.AddButton(MyFirstButton, true);
                MyPanel_dwg.CommandControls.AddButton(MyFirstButton, true);
            }

            if (!(MySecondButton == null))
            {
                MyPanel_prt.CommandControls.AddButton(MySecondButton, true);
                MyPanel_dwg.CommandControls.AddButton(MySecondButton, true);
            }

            if (!(CloseDocButton == null))
            {
                MyPanel_prt.CommandControls.AddButton(CloseDocButton, true);
                MyPanel_dwg.CommandControls.AddButton(CloseDocButton, true);
            }

            if (!(OverviewDrawingButton == null))
            {
                MyPanel_prt.CommandControls.AddButton(OverviewDrawingButton, true);
                MyPanel_asm.CommandControls.AddButton(OverviewDrawingButton, true);
            }

            if (!(CutlistButton == null))
            {
                MyPanel_dwg.CommandControls.AddButton(CutlistButton, true);
            }

            if (!(PrepToVendorButton == null))
            {
                //MyPanel_prt.CommandControls.AddButton(PrepToVendorButton, true);
                //MyPanel_asm.CommandControls.AddButton(PrepToVendorButton, true);
                MyPanel_dwg.CommandControls.AddButton(PrepToVendorButton, true);
            }

            if (!(ExportDxfButton == null))
            {
                ExportPanel_prt.CommandControls.AddButton(ExportDxfButton, true);
            }

            if (!(PrintFittedButton == null))
            {
                MyPanel_dwg.CommandControls.AddButton(PrintFittedButton, true);
            }

            if (!(ExportPartsList == null))
            {
                MyPanel_dwg.CommandControls.AddButton(ExportPartsList, true);
            }

            if(!(PrepToReleaseButton == null))
            {
                MyPanel_dwg.CommandControls.AddButton(PrepToReleaseButton, true);
            }
        }


        private RibbonTab setup_tab(string display_name, string internal_name, Ribbon inv_ribbon)
        {
            RibbonTab setup_tabRet = default(RibbonTab);
            RibbonTab ribbon_tab = null;
            try
            {
                ribbon_tab = inv_ribbon.RibbonTabs[internal_name];
            }
            catch (Exception ex)
            {
            }

            if (ribbon_tab == null)
            {
                ribbon_tab = inv_ribbon.RibbonTabs.Add(display_name, internal_name, Globals.g_addInClientID);
            }

            setup_tabRet = ribbon_tab;
            return setup_tabRet;
        }


        private RibbonPanel setup_panel(string display_name, string internal_name, RibbonTab ribbon_tab)
        {
            RibbonPanel setup_panelRet = default(RibbonPanel);
            RibbonPanel ribbon_panel = null;
            try
            {
                ribbon_panel = ribbon_tab.RibbonPanels[internal_name];
            }
            catch (Exception ex)
            {
            }

            if (ribbon_panel == null)
            {
                ribbon_panel = ribbon_tab.RibbonPanels.Add(display_name, internal_name, Globals.g_addInClientID);
            }

            setup_panelRet = ribbon_panel;
            return setup_panelRet;
        }


        private void m_uiEvents_OnResetRibbonInterface(NameValueMap Context)
        {
            // The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface();
        }
        #endregion

    }

    public static class Globals
    {
        // Inventor application object.
        public static Inventor.Application invApp;

        // The unique ID for this add-in.  If this add-in is copied to create a new add-in
        // you need to update this ID along with the ID in the .manifest file, the .addin file
        // and create a new ID for the typelib GUID in AssemblyInfo.vb
        public const string g_simpleAddInClientID = "ed06946a-9e76-4747-9038-6439ffeeb2ee";
        public const string g_addInClientID = "{" + g_simpleAddInClientID + "}";
    }
}
