//Module exists for a location on the code that does the acutal work, not just the interaction with the user/interop with inventor

using Inventor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Compression;

namespace MGBW_ENG_Commands
{
    public static class CommandFunctions
    {

        public struct Fraction
        {
            public Fraction(int n, int d)
            {
                N = n;
                D = d;
            }

            public int N { get; private set; }
            public int D { get; private set; }
        }

        public static void RunAnExe()
        {
            Process proc = new Process();
            proc = Process.Start(@"C:\path_to\some_file.exe", "");
        }

        public static void PopupMessage()
        {
            MessageBox.Show("This is a message box for PD commands");
        }

        public static void CloseDocument()
        {
            Globals.invApp.ActiveDocument.Close(true);
        }

        public static void ExportDxf()
        {
            PartDocument doc = (PartDocument)Globals.invApp.ActiveDocument;

            if (doc.DocumentSubType.DocumentSubTypeID != "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                MessageBox.Show("This can only be run on a Sheet Metal document. Exiting...");
                return;
            }

            string DXF_PATH;
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Select DXF Folder";
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                DXF_PATH = fbd.SelectedPath;
            }
            else
            {
                MessageBox.Show("Must specify DXF path");
                return;
            }

            SheetMetalComponentDefinition smcd;
            smcd = (SheetMetalComponentDefinition)doc.ComponentDefinition;

            if (!smcd.HasFlatPattern)
            {
                try
                {
                    smcd.Unfold();
                    smcd.FlatPattern.ExitEdit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not unfold part. Create flat pattern and try again.\n\n" + ex.Message);
                }
            }

            const string DXF_OPTIONS = "FLAT PATTERN DXF?AcadVersion=2004" +
                                            "&OuterProfileLayer=IV_INTERIOR_PROFILES" +
                                            "&InvisibleLayers=" +
                                                "IV_TANGENT;" +
                                                "IV_FEATURE_PROFILES_DOWN;" +
                                                "IV_BEND;" +
                                                "IV_BEND_DOWN;" +
                                                "IV_TOOL_CENTER;" +
                                                "IV_TOOL_CENTER_DOWN;" +
                                                "IV_ARC_CENTERS;" +
                                                "IV_FEATURE_PROFILES;" +
                                                "IV_FEATURE_PROFILES_DOWN;" +
                                                "IV_ALTREP_FRONT;" +
                                                "IV_ALTREP_BACK;" +
                                                "IV_ROLL_TANGENT;" +
                                                "IV_ROLL" +
                                            "&SimplifySplines=True" +
                                            "&BendLayerColor=255;255;0";

            string DISPLAY_NAME = doc.DisplayName;
            string dxf_filename = string.Format("{0}\\{1}.dxf", DXF_PATH, DISPLAY_NAME);
            try
            {
                smcd.DataIO.WriteDataToFile(DXF_OPTIONS, dxf_filename);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save DXF");
                return;
            }

            MessageBox.Show("DXF saved to path:\n\n" + DXF_PATH);
        }

        /// <summary>
        /// Return appropriate file location in released folder
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>folder location directory </returns>
        private static string ReleasedFolderLocation(string fileName)
        {
            //get styleID
            //try to split file name by "-"
            string[] nameParts = System.IO.Path.GetFileNameWithoutExtension(fileName).Split('-');
            string prefix = nameParts[0];
            double result = 0;

            //assumes destination folder structure of frame folder/sub folders by 100 increments
            #region Find/Create Group Folder
            //if the first string from the split can be converted to a number
            try
            {
                result = Convert.ToDouble(prefix.Trim());
            }
            catch (FormatException)
            {
                Console.WriteLine("Unable to convert '{0}' to a Double.", fileName);
                //continue;
            }
            catch (OverflowException)
            {
                Console.WriteLine("'{0}' is outside the range of a Double.", fileName);
                //continue;
            }

            decimal folderNum = Math.Truncate(Convert.ToDecimal(result / 100));
            folderNum = folderNum * 100;
            string folderName = folderNum.ToString() + "-" + (folderNum + 99).ToString();
            //find proper folder for the prefix or create folder
            //\\mgbwvlt\DATA2\CAD_Files\Upholstery\Frame Drawings
            string folderLoc = System.IO.Path.Combine(@"\\mgbwvlt\DATA2\CAD_Files\Upholstery\Frame Drawings", folderName, prefix);
            //string fileLoc = System.IO.Path.Combine(folderLoc, fileName);
            //does the folder exist?
            if (!System.IO.File.Exists(folderLoc))
            { Directory.CreateDirectory(folderLoc); }

            return folderLoc;
            #endregion
        }


        #region Methods for Publishing

        //TODO archive files of same name
        //if files already exist of the PDF/DWGs, then archive the older files

        //creates a PDF in the same location as Drawing unless Flagged for vendor, then it saves in the vendor folder
        public static void exportPDF(bool vendor = false, bool release = false)
        {
            //Inventor.Application InvApp =.InvApp;
            Inventor.Application InvApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            // Get the PDF translator Add-In.

            Inventor.TranslatorAddIn PDFAddin = InvApp.ApplicationAddIns.ItemById["{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}"] as Inventor.TranslatorAddIn;

            //Set a reference to the active document (the document to be published).
            Inventor.DrawingDocument oDocument = InvApp.ActiveDocument as Inventor.DrawingDocument;
            Inventor.TranslationContext oContext = InvApp.TransientObjects.CreateTranslationContext();
            oContext.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism;


            // Create a NameValueMap object
            Inventor.NameValueMap oOptions = InvApp.TransientObjects.CreateNameValueMap();
            //Inventor.NameValueMap oOptions.Value["Sheet_Range"] = PrintRangeEnum.kPrintAllSheets;

            // Create a DataMedium object
            Inventor.DataMedium oDataMedium = InvApp.TransientObjects.CreateDataMedium();

            string filePath = oDocument.FullFileName;
            //sems like the most awkward way to get a string without an extension
            //can i just use Path.GetFileNameWithoutExtension
            filePath = filePath.Substring(0, filePath.Length - 4); //***was after the if which overrided the location everytime

            //for output to a vendor file, add "vendor" to end of directory before adding fileName
            if (vendor)
            {
                // find/create a vendor folder in the encompassing directory
                string folderLoc = findVendorFolder(System.IO.Path.GetDirectoryName(oDocument.FullFileName));
                // find/create a folder to later be compressed in the vendor directory
                string compressedFolderLoc = findCompressedFolder(folderLoc, System.IO.Path.GetFileNameWithoutExtension(oDocument.DisplayName));

                //filePath = System.IO.Path.Combine(folderLoc, System.IO.Path.GetFileName(oDocument.FullFileName));  if saving to vendor instead of saving to compressed folder
                //filePath = System.IO.Path.Combine(compressedFolderLoc, System.IO.Path.GetFileName(oDocument.FullFileName));
                filePath = System.IO.Path.Combine(compressedFolderLoc, System.IO.Path.GetFileNameWithoutExtension(oDocument.FullFileName));
            }


            //in the event of releasing a folder find the style ID in the released folder or create one if new
            if (release)
            {
                string folderLoc = ReleasedFolderLocation(oDocument.FullFileName);
                filePath = System.IO.Path.Combine(folderLoc, System.IO.Path.GetFileNameWithoutExtension(oDocument.FullFileName));
            }

            // Check whether the translator has 'SaveCopyAs' options
            if (PDFAddin.HasSaveCopyAsOptions[oDocument, oContext, oOptions])
            {
                // Options for drawings...
                oOptions.Value["All_Color_AS_Black"] = 1;
                oOptions.Value["Remove_Line_Weights"] = 1;
                oOptions.Value["Vector_Resolution"] = 400;
                // oOptions.Value["Sheet_Range"] = Inventor.PrintRangeEnum.kPrintSheetRange;
                oOptions.Value["Sheet_Range"] = Inventor.PrintRangeEnum.kPrintCurrentSheet;
                //oOptions.Value["Sheet_Range"] = Inventor.PrintRangeEnum.kPrintSheetRange[1];
                oOptions.Value["Custom_Begin_Sheet"] = 2;
                oOptions.Value["Custom_End_Sheet"] = 4;
            }
            //todo is this a duplicate drawing doc oDocument
            DrawingDocument oDrawDoc = InvApp.ActiveDocument as DrawingDocument;

            foreach (Sheet oSheet in oDrawDoc.Sheets)
            {
                //ignore cutlist layouts
                if (!oSheet.Name.ToUpper().Contains("CUTLIST") &&
                    !oSheet.Name.ToUpper().Contains("NEST") &&
                    !oSheet.Name.ToUpper().Contains("CUTSHEET") &&
                    !oSheet.Name.ToUpper().Contains("SPRING") &&
                    !oSheet.Name.ToUpper().Contains("SU"))
                {
                    oSheet.Activate();

                    string sheetName = oSheet.Name.Split(':')[0];

                    //part# + sheetname
                    // Set the destination file name
                    if (oSheet.Name == "Sheet:1" || oSheet.Name.ToUpper().Contains("OVERVIEW"))
                    { oDataMedium.FileName = filePath + ".pdf"; }
                    else
                    { oDataMedium.FileName = filePath + "-" + sheetName + ".pdf"; }

                    //archive old files in formal folder
                    if(release)
                    {
                        //check if a file exists in location
                        if(System.IO.File.Exists(oDataMedium.FileName))
                        { archiveFile(oDataMedium.FileName); }
                        AddWaterMark("FOR INTERNAL USE ONLY");
                    }

                    // Publish document.
                    PDFAddin.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium);
                    RemoveWaterMarks();
                }
            }
        }

        public static void exportDWG(bool vendor = false, bool release = false)
        {
            //Inventor.Application InvApp =.InvApp;
            Inventor.Application InvApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            // Get the PDF translator Add-In.

            Inventor.TranslatorAddIn DWGAddin = InvApp.ApplicationAddIns.ItemById["{C24E3AC2-122E-11D5-8E91-0010B541CD80}"] as Inventor.TranslatorAddIn;


            //Set a reference to the active document (the document to be published).
            Inventor.DrawingDocument oDocument = InvApp.ActiveDocument as Inventor.DrawingDocument;
            Inventor.TranslationContext oContext = InvApp.TransientObjects.CreateTranslationContext();
            oContext.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism;


            // Create a NameValueMap object
            Inventor.NameValueMap oOptions = InvApp.TransientObjects.CreateNameValueMap();
            //Inventor.NameValueMap oOptions.Value["Sheet_Range"] = PrintRangeEnum.kPrintAllSheets;

            // Create a DataMedium object
            Inventor.DataMedium oDataMedium = InvApp.TransientObjects.CreateDataMedium();

            string filePath = oDocument.FullFileName;
            //for output to a vendor file, add "vendor" to end of directory before adding fileName
            if (vendor)
            {
                //string folderLoc = findVendorFolder(System.IO.Path.GetDirectoryName(oDocument.FullFileName));
                //filePath = System.IO.Path.Combine(folderLoc, System.IO.Path.GetFileName(oDocument.FullFileName));
                
                // find/create a vendor folder in the encompassing directory
                string folderLoc = findVendorFolder(System.IO.Path.GetDirectoryName(oDocument.FullFileName));
                // find/create a folder to later be compressed in the vendor directory
                string compressedFolderLoc = findCompressedFolder(folderLoc, System.IO.Path.GetFileNameWithoutExtension(oDocument.DisplayName));

                //filePath = System.IO.Path.Combine(folderLoc, System.IO.Path.GetFileName(oDocument.FullFileName));  if saving to vendor instead of saving to compressed folder
                filePath = System.IO.Path.Combine(compressedFolderLoc, System.IO.Path.GetFileName(oDocument.FullFileName));
            }
            filePath = filePath.Substring(0, filePath.Length - 4);

            if (DWGAddin.HasSaveCopyAsOptions[oDocument, oContext, oOptions])
            {
                // Options for drawings...
                oOptions.Value["All_Color_AS_Black"] = 1;
                oOptions.Value["Remove_Line_Weights"] = 1;

                // oOptions.Value["Sheet_Range"] = Inventor.PrintRangeEnum.kPrintSheetRange;
                oOptions.Value["Sheet_Range"] = Inventor.PrintRangeEnum.kPrintCurrentSheet;
                //oOptions.Value["Sheet_Range"] = Inventor.PrintRangeEnum.kPrintSheetRange[1];
                oOptions.Value["Custom_Begin_Sheet"] = 2;
                oOptions.Value["Custom_End_Sheet"] = 4;

                oOptions.set_Value("DwgVersion", 23);
                oOptions.set_Value("Export_Acad_IniFile", filePath + "export.ini");
            }
            //todo is this a duplicate draw document of oDocument?
            DrawingDocument oDrawDoc = InvApp.ActiveDocument as DrawingDocument;

            foreach (Sheet oSheet in oDrawDoc.Sheets)
            {
                //ignore certain sheet names
                //ignore cutlist layouts
                if (!oSheet.Name.ToUpper().Contains("SPRING") &&
                    !oSheet.Name.ToUpper().Contains("SU"))
                {
                    oSheet.Activate();

                    string sheetName = oSheet.Name.Split(':')[0];

                    //part# + sheetname
                    // Set the destination file name
                    if (oSheet.Name == "Sheet:1" || oSheet.Name.ToUpper().Contains("OVERVIEW"))
                    { oDataMedium.FileName = filePath + ".dwg"; }
                    else
                    { oDataMedium.FileName = filePath + "-" + sheetName + ".dwg"; }


                    // Publish document.
                    DWGAddin.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium);
                }
            }

            //todo if we want to release DWG files to formal folder
            //need to loop through the output code again for released location, unless code can only have 1 status and just has to be called multiple times

            //issues I read about on line. People couldnt loop through sheets, had to create a new drawing for each sheet and then delete the "new drawing"
            //if (File.Exists(SAVE_PATH + sSaveFile + "_Sheet_2.dwg"))
            //    File.Delete(SAVE_PATH + sSaveFile + "_Sheet_2.dwg");

            //if (File.Exists(SAVE_PATH + sSaveFile))
            //    File.Delete(SAVE_PATH + sSaveFile);

            //if (File.Exists(SAVE_PATH + sSaveFile + "_Sheet_1.dwg"))
            //    File.Move(SAVE_PATH + sSaveFile + "_Sheet_1.dwg",
            //       SAVE_PATH + sSaveFile);
        }

        /// <summary>
        /// Place old files into a folder in the same location labled "OlD"
        /// rename the file to include the last modified date
        /// </summary>
        /// <param name="fi"></param>
        private static void archiveFile(string fileFullPath)
        {
            FileInfo fi = new FileInfo(fileFullPath);
            string directory = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(fileFullPath), "OLD");
            //check if a folder named "OLD" exists in the same folder
            if(! System.IO.File.Exists(directory))
            { Directory.CreateDirectory(directory); }//create one if it doesnt
            //rename file to include its last modified date as part of the filename (before extension)
            string newFileName = System.IO.Path.Combine(directory,
                System.IO.Path.GetFileNameWithoutExtension(fileFullPath) +
                "_(" + fi.LastWriteTime.ToString("MM-dd-yyyy") + ")" +
                fi.Extension);
            //string newFileName = System.IO.Path.Combine(directory,
            //    System.IO.Path.GetFileNameWithoutExtension(fileFullPath)+
            //    "_(old write time)"+
            //    fi.Extension);
            //check if a file already exists????????
            //move or delete
            System.IO.File.Move(fileFullPath, newFileName);
        }

        public static void exportModelDWG(bool vendor = false)
        {

            //try to get asm/part from drawing name
            Inventor.Application InvApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            //Set a reference to the active document
            Inventor.DrawingDocument oDrawDocument = InvApp.ActiveDocument as Inventor.DrawingDocument;
            string drawFileLoc = oDrawDocument.FullFileName;
            // search for a part or asembly file with the same name * excluding anything in a folder of "OldVersions"

            string modelLoc = "";

            #region Find model Location
            string path = System.IO.Path.GetFullPath(drawFileLoc);
            string name = System.IO.Path.GetFileNameWithoutExtension(drawFileLoc);

            //find if the asm (iam) exists
            string modelFileName = InvApp.DesignProjectManager.ResolveFile(path, name + ".iam");
            if (modelFileName == "")
            {
                //find if the part (ipt) exists bc asm doesn't
                modelFileName = InvApp.DesignProjectManager.ResolveFile(path, name + ".ipt");

                if (modelFileName != "")
                    modelLoc = modelFileName;
                else
                {
                    //no model exists matching the name of the drawing
                    //probably just return then
                }
            }
            else
                modelLoc = modelFileName;
            #endregion

            //open model file
            Document oDocument = InvApp.Documents.Open(modelLoc, true);
            //todo do I need to break code here if open faile??!!
            oDocument.Activate();

            //good place to update and save the modelfile
            oDocument.Update();
            oDocument.Save2(true);

            //need a translator for assemblies and parts
            Inventor.TranslatorAddIn DWGAddin = InvApp.ApplicationAddIns.ItemById["{C24E3AC2-122E-11D5-8E91-0010B541CD80}"] as Inventor.TranslatorAddIn;

            //Set a reference to the active document (the document to be published).
            //Inventor.DrawingDocument oDocument = InvApp.ActiveDocument as Inventor.DrawingDocument;
            Inventor.TranslationContext oContext = InvApp.TransientObjects.CreateTranslationContext();
            oContext.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism;

            // Create a NameValueMap object
            Inventor.NameValueMap oOptions = InvApp.TransientObjects.CreateNameValueMap();

            // Create a DataMedium object
            Inventor.DataMedium oDataMedium = InvApp.TransientObjects.CreateDataMedium();

            string filePath = oDocument.FullFileName;
            //for output to a vendor file, add "vendor" to end of directory before adding fileName
            if (vendor)
            {
                // find/create a vendor folder in the encompassing directory
                string folderLoc = findVendorFolder(System.IO.Path.GetDirectoryName(oDocument.FullFileName));
                // find/create a folder to later be compressed in the vendor directory
                string compressedFolderLoc = findCompressedFolder(folderLoc, System.IO.Path.GetFileNameWithoutExtension(oDocument.DisplayName));

                //filePath = System.IO.Path.Combine(folderLoc, System.IO.Path.GetFileName(oDocument.FullFileName));  if saving to vendor instead of saving to compressed folder
                filePath = System.IO.Path.Combine(compressedFolderLoc, System.IO.Path.GetFileName(oDocument.FullFileName));
            }
            filePath = filePath.Substring(0, filePath.Length - 4);

            // Check whether the translator has 'SaveCopyAs' options
            if (DWGAddin.HasSaveCopyAsOptions[oDocument, oContext, oOptions])
            {
                // Options for drawings...
                oOptions.Value["All_Color_AS_Black"] = 1;
                oOptions.Value["Remove_Line_Weights"] = 1;

                // oOptions.Value["Sheet_Range"] = Inventor.PrintRangeEnum.kPrintSheetRange;
                oOptions.Value["Sheet_Range"] = Inventor.PrintRangeEnum.kPrintCurrentSheet;
                //oOptions.Value["Sheet_Range"] = Inventor.PrintRangeEnum.kPrintSheetRange[1];
                oOptions.Value["Custom_Begin_Sheet"] = 2;
                oOptions.Value["Custom_End_Sheet"] = 4;

                oOptions.set_Value("DwgVersion", 23);
                oOptions.set_Value("Export_Acad_IniFile", filePath + "export.ini");
            }

            //create fileName
            oDataMedium.FileName = filePath + "_3D.dwg";

            // Publish document.
            DWGAddin.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium);
        }

        //TODO naming scheme to account for specs accross multiple sheets
        public static void exportPartsList(bool release = false)
        {
            //get active document (assume a drawing of an assembly)
            //Inventor.Application InvApp =.InvApp;
            Inventor.Application InvApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            //Set a reference to the active document (the document to be published).
            Inventor.TranslationContext oContext = InvApp.TransientObjects.CreateTranslationContext();
            oContext.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism;

            DrawingDocument oDrawDoc = InvApp.ActiveDocument as DrawingDocument;
            //Sheet oSheet = oDrawDoc.ActiveSheet;
            Sheets sheetlist = oDrawDoc.Sheets;

            string directoryName = System.IO.Path.GetDirectoryName(oDrawDoc.FullFileName);
            if(release)
            {
                //search through styleIds and create folder if not there
                directoryName = ReleasedFolderLocation(oDrawDoc.DisplayName); 
            }

            foreach (Sheet oSheet in sheetlist)
            {
                if (!oSheet.Name.ToUpper().Contains("SPRING"))
                {
                    PartsLists oList = oSheet.PartsLists;

                    //get the partslist
                    int counter = 1;
                    foreach (PartsList BoM in oList)                    {

                        //name convention folder of drawing + Drawing Name + SheetName + Parts List (counter)
                        string fileName = String.Format(@"{0}\{1}_{2}_Parts List({3}).xls",
                            directoryName,
                            System.IO.Path.GetFileNameWithoutExtension(oDrawDoc.FullFileName),
                            oSheet.Name.Trim().Replace(":", " "),
                            counter
                            );

                        //using the data of the table, create an xls doc
                        //save in same folder as drawing
                        BoM.Export(fileName.Trim(), PartsListFileFormatEnum.kMicrosoftExcel);

                        counter++;
                    }
                }
            }
        }        
        #endregion


        #region Methods for creating Drawings

        //creates the basic Overview drawing for a part or asm
        public static void CreateOverViewDrawing()
        {
            //get inventor app
            Inventor.Application m_inventorApp = getInventor();

            _Document modelDoc = m_inventorApp.ActiveDocument;

            #region Read iProperties for metaData
            string metaDataTags = "";
            //get attribute info from the model
            //have to loop thorugh attributes through a property set
            //look for specific expected name in iProperties
            //tags for metaData is named "Keywords"
            foreach (PropertySet ps in modelDoc.PropertySets)
            {
                //find the correct properties tab
                //correct tab is ??? Inventor Document Summary Information
                if (ps.Name.ToString() == "Inventor Summary Information")
                {
                    foreach (Property p in ps)
                    {
                        if (p.Name.ToString() == "Keywords")
                        {
                            //grab the keywords to add to the drawing later
                            metaDataTags = p.Value.ToString();
                        }
                    }
                }
            }
            #endregion

            //TODO- try/catch ask user for template if default fails
            //attempt to get the default drawing template
            //if unavailable, then prompt user to locate *******************************!!!!!!!!!!!!!!!!!!!!!!!!!!            
            //****maybe try get the default template first, if null then get the template folder and add constant name
            //if assumed name and template location doesnt work, then request feedback from user*****

            //"C:\\Users\\tylerhenderson\\Desktop\\OpenWork\\Templates\\en-US\\English\\Fractional.idw"
            //get default location instead of static
            string defaultDraw = m_inventorApp.DesignProjectManager.ActiveDesignProject.TemplatesPath;
            defaultDraw = System.IO.Path.Combine(defaultDraw, "Fractional.idw");

            DrawingDocument oDrawDoc = m_inventorApp.Documents.Add(
                DocumentTypeEnum.kDrawingDocumentObject,
                defaultDraw,
                true) as DrawingDocument;

            //sheet for overview
            Sheet oSheet = oDrawDoc.ActiveSheet;
            oSheet.Name = "Overview";

            TransientGeometry oTG = m_inventorApp.TransientGeometry;
            ComponentDefinition oCompDef = null;

            #region add MetaData to iProperties
            foreach (PropertySet ps in oDrawDoc.PropertySets)
            {
                //find the correct properties tab
                if (ps.Name.ToString() == "Inventor Summary Information")
                {
                    foreach (Property p in ps)
                    {
                        if (p.Name.ToString() == "Keywords")
                        {
                            //grab the keywords to add to the drawing later
                            p.Value = metaDataTags;
                        }
                    }
                }
            }
            #endregion

            if (modelDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyDocument oDoc = modelDoc as AssemblyDocument;
                oCompDef = oDoc.ComponentDefinition as ComponentDefinition;

            }
            else if (modelDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartDocument oDoc = modelDoc as PartDocument;
                oCompDef = oDoc.ComponentDefinition as ComponentDefinition;
            }
            else
                return;

            Box rngBx = oCompDef.RangeBox;

            //assemblies will most likely prefer a front view, parts may prefer to find the shallowest view (FIND SIMILAR CODE TO CUTLIST)
            //probably move to a separate method so that later it can automate creating part drawings for an assembly

            #region add views
            //simple algo to determine view
            //smallest dimension of view should be the axis aimed away from the view
            //  x   =   front
            //  y   =   side
            //  z   =   top
            double xDim = rngBx.MaxPoint.X - rngBx.MinPoint.X;
            double yDim = rngBx.MaxPoint.Y - rngBx.MinPoint.Y;
            double zDim = rngBx.MaxPoint.Z - rngBx.MinPoint.Z;
            //calc scale
            double oScale = calcAdjustScale(xDim, yDim, zDim, 14, 10); //may need to adjust target quadrant size
            Fraction wholeScale = RealToFraction(oScale, .1);
            oScale = Convert.ToDouble(wholeScale.N) / Convert.ToDouble(wholeScale.D);

            //place Base and consequential ortho views
            DrawingView oFrontView = oSheet.DrawingViews.AddBaseView(modelDoc,
                oTG.CreatePoint2d((xDim / 2) * oScale + 3.5, (yDim / 2) * oScale + 5),
                oScale,//need to calculate scale*******
                ViewOrientationTypeEnum.kFrontViewOrientation,//verify if Front is view we want. isnt there a "Master" View?
                DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);

            //add projected ORTHO views
            //side view
            DrawingView oSideView = oSheet.DrawingViews.AddProjectedView(oFrontView,
                oTG.CreatePoint2d((oFrontView.Center.X + oFrontView.Width / 2 + zDim / 2 * oScale) + 2, oFrontView.Center.Y),
                DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);
            //location is the center of the base view, plus half the widths of each view plus some portion for a gap
            //top view
            DrawingView oTopView = oSheet.DrawingViews.AddProjectedView(oFrontView,
                oTG.CreatePoint2d(oFrontView.Center.X, (oFrontView.Center.Y + oFrontView.Height / 2 + zDim / 2 * oScale) + 1),
                DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);
            //location is the center of the base view, plus half the heights of each view plus somE portion for a gap
            //create iso view in corner
            //TODO create a new view so that drawing states can be different
            //DrawingView oIsoView = oSheet.DrawingViews.AddProjectedView(oFrontView,
            //    oTG.CreatePoint2d(40, 30),
            //    DrawingViewStyleEnum.kShadedDrawingViewStyle);
            DrawingView oIsoView = oSheet.DrawingViews.AddBaseView(modelDoc,
                oTG.CreatePoint2d(40, 30),
                oScale,
                ViewOrientationTypeEnum.kIsoTopLeftViewOrientation,
                DrawingViewStyleEnum.kShadedDrawingViewStyle);

            #endregion

            //TODO - autodimension the ortho views
            //would be good to automate some Dimensions right here
            //********************        


            //*************only add a partslist if an asm, not for parts
            if (modelDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {//add in the BOM
                AssemblyDocument oDoc = modelDoc as AssemblyDocument;
                AssemblyComponentDefinition ooCompDef = oDoc.ComponentDefinition;
                ooCompDef.BOM.StructuredViewEnabled = true; //default is false, has to be enabled to create a BOM
                CreatePartsList(oIsoView, oSheet, m_inventorApp);
            }
        }


        //TODO, complete the cases for different interpretations
        //creates a cutlist of all parts in the assembly, in its best attempt at the detail view for each part
        public static void CreateCutlistDrawing()
        {
            //get inventor app
            Inventor.Application m_inventorApp = getInventor();

            switch (m_inventorApp.ActiveDocument.DocumentType)
            {
                case DocumentTypeEnum.kDrawingDocumentObject:
                    //create cutlist in drawing

                    //get the name of the model for the drawing *assumes typical name conventions
                    DrawingDocument oDrawing = (DrawingDocument)m_inventorApp.ActiveDocument;
                    string fileName = oDrawing.FullFileName;
                    //assuming the asm is in the same folder of the drawing
                    string asmLoc = findModelFromDrawing(fileName, m_inventorApp);

                    //confirm that this is a drawing of an assembly not a part before running it
                    //probably just check that the extension is correct

                    //run command
                    createCutlistSheet(asmLoc, oDrawing, m_inventorApp);
                    break;
                case DocumentTypeEnum.kAssemblyDocumentObject:
                    //try to open a drawing then create a cutlist
                    break;
                default:
                    return;

            }
        }

        //creates PDFs and DWGs of the best sheets in the drawing and the 3D model
        public static void PrepForVendors()
        {
            //update the drawing field that the specs have been updated and who did it (if field exists)
            //in order to keep the same titleblock/drawing we have been using I'm going to use:
            //eng approv for a revision tracker
            //mfg approv for a released tracker
            //as well as updating status. if a drawing is new it is by default work in progress
            //on prepping for a vendor, assuming this is intended to send to a vendor, I'll change to 

            Inventor.Application InvApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            //Set a reference to the active document (the document to be published).
            Inventor.DrawingDocument oDrawDoc = InvApp.ActiveDocument as Inventor.DrawingDocument;
            #region setStatus

            foreach (PropertySet ps in oDrawDoc.PropertySets)
            {
                switch (ps.Name.ToString())
                {
                    case "Inventor Summary Information":
                        //loop through the different properties
                        break;
                    case "Inventor Document Summary Information":
                        //loop through the different properties
                        break;
                    case "Design Tracking Properties":
                        foreach (Property p in ps)
                        {
                            switch (p.Name.ToString())
                            {
                                case "Engr Approved By": //12
                                    p.Value = InvApp.GeneralOptions.UserName;
                                    break;
                                case "Engr Date Approved": //13
                                    p.Value = DateTime.Now;
                                    break;
                                case "Mfg Approved By": //34
                                                        //code to change value?
                                    break;
                                case "Mfg Date Approved": //35
                                                          //code to change value?
                                    break;
                                case "Design Status": //40
                                    //1 Work in Progress
                                    //2 Pending
                                    //3 Released
                                    p.Value = 2;
                                    break;
                            }
                        }
                        break;
                    case "Inventor User Defined Properties":
                        //loop through the different properties
                        break;

                }
            }

            #endregion

            //save and update file
            oDrawDoc.Update();
            oDrawDoc.Save2(true);
            //oDrawDoc.Save();
            
            //TODO add a watermark here

            //creates a PDF of all sheets in the folder where the drawing is saved
            //TODO add watermark here for vendors
            exportPDF(true);
            //REmove the watermark here and re-save

            //dwgs of all sheets
            exportDWG(true);

            //TODO remove all watermarks

            //dwg of the 3D model
            exportModelDWG(true);

            //create a compressed file of the folder of specs then delete the folder
            ZipCompressFolder();
        }

        public static void PrepForRelease()
        {
            Inventor.Application InvApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            //Set a reference to the active document (the document to be published).
            Inventor.DrawingDocument oDrawDoc = InvApp.ActiveDocument as Inventor.DrawingDocument;

            //update titleblock for a release date and status to released
            #region setStatus

            foreach (PropertySet ps in oDrawDoc.PropertySets)
            {
                switch (ps.Name.ToString())
                {
                    case "Inventor Summary Information":
                        //loop through the different properties
                        break;
                    case "Inventor Document Summary Information":
                        //loop through the different properties
                        break;
                    case "Design Tracking Properties":
                        foreach (Property p in ps)
                        {
                            switch (p.Name.ToString())
                            {
                                case "Engr Approved By": //12
                                    
                                    break;
                                case "Engr Date Approved": //13
                                    
                                    break;
                                case "Mfg Approved By": //34
                                    p.Value = InvApp.GeneralOptions.UserName;
                                    break;
                                case "Mfg Date Approved": //35
                                    p.Value = DateTime.Now;
                                    break;
                                case "Design Status": //40
                                    //1 Work in Progress
                                    //2 Pending
                                    //3 Released
                                    p.Value = 3;
                                    break;
                            }
                        }
                        break;
                    case "Inventor User Defined Properties":
                        //loop through the different properties
                        break;

                }
            }

            #endregion
            //save and update file
            oDrawDoc.Update();
            oDrawDoc.Save2(true);

            //TODO add watermark here
            
            //export excel table of BoM? *not sure if i want to do this obviously requires an assembly drawing
            exportPartsList(true);
            //create PDFs in the home folder and in the released folder (archive old specs)
            exportPDF(true);//update local drive **REMOVE this line????*******
            exportPDF(false, true); //update formal drive

            //TODO remove watermark here
            
        }

        public static void printFitted()
        {
            Inventor.Application InvApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            //Set a reference to the active document (the document to be published).
            Inventor.TranslationContext oContext = InvApp.TransientObjects.CreateTranslationContext();
            oContext.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism;

            DrawingDocument oDrawDoc = InvApp.ActiveDocument as DrawingDocument;
            DrawingPrintManager oPrintMgr = oDrawDoc.PrintManager as DrawingPrintManager;

            string printerName = "D1";
            var printerColl = System.Drawing.Printing.PrinterSettings.InstalledPrinters;
            if(printerColl.Count >0)
            {
                foreach (string name in printerColl)
                {
                    if (name.Contains(printerName))
                    {
                        printerName = name;
                        //stop loop
                        break;
                    }
                }
            }

            //could loop through names for one that contains "D1" or whatever name we want

            oPrintMgr.Printer = printerName;
            oPrintMgr.NumberOfCopies = 1;
            oPrintMgr.ScaleMode = PrintScaleModeEnum.kPrintBestFitScale;//or full scale or whatever
            oPrintMgr.PaperSize = PaperSizeEnum.kPaperSizeDefault; //8.5 x 11 or custom
            oPrintMgr.PrintRange = PrintRangeEnum.kPrintCurrentSheet;

            //determine orientation
            Inventor.Sheet oSheet = oDrawDoc.ActiveSheet;
            //compare and decide Orientaion
            if (oSheet.Width > oSheet.Height)
            { oPrintMgr.Orientation = PrintOrientationEnum.kLandscapeOrientation; }
            else
                oPrintMgr.Orientation = PrintOrientationEnum.kPortraitOrientation;


            oPrintMgr.SubmitPrint();
        }
        #endregion


        #region Methods that arent called by buttons

        //both watermark commands assume that we're in a drawing sheet already
        //TODO add code to verify that the form is a drawing before executing code
        //removes all text from active sheet with watermark text style
        private static void RemoveWaterMarks()
        {
            Inventor.Application InvApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            Inventor.DrawingDocument oDrawDocument = InvApp.ActiveDocument as Inventor.DrawingDocument;
            DrawingStylesManager oDStylesMan = oDrawDocument.StylesManager;
            Sheet oSheet = oDrawDocument.ActiveSheet;

            // pull genNotes, loop through and delete any thing with style "watermark"
            GeneralNotes oGenNotes = oSheet.DrawingNotes.GeneralNotes;
            foreach (GeneralNote note in oGenNotes)
            {
                if (note.TextStyle.Name == "WaterMark Style")
                {
                    note.Delete();

                }
            }
        }

        //adds a text watermark to the active sheet for the text passed
        private static void AddWaterMark(string message)
        {
            Inventor.Application InvApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            Inventor.DrawingDocument oDrawDocument = InvApp.ActiveDocument as Inventor.DrawingDocument;
            TransientGeometry oTG = InvApp.TransientGeometry;
            DrawingStylesManager oDStylesMan = oDrawDocument.StylesManager;
            //check if the style were using already is in the list
            TextStyle oStyle = default(TextStyle);
            bool styleExists = false;
            foreach (TextStyle txtstyle in oDStylesMan.TextStyles)
            {
                if (txtstyle.Name == "WaterMark Style")
                {
                    oStyle = txtstyle;
                    styleExists = true;
                    break;
                }
            }
            if (!styleExists)
            {
                oStyle = (TextStyle)oDStylesMan.TextStyles["Label Text (ANSI)"].Copy("WaterMark Style");
                oStyle.Font = "Swis721 BlkOul BT";
                oStyle.FontSize = (double)1;
                oStyle.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter;
                oStyle.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextMiddle;

                oStyle.Color = InvApp.TransientObjects.CreateColor(255, 0, 0);
            }

            //add watermark to the sheet (in live, this needs to ONLY be the overview sheet)
            //TODO verify if this is the overview sheet, possibly pass the sheet at start?
            // Loop through the available sheeets and only add to the "Overview or Sheet 1"
            Sheet oSheet = oDrawDocument.ActiveSheet;
            Point2d oCtrPoint = oTG.CreatePoint2d(oSheet.Width / 2, oSheet.Height / 2);//center point
            GeneralNotes oGenNotes = oSheet.DrawingNotes.GeneralNotes;
            GeneralNote oGeneralNote = oGenNotes.AddFitted(oCtrPoint, message, oStyle); //pass different watermark labels here

            //rotations in oStyle are limited to 90 deg increments
            oGeneralNote.Rotation = (double)(Math.PI / 6);
        }

        //get the inventor app
        private static Inventor.Application getInventor()
        {
            Inventor.Application m_inventorApp = null;

            try //Try to get an active instance of Inventor
            {
                try
                {
                    m_inventorApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
                }
                catch
                {
                    Type inventorAppType = System.Type.GetTypeFromProgID("Inventor.Application");
                    m_inventorApp = System.Activator.CreateInstance(inventorAppType) as Inventor.Application;
                    //Must be set visible explicitly
                    m_inventorApp.Visible = true;
                    return m_inventorApp;
                }
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Error: couldn't create Inventor instance");
                return null;
            }
            return m_inventorApp;
        }

        //check if a drawing file with the same name of a part exists
        private static string FindDrawingFile(PartDocument partDoc)
        {
            string fullFileName = partDoc.FullFileName;

            //extract the parth from the full filename
            string path = System.IO.Path.GetDirectoryName(fullFileName);

            //filename
            string fileName = System.IO.Path.GetFileNameWithoutExtension(fullFileName);

            //name of file if it was drawing extension
            string searchFile = fileName + ".ids";

            //see if file exists
            if (System.IO.File.Exists(searchFile))
            {
                return searchFile + DateTime.Today.ToString();
                //need to alert existance of file to user and possible open it.
            }

            else return searchFile;
            //need to create a new file and open it
        }

        //given a folder location look for and return the folder location of where to save vendors specs, or create one if one doesnt exist
        private static string findVendorFolder(string containFolder)
        {
            bool createFolder = true;
            string folderLoc = null;

            //look for a folder named "VENDOR" 
            DirectoryInfo dir = new DirectoryInfo(containFolder);
            IEnumerable<DirectoryInfo> folderList = dir.GetDirectories("*", SearchOption.TopDirectoryOnly);

            folderList =
                from folder in folderList
                where folder.Name.ToUpper() == "VENDOR"
                select folder;

            foreach (DirectoryInfo folder in folderList)
            {
                folderLoc = folder.FullName;
                createFolder = false;
            }

            //If no vendor folder is to be found, create one and return it
            if (createFolder)
            {
                folderLoc = System.IO.Path.Combine(containFolder, "VENDOR");
                Directory.CreateDirectory(folderLoc);
                return folderLoc;
            }

            return folderLoc;
        }
        private static string findCompressedFolder(string vendorDirectory, string styleID)
        {
            //need to get styleID
            //compress name scheme is:  StyleID (month-day-year)
            string compressFolderName = styleID + " (" + DateTime.Now.ToString("MM-dd-yy") + ")";
            string compressLocation = System.IO.Path.Combine(vendorDirectory, compressFolderName);
            if (!Directory.Exists(compressLocation))
            { Directory.CreateDirectory(compressLocation); }

            return compressLocation;
        }

        public static void ZipCompressFolder()
        {
            //get app to get name of file we're working with
            Inventor.Application InvApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            //get the active document in theory could be a model or drawing
            Document oDocument = InvApp.ActiveDocument;

            // find/create a vendor folder in the encompassing directory
            string folderLoc = findVendorFolder(System.IO.Path.GetDirectoryName(oDocument.FullFileName));
            // find/create a folder to later be compressed in the vendor directory
            string compressedFolderLoc = findCompressedFolder(folderLoc, System.IO.Path.GetFileNameWithoutExtension(oDocument.DisplayName));
            string saveFileLoc = System.IO.Path.Combine(folderLoc, System.IO.Path.GetFileName(compressedFolderLoc)) + ".zip"; //folder loc with the same name but with .zip

            //one last verification
            if (Directory.Exists(compressedFolderLoc))
            {
                ZipFile.CreateFromDirectory(compressedFolderLoc, saveFileLoc);
                //delete old folder
                Directory.Delete(compressedFolderLoc, true);
            }
        }

        //derive ideal scale
        private static double calcScale(double widthModel, double heightModel)
        {
            //assuming insert area of standard template
            //trying to limit to 1 quadrant of drawing template
            //template is roughly comperable to 8.5 x 11 sheet
            double targetDistX = 25; //use cm instead of in
            double targetDisty = targetDistX * 17 / 22;

            double scaleX = targetDistX / widthModel;
            double scaleY = targetDisty / heightModel;

            //which ever is lower is the scale to work from
            if (scaleX < scaleY)
                return scaleX;
            else
                return scaleY;
        }

        //given dimension of the unit and having assumed quadrants of the orhto views
        //calculate the scale that would fit a drawing into the grid with regards to how small or large all views are
        private static double calcAdjustScale(double xDim, double yDim, double zDim, double quadX = 14, double quadY = 10)
        {
            double scaleX = (quadX * 2) / (xDim + zDim);
            double scaleY = (quadY * 2) / (yDim + zDim);

            //which ever is lower is the scale to work from
            if (scaleX < scaleY)
                return scaleX;
            else
                return scaleY;
        }

        /// <summary>
        /// Binary seek for the value where f() becomes false.
        /// </summary>
        private static void Seek(ref int a, ref int b, int ainc, int binc, Func<int, int, bool> f)
        {
            a += ainc;
            b += binc;

            if (f(a, b))
            {
                int weight = 1;

                do
                {
                    weight *= 2;
                    a += ainc * weight;
                    b += binc * weight;
                }
                while (f(a, b));

                do
                {
                    weight /= 2;

                    int adec = ainc * weight;
                    int bdec = binc * weight;

                    if (!f(a - adec, b - bdec))
                    {
                        a -= adec;
                        b -= bdec;
                    }
                }
                while (weight > 1);
            }
        }

        // "Fraction" code from stackoverflow user Kay Zed
        //https://stackoverflow.com/questions/5124743/algorithm-for-simplifying-decimal-to-fractions/32903747#32903747
        public static Fraction RealToFraction(double value, double accuracy)
        {
            if (accuracy <= 0.0 || accuracy >= 1.0)
            {
                throw new ArgumentOutOfRangeException("accuracy", "Must be > 0 and < 1.");
            }

            int sign = Math.Sign(value);

            if (sign == -1)
            {
                value = Math.Abs(value);
            }

            // Accuracy is the maximum relative error; convert to absolute maxError
            double maxError = sign == 0 ? accuracy : value * accuracy;

            int n = (int)Math.Floor(value);
            value -= n;

            if (value < maxError)
            {
                return new Fraction(sign * n, 1);
            }

            if (1 - maxError < value)
            {
                return new Fraction(sign * (n + 1), 1);
            }

            // The lower fraction is 0/1
            int lower_n = 0;
            int lower_d = 1;

            // The upper fraction is 1/1
            int upper_n = 1;
            int upper_d = 1;

            while (true)
            {
                // The middle fraction is (lower_n + upper_n) / (lower_d + upper_d)
                int middle_n = lower_n + upper_n;
                int middle_d = lower_d + upper_d;

                if (middle_d * (value + maxError) < middle_n)
                {
                    // real + error < middle : middle is our new upper
                    //upper_n = middle_n;
                    //upper_d = middle_d;
                    //^^^^^^^old, struggled with .01 and .001 etc
                    Seek(ref upper_n, ref upper_d, lower_n, lower_d, (un, ud) => (lower_d + ud) * (value + maxError) < (lower_n + un));
                }
                else if (middle_n < (value - maxError) * middle_d)
                {
                    // middle < real - error : middle is our new lower
                    //lower_n = middle_n;
                    //lower_d = middle_d;
                    //^^^^^^^old, struggled with .01 and .001 etc
                    Seek(ref lower_n, ref lower_d, upper_n, upper_d, (ln, ld) => (ln + upper_n) < (value - maxError) * (ld + upper_d));
                }
                else
                {
                    // Middle is our best fraction
                    return new Fraction((n * middle_d + middle_n) * sign, middle_d);
                }
            }
        }

        //creates a parts list.
        private static void CreatePartsList(DrawingView balloonView, Sheet oSheet, Inventor.Application m_inventorApp)//probably need sheet and view
        {
            Border oBorder = oSheet.Border;
            Point2d oPlacementPoint = default(Point2d);
            oPlacementPoint = m_inventorApp.TransientGeometry.CreatePoint2d(0, oSheet.Height);

            //place table to top right of border, unless border doesnt exist
            if ((oBorder != null))
            {
                // A border exists. The placement point
                // is the top-right corner of the border.
                //oPlacementPoint = oBorder.RangeBox.MaxPoint; // top right
                oPlacementPoint = m_inventorApp.TransientGeometry.CreatePoint2d(
                    oBorder.RangeBox.MinPoint.X,
                    oBorder.RangeBox.MaxPoint.Y); //top left
            }

            //BOM view must first be enabled.
            //balloonView.

            // Create the parts list.
            PartsList oPartsList = default(PartsList);
            try
            {
                oPartsList = oSheet.PartsLists.Add(
                    balloonView,
                    oPlacementPoint,
                    PartsListLevelEnum.kFirstLevelComponents,
                    false); //breaking here on some drawings. Possibly just wrap in a try block
            }
            catch
            {
                try
                {
                    oPartsList = oSheet.PartsLists.Add(
                    balloonView,
                    oPlacementPoint,
                    PartsListLevelEnum.kStructuredAllLevels,
                    false);
                }
                catch { return; }
            }
            //breaking bc partlist level
            //2283
            //unresolved Doc:   "Parts Only"                Works
            //                  "Structured"                Works
            //                  "Structured all levels"     Breaks on add partslist
            //                  "First level Components"    Works

            //2284 sofa
            //Nested Doc:       "Parts Only"                Breaks on add partslist
            //                  "Structured"                Breaks on add partslist
            //                  "Structured all levels"     Works
            //                  "First level Components"    Breaks on add partslist

            //edit the partslist
            //--------------------------------

            //font info
            int colCount = 0;
            double fontWidth = oPartsList.DataTextStyle.FontSize * oPartsList.DataTextStyle.WidthScale;

            //remove the description column
            //adjust width of columns fit contents of known widths we want
            foreach (PartsListColumn col in oPartsList.PartsListColumns)
            {
                colCount++;
                if (col.Title == "DESCRIPTION")
                {
                    col.Remove();
                    continue;
                }

                if (col.Title == "QTY" || col.Title == "ITEM")
                { col.Width = 4 * fontWidth / 1.3; }

                if (col.Title == "PART NUMBER")
                {
                    double cellLength = 0;
                    PartsListCell oCell = default(PartsListCell);
                    foreach (PartsListRow row in oPartsList.PartsListRows)
                    {
                        //loop through all rows of the "PART NUMBER" column
                        //while looping remove any part with any flag meant for removal   ie "NO VIS, NO BOM"
                        //track the width of ofther lines, to find the ideal width for column
                        if (row.Visible)
                        {
                            oCell = row[colCount];
                            if (oCell.Value.Contains("NO VIS") || oCell.Value.Contains("NO BOM"))
                            {
                                row.Visible = false;
                                continue;
                            }

                            if (cellLength < oCell.Value.Length)
                                cellLength = oCell.Value.Length;
                        }
                    }
                    col.Width = fontWidth * (cellLength + 2) / 1.3; //adding 2 for padding
                }

            }


            //split at 10 if there are too many rows
            if (oPartsList.PartsListRows.Count > 13)
            {
                //align partslist to right
                oPartsList.WrapLeft = false;

                //oPartsList.TableDirection;
                oPartsList.MaximumRows = 10;
            }

            oPartsList.Renumber();
        }

        public static string findModelFromDrawing(string fileLoc, Inventor.Application m_inventorApp)
        {
            string path = System.IO.Path.GetFullPath(fileLoc);
            string name = System.IO.Path.GetFileNameWithoutExtension(fileLoc);

            //find if the asm (iam) exists
            string modelFileName = m_inventorApp.DesignProjectManager.ResolveFile(path, name + ".iam");
            if (modelFileName == "")
            {
                //find if the part (ipt) exists bc asm doesn't
                modelFileName = m_inventorApp.DesignProjectManager.ResolveFile(path, name + ".ipt");

                if (modelFileName != "")
                    return modelFileName;
                else
                    return "";
            }
            return modelFileName;
        }

        private static void createCutlistSheet(string asmFileLoc, DrawingDocument oDraw, Inventor.Application m_inventorApp)
        {
            //create new sheet labeled CUTLIST
            Sheets sheetList = oDraw.Sheets;
            sheetList.Add(
                DrawingSheetSizeEnum.kCustomDrawingSheetSize,
                PageOrientationTypeEnum.kLandscapePageOrientation,
                "CUTLIST",
                244,
                122);

            Sheet oSheet = sheetList[sheetList.Count];
            oSheet.Activate();

            //get document of assembly to loop through
            AssemblyDocument asmDoc = (AssemblyDocument)m_inventorApp.Documents.ItemByName[asmFileLoc];

            //loop through partslist (recursively)
            traverseAssembly(asmDoc.ComponentDefinition.Occurrences, m_inventorApp);

        }

        private static void traverseAssembly(ComponentOccurrences oOccs, Inventor.Application m_inventorApp)
        {
            foreach (ComponentOccurrence occ in oOccs)
            {
                //check if occurance is a subassembly
                if (occ.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                { traverseSubAssembly(occ.SubOccurrences, m_inventorApp); }
                else
                {
                    //add to cutlist
                    placePartInCutlist(occ, m_inventorApp);
                }
            }
        }

        private static void traverseSubAssembly(ComponentOccurrencesEnumerator oOccs, Inventor.Application m_inventorApp)
        {
            //iterate through all occurances in collection
            foreach (ComponentOccurrence occ in oOccs)
            {
                //check if occurance is a subassembly
                if (occ.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                { traverseSubAssembly(occ.SubOccurrences, m_inventorApp); }
                else
                {
                    //add to  cutlist *ignores cutlist, cutsheet, nest, su
                    placePartInCutlist(occ, m_inventorApp);
                }
            }
        }

        //place each component occurance onto the cutlist, components have been filtered to part level not subassembly
        private static void placePartInCutlist(ComponentOccurrence oOcc, Inventor.Application m_inventorApp)
        {

            if (oOcc.Name.ToUpper().Contains("NO BOM") || oOcc.Name.ToUpper().Contains("NO VIS"))
                return;

            _Document doc = oOcc.Definition.Document as _Document;

            //do not place the part if it is found in a "hardware" folder
            if (doc.FullFileName.ToUpper().Contains("HARDWARE") || doc.FullFileName.ToUpper().Contains("LEGS"))
                return;
            else
            {
                DrawingDocument oDrawing = (DrawingDocument)m_inventorApp.ActiveDocument;
                Sheet oSheet = oDrawing.ActiveSheet;
                TransientGeometry oTG = m_inventorApp.TransientGeometry;
                DrawingViews oViews = oSheet.DrawingViews;

                Box rangeBox = oOcc.Definition.RangeBox;

                ViewOrientationTypeEnum viewOrient = ViewOrientationTypeEnum.kTopViewOrientation;
                double padding = 10;
                bool rotate = false;
                #region determine View
                //simple algo to determine view
                //smallest dimension of view should be the axis aimed away from the view
                //  x   =   front
                //  y   =   side
                //  z   =   top
                double xDim = rangeBox.MaxPoint.X - rangeBox.MinPoint.X;
                double yDim = rangeBox.MaxPoint.Y - rangeBox.MinPoint.Y;
                double zDim = rangeBox.MaxPoint.Z - rangeBox.MinPoint.Z;

                //compare for smallest dimension
                if (xDim < yDim && xDim < zDim)
                {
                    viewOrient = ViewOrientationTypeEnum.kRightViewOrientation;
                    if (zDim > yDim)
                    {
                        padding = yDim / 2 + 10;
                        rotate = true;
                    }
                    else
                        padding = zDim / 2 + 10;
                }
                else if (yDim < xDim && yDim < zDim)
                {
                    viewOrient = ViewOrientationTypeEnum.kTopViewOrientation;
                    if (xDim > zDim)
                    {
                        padding = zDim / 2 + 10;
                        rotate = true;
                    }
                    else
                        padding = xDim / 2 + 10;
                }
                else if (zDim < xDim && zDim < yDim)
                {
                    viewOrient = ViewOrientationTypeEnum.kFrontViewOrientation;
                    if (xDim > yDim)
                    {
                        padding = yDim / 2 + 10;
                        rotate = true;
                    }
                    else
                        padding = xDim / 2 + 10;
                }
                #endregion

                //find furthest Y dimension iterating through views.
                double x = 0;
                foreach (DrawingView view in oViews)
                {
                    if ((view.Center.X + view.Width / 2) > x)
                        x = view.Center.X + (view.Width / 2);
                }

                x += padding;

                DrawingView oView = oSheet.DrawingViews.AddBaseView(
                    doc,
                    oTG.CreatePoint2d(x, 61),
                    1,
                    viewOrient,
                    DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);

                //rotate if needed
                if (rotate)
                    oView.Rotation = oView.Rotation + Math.PI / 2;

                //label each piece
                #region part labeling
                //leader annotation text
                //leader line start top left corner
                //text vertical
                //text near top, offset by constant + % of total length  ????does length x,y,z change based on view/rotation??

                //drawing curve segment *typically selected by user


                //until i can get a curve section without a user selecting, I should just stick regular text inserted at the center

                //get Drawing Styles Manager
                DrawingStylesManager oDStylesMan = default(DrawingStylesManager);
                DrawingDocument oDoc = (DrawingDocument)m_inventorApp.ActiveDocument;//________
                oDStylesMan = oDoc.StylesManager;
                //check if textstyle exists already?


                //get material to add to label for vendors to recognize
                //get property set
                PropertySets oPropSets = doc.PropertySets;
                PropertySet oPropSet = oPropSets["Design Tracking Properties"];
                Property oProp = oPropSet["Material"];
                string oMatType = oProp.Value as string;
                if (oMatType == "Generic")
                    oMatType = "Standard Thickness";

                GeneralNotes oGenNotes = oSheet.DrawingNotes.GeneralNotes;
                string partName = oOcc.Name + "_" + oMatType;
                Point2d center = oTG.CreatePoint2d(x, 61);

                double rightX = x + oView.Width / 2;
                Point2d rightOfView = oTG.CreatePoint2d(rightX, 61);
                //Math.abs(rangeBox.MaxPoint.Y - rangeBox.MinPoint.Y))/ 2 + rangeBox.MinPoint.Y
                TextStyle oStyle = default(TextStyle);
                bool styleExists = false;
                foreach (TextStyle txtstyle in oDStylesMan.TextStyles)
                {
                    if (txtstyle.Name == "Cutlist Style")
                    {
                        oStyle = txtstyle;
                        styleExists = true;
                        break;
                    }

                }
                if (!styleExists)
                {
                    oStyle = (TextStyle)oDStylesMan.TextStyles["Label Text (ANSI)"].Copy("Cutlist Style");
                    oStyle.FontSize = 1.27;
                    oStyle.Rotation = (Math.PI / 2);
                }

                //Instead of putting label on center, it needs to be at the far edge of the view
                //GeneralNote oGenNote = oGenNotes.AddFitted(center, partName, oStyle);
                GeneralNote oGenNote = oGenNotes.AddFitted(rightOfView, partName, oStyle);


                //GenNote.Rotation = (Math.PI / 2);
                //oGenNote.TextStyle.FontSize = 1.27; //overrides all text sizes

                #endregion
            }
        }

        //Update released PDFs folder from formals inventor folder
        public static void updateReleasedPDFs_OLDIGNORE(string dirPDFs, string dirCAD)
        {
            //get inventor app
            Inventor.Application m_inventorApp = getInventor();

            //create a list of pdfs from released spec folder (formal)
            #region pdfCollection
            DirectoryInfo dirPDF = new DirectoryInfo(dirPDFs);
            IEnumerable<FileInfo> fileInfoEnumPDF = dirPDF.GetFiles(".", SearchOption.AllDirectories);

            //filter for PDFS
            fileInfoEnumPDF =
                from file in fileInfoEnumPDF
                where file.Extension == ".pdf"
                orderby file.DirectoryName
                select file;
            #endregion

            //create list of drawings in formal drive (Charles)
            #region drawingCollection pre-filtered
            DirectoryInfo dirDraw = new DirectoryInfo(dirCAD);
            IEnumerable<FileInfo> fileInfoEnumDraw = dirDraw.GetFiles(".", SearchOption.AllDirectories);

            //filter for drawing files "IDW", filter out back ups in OLD folder
            fileInfoEnumDraw =
                from file in fileInfoEnumDraw
                where file.Extension == ".idw"
                where !file.Directory.ToString().Contains("OldVersions")
                orderby file.DirectoryName
                select file;

            //filter out drawings of parts lists, spring up,
            fileInfoEnumDraw =
                from file in fileInfoEnumDraw
                where !file.Name.ToString().ToUpper().Contains("SPRING")
                where !file.Name.ToString().ToUpper().Contains("PARTS")
                where !file.Name.ToString().ToUpper().Contains("FULL SIZE")
                orderby file.DirectoryName
                select file;

            //filter out drawings that dont have a part or asm of the same name
            fileInfoEnumDraw =
                from file in fileInfoEnumDraw
                where findModelFromDrawing(file.FullName, m_inventorApp) != ""
                orderby file.DirectoryName
                select file;

            #endregion

            //create a list of files to update to released PDFs
            #region process HashSet
            HashSet<FileInfo> processDraws = new HashSet<FileInfo>();

            //New drawings
            //if pdf doesnt exist , then it needs to be re-printed
            var compareCollectionsName =
                from file in fileInfoEnumDraw
                where !fileInfoEnumPDF.Any(x => System.IO.Path.GetFileNameWithoutExtension(x.Name) == System.IO.Path.GetFileNameWithoutExtension(file.Name))
                where !file.Name.ToUpper().Contains("FULL SIZE")
                where !file.Name.ToUpper().Contains("CUTSHEET")
                where !file.Name.ToUpper().Contains("SU")
                select file;

            //compare duplicate names by last modified date. If formal drawing is newer than released PDF, add to list
            //or idw has same name but newer modified date
            //Datetime.compare(t1, t2)
            //t1 is earlier than t2      < 0
            //t1 is same as t2           = 0
            //t1 is later than t2        > 0
            var compareCollectionsDate =
                from file in fileInfoEnumDraw
                where fileInfoEnumPDF.Any(x => System.IO.Path.GetFileNameWithoutExtension(x.Name) == System.IO.Path.GetFileNameWithoutExtension(file.Name) &&
                DateTime.Compare(x.LastWriteTime, file.LastWriteTime) < 0)
                select file;


            List<FileInfo> checkdatahere = compareCollectionsDate.ToList();
            List<FileInfo> checkwhaterver = compareCollectionsName.ToList();
            List<FileInfo> concatedData = compareCollectionsDate.ToList().Concat(compareCollectionsName.ToList()).ToList();
            #endregion

            //Create new PDF of all drawings in list
            //open and activate the file
            m_inventorApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            m_inventorApp.SilentOperation = true;
            foreach (FileInfo fi in concatedData)
            {
                try
                {
                    Document oDoc = m_inventorApp.Documents.Open(fi.FullName, true);
                    oDoc.Activate();
                    //create the PDF in the file location *or should it be vendor folder?
                    exportPDF(false, true);
                    oDoc.Close(true);
                }
                catch (System.Exception ex)
                {
                    //error log for failed files
                    StringBuilder logOut = new StringBuilder();
                    logOut.AppendLine();
                    logOut.AppendFormat("{0} failed from error: {1}", fi.FullName, ex.ToString());
                    System.IO.File.AppendAllText(@"\\mgbwvlt\DATA2\CAD_Files\Upholstery\Frame PDF Error Log.txt", logOut.ToString());
                }
            }
            m_inventorApp.SilentOperation = false;
            //for duplicate names move old spec into a "old folder" rename with date from last modified (just delete and dont move if filename with date is duplicate)
        }


        //Adds a form of watermark to the chosen sheet
        #endregion
    }
    
}
