// (C) Copyright 2023 by Autodesk, Inc. 
//
// Permission to use, copy, modify, and distribute this software
// in object code form for any purpose and without fee is hereby
// granted, provided that the above copyright notice appears in
// all copies and that both that copyright notice and the limited
// warranty and restricted rights notice below appear in all
// supporting documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK,
// INC. DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL
// BE UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is
// subject to restrictions set forth in FAR 52.227-19 (Commercial
// Computer Software - Restricted Rights) and DFAR 252.227-7013(c)
// (1)(ii)(Rights in Technical Data and Computer Software), as
// applicable.
//

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using DesignAutomationFramework;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Revit.IFC.Import;
using Revit.IFC.Import.Data;
using Revit.IFC.Import.Utility;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace RevitIfcLinkPropDumper
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalDBApplication
    {
        public ExternalDBApplicationResult OnStartup(ControlledApplication application)
        {
            DesignAutomationBridge.DesignAutomationReadyEvent += HandleDesignAutomationReadyEvent;
            return ExternalDBApplicationResult.Succeeded;
        }

        public ExternalDBApplicationResult OnShutdown(ControlledApplication application)
        {
            return ExternalDBApplicationResult.Succeeded;
        }

        [Obsolete]
        public void HandleApplicationInitializedEvent(object sender, Autodesk.Revit.DB.Events.ApplicationInitializedEventArgs e)
        {
            var app = sender as Autodesk.Revit.ApplicationServices.Application;
            DesignAutomationData data = new DesignAutomationData(app, "InputFile.rvt");
            this.DoTask(data);
        }

        private void HandleDesignAutomationReadyEvent(object sender, DesignAutomationReadyEventArgs e)
        {
            LogTrace("Design Automation Ready event triggered ...");
            // Hook up the CustomFailureHandling failure processor.
            Application.RegisterFailuresProcessor(new ExportIfcFailuresProcessor());

            e.Succeeded = true;
            e.Succeeded = this.DoTask(e.DesignAutomationData);
        }

        private bool DoTask(DesignAutomationData data)
        {
            if (data == null)
                return false;

            Application app = data.RevitApp;
            if (app == null)
            {
                LogTrace("Error occurred");
                LogTrace("Invalid Revit App");
                return false;
            }

            LogTrace("Creating an empty `host.rvt` ...");
            var hostDoc = app.NewProjectDocument(UnitSystem.Metric);
            if (hostDoc == null)
            {
                LogTrace("Error occurred");
                LogTrace("Invalid Revit DB Document");
                return false;
            }
            LogTrace(" - DONE.");

            LogTrace("Linking IFC ...");
            var folder = new DirectoryInfo(Directory.GetCurrentDirectory());
            FileInfo[] ifcLinkFiles = folder.GetFiles("*.ifc", SearchOption.AllDirectories);

            if (ifcLinkFiles.Length <= 0)
            {
                LogTrace("Error occurred");
                LogTrace("No IFC found to be linked");
                return false;
            }

            IDictionary<string, string> options = new Dictionary<string, string>();
            options["Action"] = "Link";   // default is Open.
            options["Intent"] = "Reference"; // This is the default.

            foreach (FileInfo ifcFile in ifcLinkFiles)
            {
                var ifcName = ifcFile.FullName;
                LogTrace($" - Linking `{ifcName}` ...");

                try
                {
                    // Clear the maps at the start of import, to force reload of options.
                    IFCCategoryUtil.Clear();

                    string fullIFCFileName = ifcFile.FullName; //Path.Combine(outputPath, ifcName);
                    Importer importer = Importer.CreateImporter(hostDoc, fullIFCFileName, options);

                    importer.ReferenceIFC(hostDoc, fullIFCFileName, options);
                }
                catch (Exception ex)
                {
                    LogTrace("Exception in linking IFC document. " + ex.Message);
                    if (Importer.TheLog != null)
                        Importer.TheLog.LogError(-1, ex.Message, false);

                    return false;
                }
                finally
                {
                    if (Importer.TheLog != null)
                        Importer.TheLog.Close();

                    if (IFCImportFile.TheFile != null)
                        IFCImportFile.TheFile.Close();
                }

                LogTrace(" -- DONE.");
            }

            LogTrace("Saving `host.rvt` ...");
            ModelPath hostPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(Path.Combine(Directory.GetCurrentDirectory(), "host.rvt"));
            var saveHostOptions = new SaveAsOptions();
            saveHostOptions.OverwriteExistingFile = true;
            hostDoc.SaveAs(hostPath, saveHostOptions);
            hostDoc.Close(false);
            LogTrace(" - DONE.");

            LogTrace("IFC link completed ...");

            FileInfo[] ifcLinkRvtFiles = folder.GetFiles("*.ifc.RVT", SearchOption.AllDirectories);

            if (ifcLinkRvtFiles.Length <= 0)
            {
                LogTrace("Error occurred");
                LogTrace("No IFC found to be linked");
                return false;
            }

            LogTrace($"Exporting wall props from ifc ...");

            IWorkbook workbook = new XSSFWorkbook();

            foreach (FileInfo ifcLinkRvtFile in ifcLinkRvtFiles)
            {
                var ifcLinkRvtName = ifcLinkRvtFile.FullName;
                var sheeName = ifcLinkRvtFile.Name.Substring(0, 15);
                LogTrace($" - Opening `{ifcLinkRvtName}` ...");

                try
                {
                    ModelPath ifcLinkRvtPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(ifcLinkRvtName);
                    var ifcDoc = app.OpenDocumentFile(ifcLinkRvtPath, new OpenOptions());
                    LogTrace(" -- DONE.");

                    LogTrace($" - Creating Excel sheet for `{ifcLinkRvtName}` ...");
                    int rowIndex = 0;
                    ISheet workSheet = workbook.CreateSheet(sheeName);
                    workSheet.CreateRow(rowIndex);

                    workSheet.GetRow(rowIndex).CreateCell(0).SetCellValue("ElementId");
                    workSheet.GetRow(rowIndex).CreateCell(1).SetCellValue("Name");
                    workSheet.GetRow(rowIndex).CreateCell(2).SetCellValue("IFC Guid");
                    workSheet.GetRow(rowIndex).CreateCell(3).SetCellValue("Category");
                    workSheet.GetRow(rowIndex).CreateCell(4).SetCellValue("Element Type");
                    LogTrace(" -- DONE.");

                    using (var elemCollector = new FilteredElementCollector(ifcDoc).WhereElementIsNotElementType().OfClass(typeof(DirectShape)))
                    {
                        LogTrace($" - Collecting IFC props ...");
                        if(elemCollector.Count() <= 0)
                        {
                            LogTrace($" -- No IFC objects found ...");
                            continue;
                        }

                        foreach (Element elem in elemCollector)
                        {
                            var cate = elem.Category;
                            var elemType = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);
                            var name = elem.Name;
                            var id = elem.Id;
                            var ifcGuid = elem.get_Parameter(BuiltInParameter.IFC_GUID);

                            workSheet.CreateRow(++rowIndex);

                            workSheet.GetRow(rowIndex).CreateCell(0).SetCellValue(id.ToString());
                            workSheet.GetRow(rowIndex).CreateCell(1).SetCellValue(name);
                            workSheet.GetRow(rowIndex).CreateCell(2).SetCellValue(ifcGuid?.AsString());
                            workSheet.GetRow(rowIndex).CreateCell(3).SetCellValue(cate?.Name);
                            workSheet.GetRow(rowIndex).CreateCell(4).SetCellValue(elemType?.AsString());
                        }
                        LogTrace(" -- DONE.");
                    }

                    ifcDoc.Close(false);
                }
                catch (Exception ex)
                {
                    LogTrace("Exception in export props from IFC document. " + ex.Message);

                    return false;
                }
            }

            try
            {
                using (FileStream stream = new FileStream("result.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(stream, true);
                    workbook.Close();
                }
            }
            catch (Exception ex)
            {
                LogTrace("Exception in saving Excel result. " + ex.Message);
                return false;
            }

            return true;
        }

        /// <summary>
        /// This will appear on the Design Automation output
        /// </summary>
        private static void LogTrace(string format, params object[] args)
        {
#if DEBUG
            System.Diagnostics.Trace.WriteLine(string.Format(format, args));
#endif
            System.Console.WriteLine(format, args);
        }

    }
}
