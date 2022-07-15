#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class cmdProjectSetup : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            string excelFile = @"C:\SourceXL\Excel\Session02_Challenge-220706-113155.xlsx";

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);
            Excel.Worksheet excelWs = excelWb.Worksheets.Item[1];
            Excel.Worksheet excelWs1 = excelWb.Worksheets.Item[2];
            Excel.Range excelRng = excelWs.UsedRange;
            Excel.Range excelRng1 = excelWs1.UsedRange;

            int rowCount = excelRng.Rows.Count;
            int rowCount1 = excelRng1.Rows.Count;
            List<string[]> levelList = new List<string[]>();
            List<string[]> sheetList = new List<string[]>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            collector.WhereElementIsElementType();
            var tblk = collector.FirstElementId();

            for (int i = 1; i<= rowCount; i++)
            {
                Excel.Range cell1 = excelWs.Cells[i, 1];
                Excel.Range cell2 = excelWs.Cells[i, 2];

                string data1 = cell1.Value.ToString();
                string data2 = cell2.Value.ToString();
                string[] dataArray = new string[2];
                dataArray[0] = data1;
                dataArray[1] = data2;

                levelList.Add(dataArray);
            }

            for (int i = 1; i<= rowCount; i++)
            {
                Excel.Range cell1 = excelWs.Cells[i, 1];
                Excel.Range cell2 = excelWs.Cells[i, 2];

                string data1 = cell1.Value.ToString();
                string data2 = cell2.Value.ToString();
                string[] dataArray = new string[2];
                dataArray[0] = data1;
                dataArray[1] = data2;

                sheetList.Add(dataArray);
            }

            levelList.RemoveAt(0);
            sheetList.RemoveAt(0);

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create some levels");

                foreach (string[] levelData in levelList)
                {
                    string levelName = levelData[0];
                    string levelFeetstr = levelData[1];

                    double levelFeet = Double.Parse(levelFeetstr);
                    Level curLevel = Level.Create(doc, levelFeet);
                    string tname = "Element " + curLevel.Id.ToString();
                    curLevel.Name = tname;
                    curLevel.Name = levelName;
                }

                t.Commit();
            }

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create some Sheets");

                foreach (string[] sheetData in sheetList)
                {
                    string sheetNum = sheetData[0];
                    string sheetName = sheetData[1];

                    ViewSheet curSheet = ViewSheet.Create(doc, tblk);
                    curSheet.SheetNumber = sheetNum;
                    curSheet.Name = sheetName;
                }

                t.Commit();

                excelWb.Close();
                excelApp.Quit();

                return Result.Succeeded;
            }
        }
    }
}