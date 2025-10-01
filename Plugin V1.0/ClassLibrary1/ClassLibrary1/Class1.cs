using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using System.Text.RegularExpressions;

namespace ClassLibrary1
{
    public class CableCommands
    {
        [CommandMethod("JJ-stage1-2")]
        public void TestCableRowsForSwitchboard()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            var ed = doc.Editor;

            try
            {
                // 1) Ask for the Excel file
                var pfo = new PromptOpenFileOptions("\nSelect Excel file (.xlsx / .xlsm)");
                pfo.Filter = "Excel Workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*";
                var pr = ed.GetFileNameForOpen(pfo);
                if (pr.Status != PromptStatus.OK) return;
                string xlsxPath = pr.StringResult;

                // 2) Open workbook with ClosedXML
                List<string> ids = GetSwitchboardIds(xlsxPath, ed);

                // 3) Report results
                if (ids.Count == 0)
                {
                    ed.WriteMessage("\nNo Id No. values found  in Switchboard! (Column B starting at B10).");
                }
                else
                {
                    ed.WriteMessage($"\nFound {ids.Count} Id No. value(s) in Switchboard (B10↓):");
                    foreach (var id in ids)
                        ed.WriteMessage($"\n  - {id}");
                }


                foreach (var switchboardId in ids)
                {
                    Tools.ProcessAswitchBoard(switchboardId, xlsxPath, doc);
                }   // foreach (var switchboardId in ids)

                

            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }

        private static List<string> GetSwitchboardIds(string workbookPath, Autodesk.AutoCAD.EditorInput.Editor ed = null)
        {
            var result = new List<string>();

            using (var wb = new XLWorkbook(workbookPath))
            {
                // Find the "Switchboard" sheet (case-insensitive)
                IXLWorksheet ws = null;
                foreach (var sh in wb.Worksheets)
                {
                    if (string.Equals(sh.Name?.Trim(), "Switchboard", StringComparison.OrdinalIgnoreCase))
                    {
                        ws = sh;
                        break;
                    }
                }

                if (ws == null)
                {
                    ed?.WriteMessage("\nSheet 'Switchboard' not found.");
                    return result;
                }

                // Column B, starting at row 10
                int row = 10;
                while (true)
                {
                    var cell = ws.Cell(row, 2); // Column B = 2
                    string val = cell.GetString()?.Trim();

                    if (string.IsNullOrWhiteSpace(val))
                        break; // stop at the first blank (adjust if you need to skip sporadic blanks)

                    result.Add(val);
                    row++;
                }
            }

            return result;
        }
    }


}
