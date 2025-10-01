using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Vml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary1
{
    public class Network
    {
        [CommandMethod("JJ-NetworkPanel")]

        public static void NetworkPanels()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            // 1) Ask for the Excel file
            var pfo = new PromptOpenFileOptions("\nSelect Excel file (.xlsx / .xlsm)");
            pfo.Filter = "Excel Workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*";
            var pr = ed.GetFileNameForOpen(pfo);
            if (pr.Status != PromptStatus.OK) return;
            string xlsxPath = pr.StringResult;

            using var wb = new XLWorkbook(xlsxPath);
            var ws = wb.Worksheet("Source");

            
            int firstDataRow = 10;
            int lastRow = ws.LastRowUsed()?.RowNumber() ?? firstDataRow - 1;
            if (lastRow < firstDataRow) return; // nothing to do

            for (int r = firstDataRow; r <= lastRow; r++)
            {
                string idNo = ws.Cell(r, 2).GetString().Trim(); // Column B = 2
                if (string.IsNullOrWhiteSpace(idNo)) continue;

                ProcessAsource(idNo, xlsxPath, wb, ws, doc); // your existing function
            }
        }

        // Helper: first matching row (case-insensitive), starting at startRow
        private static int FindFirstRow(IXLWorksheet ws, int col, string target, int startRow)
        {
            if (string.IsNullOrWhiteSpace(target)) return -1;
            target = target.Trim();
            var used = ws.RangeUsed();
            if (used is null) return -1;

            int lastRow = ws.LastRowUsed().RowNumber();
            for (int r = Math.Max(startRow, used.FirstRow().RowNumber()); r <= lastRow; r++)
            {
                var val = ws.Cell(r, col).GetString().Trim();
                if (val.Equals(target, StringComparison.OrdinalIgnoreCase))
                    return r;
            }
            return -1;
        }

        // Helper: all matching rows (Column col equals target), starting at startRow
        private static IEnumerable<int> FindAllRows(IXLWorksheet ws, int col, string target, int startRow)
        {
            if (string.IsNullOrWhiteSpace(target)) yield break;
            target = target.Trim();
            var used = ws.RangeUsed();
            if (used is null) yield break;

            int lastRow = ws.LastRowUsed().RowNumber();
            for (int r = Math.Max(startRow, used.FirstRow().RowNumber()); r <= lastRow; r++)
            {
                var val = ws.Cell(r, col).GetString().Trim();
                if (val.Equals(target, StringComparison.OrdinalIgnoreCase))
                    yield return r;
            }
        }

        /// <summary>
        /// Process a single Id (idNo).
        /// Current sheet (ws): find idNo in Column B (from row 10), read Column F => ConnectedTo.
        /// Cable sheet: find ConnectedTo in Column B; from same row read Column E => ConnectedToCable.
        /// Cable sheet: find all rows where Column D == ConnectedToCable; collect Column B values => ListConnectedFrom.
        /// Print each X in ListConnectedFrom via ed.WriteMessage.
        /// </summary>
        public static void ProcessAsource(string idNo, string xlsxPath, XLWorkbook wb, IXLWorksheet ws, Document doc)
        {

            var ed = doc.Editor;

            if (string.IsNullOrWhiteSpace(idNo)) return;

            // 1) Current sheet: find idNo in Column B (from row 10)
            int rowOnCurrent = FindFirstRow(ws, col: 2, target: idNo, startRow: 10);
            if (rowOnCurrent < 0)
            {
                ed.WriteMessage($"\n[id {idNo}] Not found on current sheet Column B.");
                return;
            }

            // Column F => ConnectedTo
            string ConnectedTo = ws.Cell(rowOnCurrent, 6).GetString().Trim();
            if (string.IsNullOrWhiteSpace(ConnectedTo))
            {
                ed.WriteMessage($"\n[id {idNo}] ConnectedTo (Column F) is empty.");
                return;
            }

            // 2) Cable sheet
            var wsCable = wb.Worksheet("Cable");
            if (wsCable == null)
            {
                ed.WriteMessage("\nSheet 'Cable' not found.");
                return;
            }

            // Find ConnectedTo in Cable!B
            int rowCableForConnectedTo = FindFirstRow(wsCable, col: 2, target: ConnectedTo, startRow: 1);
            if (rowCableForConnectedTo < 0)
            {
                ed.WriteMessage($"\n[id {idNo}] ConnectedTo '{ConnectedTo}' not found in Cable!B.");
                return;
            }

            // CableID (Cable!B on that row) – kept if needed
            string CableID = wsCable.Cell(rowCableForConnectedTo, 2).GetString().Trim();

            // Cable!E on same row => ConnectedToCable
            string ConnectedToCable = wsCable.Cell(rowCableForConnectedTo, 5).GetString().Trim();
            if (string.IsNullOrWhiteSpace(ConnectedToCable))
            {
                ed.WriteMessage($"\n[id {idNo}] ConnectedToCable (Cable!E) is empty for ConnectedTo '{ConnectedTo}'.");
                return;
            }

            // 3) All rows where Cable!D == ConnectedToCable
            var matchingRows = FindAllRows(wsCable, col: 4, target: ConnectedToCable, startRow: 1).ToList();
            if (matchingRows.Count == 0)
            {
                ed.WriteMessage($"\n[id {idNo}] No entries in Cable!D matching '{ConnectedToCable}'.");
                return;
            }

            ed.WriteMessage($"\n[id {idNo}] ConnectedTo='{ConnectedTo}', ConnectedToCable='{ConnectedToCable}'. Pairs: {matchingRows.Count}");

            // insert that board first

            
            Tools.ProcessAswitchBoard(ConnectedToCable, xlsxPath, doc);


            // Print as: --> <Cable!B> ---> <Cable!E>
            foreach (var r in matchingRows)
            {
                var fromVal = wsCable.Cell(r, 2).GetString().Trim(); // Cable!B
                var toVal = wsCable.Cell(r, 5).GetString().Trim(); // Cable!E
                if (string.IsNullOrWhiteSpace(fromVal) && string.IsNullOrWhiteSpace(toVal)) continue;

                string toValType = NetworkTools.ClassifyId(wb,toVal);

                ed.WriteMessage($"\n--> {fromVal} ---> {toVal} ==> Type of {toVal} is: {toValType}");

                if (string.Equals(toValType, "Switchboard", StringComparison.OrdinalIgnoreCase))
                {
                    ed.WriteMessage($"\n-------> {toVal} is a : Switchboard ==> Run previous code (stage 1 - 2)");


                    Tools.ProcessAswitchBoard(toVal, xlsxPath, doc);

                }// 


            }

        }





    }
}
