using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary1
{
    public class NetworkTools
    {
        // Returns one of: "Switchboard", "Distribution Board", "Consumer Unit", "Load3", "Load1", "Fap", or null if not found
        public static string ClassifyId(XLWorkbook wb, string idNo)
        {
            if (string.IsNullOrWhiteSpace(idNo)) return null;


            // 1) Switchboard!B (from row 10)
            var wsSw = wb.Worksheet("Switchboard");
            if (wsSw != null && FindFirstRow(wsSw, 2, idNo, 10) > 0)
                return "Switchboard";

            // 2) Distribution Board!B (from row 10)
            var wsDb = wb.Worksheet("Distribution Board");
            if (wsDb != null && FindFirstRow(wsDb, 2, idNo, 10) > 0)
                return "Distribution Board";

            // 3) Consumer Unit!B (from row 10)
            var wsCu = wb.Worksheet("Consumer Unit");
            if (wsCu != null && FindFirstRow(wsCu, 2, idNo, 10) > 0)
                return "Consumer Unit";

            // 4) Load!B (from row 10)
            var wsLoad = wb.Worksheet("Load");
            int r = (wsLoad == null) ? -1 : FindFirstRow(wsLoad, 2, idNo, 10);
            if (r > 0)
            {
                // Column E decides if it's a "Load" row vs FAP
                string colE = wsLoad.Cell(r, 5).GetString().Trim(); // Load!E
                if (colE.Equals("Load", StringComparison.OrdinalIgnoreCase))
                {
                    // Column F decides phase: "three phase" -> Load3, otherwise -> Load1
                    string colF = wsLoad.Cell(r, 6).GetString();
                    bool three = colF?.IndexOf("three phase", StringComparison.OrdinalIgnoreCase) >= 0;
                    return three ? "Load3" : "Load1";
                }
                else
                {
                    // Not "Load" in E => Fap
                    return "Fap";
                }
            }

            // Not found anywhere
            return null;
        }

        // Helper: first matching row (case-insensitive), starting at startRow
        private static int FindFirstRow(IXLWorksheet ws, int col, string target, int startRow)
        {
            var used = ws.RangeUsed();
            if (used == null || string.IsNullOrWhiteSpace(target)) return -1;

            int last = ws.LastRowUsed().RowNumber();
            string t = target.Trim();

            for (int r = Math.Max(startRow, used.FirstRow().RowNumber()); r <= last; r++)
            {
                if (ws.Cell(r, col).GetString().Trim()
                      .Equals(t, StringComparison.OrdinalIgnoreCase))
                    return r;
            }
            return -1;
        }
    }
}
