using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using ClosedXML.Excel;

public static class Tools
{
    private static readonly Regex Suffix = new Regex(@"^(.*?)-(\d+)$", RegexOptions.Compiled);

    private static (string Prefix, int Number, bool HasNumber) Key(string s)
    {
        var m = Suffix.Match(s ?? string.Empty);
        return m.Success
            ? (m.Groups[1].Value, int.Parse(m.Groups[2].Value), true)
            : (s ?? string.Empty, int.MaxValue, false);
    }

    private static int CompareKey((string Prefix, int Number, bool HasNumber) a,
                                  (string Prefix, int Number, bool HasNumber) b)
    {
        int c = string.Compare(a.Prefix, b.Prefix, StringComparison.OrdinalIgnoreCase);
        if (c != 0) return c;

        c = a.Number.CompareTo(b.Number);
        if (c != 0) return c;

        // If same prefix & number, ones that *have* a number come first
        return a.HasNumber == b.HasNumber ? 0 : (a.HasNumber ? -1 : 1);
    }

    public static List<T> SortByRowId<T>(IEnumerable<T> rows, Func<T, string> rowIdSelector)
    {
        return rows
            .OrderBy(r => Key(rowIdSelector(r)),
                     Comparer<(string Prefix, int Number, bool HasNumber)>.Create(CompareKey))
            .ToList();
    }


    public static List<string> FindDistributionBoardData(string filePath, string subMain)
    {
        using (var workbook = new XLWorkbook(filePath))
        {
            var ws = workbook.Worksheet("Distribution Board");
            var lastRow = ws.LastRowUsed().RowNumber();

            for (int i = 2; i <= lastRow; i++) // assuming row 1 is headers
            {
                string valG = ws.Cell(i, 7).GetString(); // Column G
                if (string.Equals(valG, subMain, StringComparison.OrdinalIgnoreCase))
                {
                    string colB = ws.Cell(i, 2).GetString();  // Column B
                    string colI = ws.Cell(i, 9).GetString(); // Column I

                    return new List<string> { colB, colI};
                }
            }
        }

        return new List<string>(); // return empty if not found
    }

    public static List<string> FindConsumerUnitData(string filePath, string subMain)
    {
        using (var workbook = new XLWorkbook(filePath))
        {
            var ws = workbook.Worksheet("Consumer Unit");
            var lastRow = ws.LastRowUsed().RowNumber();

            for (int i = 2; i <= lastRow; i++) // assuming row 1 is headers
            {
                string valG = ws.Cell(i, 7).GetString(); // Column G
                if (string.Equals(valG, subMain, StringComparison.OrdinalIgnoreCase))
                {
                    string colB = ws.Cell(i, 2).GetString();  // Column B
                    string colQ = ws.Cell(i, 17).GetString(); // Column Q
                    string colU = ws.Cell(i, 21).GetString(); // Column U

                    return new List<string> { colB, colQ, colU };
                }
            }
        }

        return new List<string>(); // return empty if not found
    }



    /// <summary>
    /// Returns one of: "Distribution Board", "Consumer Unit", "FAP", or "Unknown".
    /// Looks up X in:
    ///   - "Distribution Board" sheet, Column G
    ///   - "Consumer Unit"       sheet, Column G
    ///   - "Load"                sheet, Column D   (for FAP match)
    /// Matching is case-insensitive and trimmed. Row 1 is treated as header.
    /// </summary>
    public static string GetSWITCHType(string xlsPath, string X)
    {
        if (string.IsNullOrWhiteSpace(xlsPath) || string.IsNullOrWhiteSpace(X))
            return "Unknown";

        string candidate = X.Trim();

        try
        {
            using var wb = new XLWorkbook(xlsPath);

            var distBoards = ReadColumnAsSet(wb, "Distribution Board", "G");
            if (distBoards.Contains(candidate))
                return "Distribution Board";

            var consumerUnits = ReadColumnAsSet(wb, "Consumer Unit", "G");
            if (consumerUnits.Contains(candidate))
                return "Consumer Unit";

            // FAP: presence of X in Load!D and matching column A is "FAP"
            if (wb.TryGetWorksheet("Load", out var loadWs))
            {
                var lastRow = loadWs.LastRowUsed()?.RowNumber() ?? 0;

                for (int r = 2; r <= lastRow; r++) // skip header row
                {
                    string colD = loadWs.Cell(r, "D").GetString().Trim();
                    string colA = loadWs.Cell(r, "B").GetString().Trim();

                    if (string.Equals(colD, candidate, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(colA, "FAP", StringComparison.OrdinalIgnoreCase))
                    {
                        return "FAP";
                    }
                }
            }

            if (wb.TryGetWorksheet("Load", out var loadWs2))
            {
                var lastRow = loadWs2.LastRowUsed()?.RowNumber() ?? 0;

                for (int r = 2; r <= lastRow; r++) // skip header row
                {
                    string colD = loadWs2.Cell(r, "D").GetString().Trim();
                    string colF = loadWs2.Cell(r, "F").GetString().Trim();

                    if (string.Equals(colD, candidate, StringComparison.OrdinalIgnoreCase) )
                    {
                        if (colF.IndexOf("three phase", StringComparison.OrdinalIgnoreCase) >= 0)
                            return "Load3";
                        else if (colF.IndexOf("three phase", StringComparison.OrdinalIgnoreCase) >= 0)
                            return "Load1";
                    }
                }
            }


            return "Unknown";
        }
        catch
        {
            // If workbook/sheets/columns are missing or unreadable, just return Unknown.
            return "Unknown";
        }
    }

    /// <summary>
    /// Reads a column from a worksheet into a HashSet<string> (case-insensitive),
    /// skipping blanks and the header row (row 1).
    /// </summary>
    private static HashSet<string> ReadColumnAsSet(XLWorkbook wb, string sheetName, string colLetter)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (!wb.TryGetWorksheet(sheetName, out var ws))
            return set;

        var col = ws.Column(colLetter);
        foreach (var cell in col.CellsUsed())
        {
            // Skip header row (assumed row 1). Adjust if your header is elsewhere.
            if (cell.Address.RowNumber <= 1) continue;

            var val = cell.GetString()?.Trim();
            if (!string.IsNullOrEmpty(val))
                set.Add(val);
        }

        return set;
    }







}
