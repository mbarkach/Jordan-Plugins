using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using ClosedXML.Excel;
using System.Text;

namespace ClassLibrary1
{
    public class CableCommands
    {
        [CommandMethod("TESTCABLESB")]
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
                    string FieldExpress = "%<\\\\AcExpr ";

                    StringBuilder sb = new StringBuilder();
                    sb.Append("%<\\AcExpr ");
                    List<string> fieldParts = new List<string>();

                    // 3) Call the helper
                    var rows = CableLookup.GetRowsForSwitchboardId(xlsxPath, switchboardId);
                    rows = rows.OrderBy(r => r.RowId).ToList();

                    ed.WriteMessage($"\nFound {rows.Count} row(s) where 'Connected From' = '{switchboardId}':");
                    
                    int n = rows.Count;
                    int boardWidth = Math.Max(8000, (n * 1100) + 2200);

                    // 1) Ask user to pick insertion point
                    PromptPointOptions ppo = new PromptPointOptions("\nPick insertion point for board:");
                    PromptPointResult ppr = ed.GetPoint(ppo);
                    if (ppr.Status != PromptStatus.OK) return;

                    Point3d insPt = ppr.Value;

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        // 2) Ensure block "board" exists
                        if (!bt.Has("board"))
                        {
                            ed.WriteMessage("\nBlock 'board' not found in drawing.");
                            return;
                        }

                        // 3) Create reference
                        BlockReference br = new BlockReference(insPt, bt["board"]);

                        ms.AppendEntity(br);
                        tr.AddNewlyCreatedDBObject(br, true);

                        if (br.IsDynamicBlock)
                        {
                            //ed.WriteMessage($"\nBlock is dynamic");

                            foreach (DynamicBlockReferenceProperty prop in br.DynamicBlockReferencePropertyCollection)
                            {
                                if (string.Equals(prop.PropertyName, "Distance1", StringComparison.OrdinalIgnoreCase))
                                {
                                    prop.Value = (double)boardWidth;    // .ToString();
                                    break;
                                }
                            }
                        }

                        // Place Switches_X blocks:
                        double dx = boardWidth/2;
                        double dy = 153.0;
                        const string switchBlockName = "Switches_x";

                        Point3d pmidl = new Point3d(insPt.X + dx, insPt.Y + dy, insPt.Z);
                        // Insert the block at pmidl
                        BlockReference swRefIncom = new BlockReference(pmidl, bt[switchBlockName])
                        {
                            ScaleFactors = new Scale3d(40.5)
                        };

                        if (swRefIncom.IsDynamicBlock)
                        {
                            foreach (DynamicBlockReferenceProperty prop in swRefIncom.DynamicBlockReferencePropertyCollection)
                            {
                                if (string.Equals(prop.PropertyName, "Visibility1", StringComparison.OrdinalIgnoreCase))
                                {
                                    prop.Value = "Switch-Disconnector";    // .ToString();
                                    break;
                                }
                            }
                        }

                        ms.AppendEntity(swRefIncom);
                        tr.AddNewlyCreatedDBObject(swRefIncom, true);

                        // Place Switches_X blocks:
                        dx = 716.8;
                        dy = 1510.1;

                        int countRow = rows.Count;
                        foreach (var row in rows)
                        {
                            // Compute P1 = (X0 + 716.8, Y0 + 1510.1)
                            Point3d p1 = new Point3d(insPt.X + dx, insPt.Y + dy, insPt.Z);

                            // Ensure "Switches_x" block exists
                            
                            if (!bt.Has(switchBlockName))
                            {
                                ed.WriteMessage($"\nBlock '{switchBlockName}' not found in drawing.");
                                return;
                            }

                            // Insert the block at P1
                            BlockReference swRef = new BlockReference(p1, bt[switchBlockName])
                            {
                                ScaleFactors = new Scale3d(40.5)   // uniform scaling on X, Y, Z
                            };

                            if (swRef.IsDynamicBlock)
                            {
                                //ed.WriteMessage($"\nBlock is dynamic");

                                foreach (DynamicBlockReferenceProperty prop in swRef.DynamicBlockReferencePropertyCollection)
                                {
                                    if (string.Equals(prop.PropertyName, "Visibility1", StringComparison.OrdinalIgnoreCase))
                                    {
                                        prop.Value = "Isolator";    // .ToString();
                                        break;
                                    }
                                }
                            }

                            ms.AppendEntity(swRef);
                            tr.AddNewlyCreatedDBObject(swRef, true);

                            // This is for Fiedl + 
                            List<string> msvdTerms = new List<string>();

                            // Fill attributes from row.Values (keys are your exact attribute TAGs)
                            BlockTableRecord defBtr = (BlockTableRecord)tr.GetObject(bt[switchBlockName], OpenMode.ForRead);

                            // BEFORE your foreach loop (once per board)
                            List<ObjectId> msvdFieldIds = new List<ObjectId>();

                            foreach (ObjectId entId in defBtr)
                            {
                                var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                                if (ent is AttributeDefinition attDef && !attDef.Constant)
                                {
                                    // Create an AttributeReference based on the ATTDEF
                                    var ar = new AttributeReference();
                                    ar.SetAttributeFromBlock(attDef, swRef.BlockTransform);

                                    // Look up value by TAG (case-insensitive because your dictionary uses OrdinalIgnoreCase)
                                    if (!row.Values.TryGetValue(attDef.Tag, out string val) || val == null)
                                        val = string.Empty;

                                    ar.TextString = val;

                                    // Attach attribute to the inserted block reference
                                    swRef.AttributeCollection.AppendAttribute(ar);
                                    tr.AddNewlyCreatedDBObject(ar, true);


                                    // Inside your foreach over attributes of each Switches_x
                                    if (string.Equals(attDef.Tag, "MSVD-Max-Diversified-Load-(A)", StringComparison.OrdinalIgnoreCase))
                                    {
                                        long objId = ar.ObjectId.OldIdPtr.ToInt64();

                                        string fieldPart = $"%<\\AcObjProp Object(%<\\_ObjId {objId}>%).Textstring>%";
                                        fieldParts.Add(fieldPart);

                                    }
                                }
                            }

                            ed.Command("_.ATTSYNC", "_N", "board");
                            ed.Command("_.REGEN");

                            dx = dx + 1100;
                            countRow = countRow - 1;
                        }

                        sb.Append(string.Join(" + ", fieldParts));
                        sb.Append(" \\f \"%lu2%pr3\">%");
                        string NewFieldExpression = sb.ToString();

 
                        foreach (ObjectId attId in br.AttributeCollection) // 'br' is your board BlockReference
                        {
                            var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                            if (attRef != null && string.Equals(attRef.Tag, "REF", StringComparison.OrdinalIgnoreCase))
                            {
                                attRef.TextString = switchboardId;
                            }

                            if (attRef != null && string.Equals(attRef.Tag, "TOTAL-AMPS", StringComparison.OrdinalIgnoreCase))
                            {
                                attRef.TextString = NewFieldExpression;
                            }
                        }

                        ed.Command("_.ATTSYNC", "_N", "board");
                        ed.Command("_.REGEN");

                        tr.Commit();
                        ed.WriteMessage($"\n Full expression === > {FieldExpress}");

                    }
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

    public static class CableLookup
    {
        private static string? MapHeaderToTag(string header)
        {
            switch (header.Trim().ToUpperInvariant())
            {
                case "ID NO.":
                case "ID. NO":
                case "ID NO":
                    return "Id.-No";

                case "LENGTH (M)":
                case "LENGTH":
                    return "LENGTH-(M)";

                case "BREAKING CAPACITY (KA)":
                    return "Breaking-Capacity-(kA)";

                case "MSVD MAX DIVERSIFIED LOAD (A)":
                    return "MSVD-Max-Diversified-Load-(A)";

                case "DEVICE TYPE":
                    return "Device-Type";

                case "CABLE TYPE":
                    return "Cable-Type";

                case "CABLE MAKE UP":
                    return "CABLE-MAKE-UP";

                case "TPN SPN":
                    return "TPN-SPN";

                case "RATING (A)":
                    return "Rating-(A)";

                case "OVERLOAD SETTING":
                    return "Overload-Setting";

                case "CPC Description":
                    return "CPC-Description";

                default:
                    return null; // fallback: keep original if unmapped
            }
        }

        public static List<CableRow> GetRowsForSwitchboardId(string workbookPath, string switchboardId)
        {
            //var doc = Application.DocumentManager.MdiActiveDocument;
            //var ed = doc.Editor;

            using var wb = new XLWorkbook(workbookPath);
            var ws = wb.Worksheets.FirstOrDefault(s => string.Equals(s.Name?.Trim(), "Cable", StringComparison.OrdinalIgnoreCase))
                     ?? throw new InvalidOperationException("Sheet 'Cable' not found.");

            var used = ws.RangeUsed();
            if (used == null) return new List<CableRow>();

            int headerRow = 9;
            int firstDataRow = headerRow + 1;
            int lastRow = used.LastRow().RowNumber();

            // Read headers
            var headers = new List<string>();
            for (int c = used.FirstColumn().ColumnNumber(); c <= used.LastColumn().ColumnNumber(); c++)
                headers.Add(ws.Cell(headerRow, c).GetString().Trim());

            int colD = 4;
            var results = new List<CableRow>();
            string IdNo = "";
            for (int r = firstDataRow; r <= lastRow; r++)
            {
                var val = ws.Cell(r, colD).GetString()?.Trim();
                if (string.Equals(val, switchboardId.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    
                    var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    for (int c = used.FirstColumn().ColumnNumber(); c <= used.LastColumn().ColumnNumber(); c++)
                    {
                        string header = headers[c - used.FirstColumn().ColumnNumber()];
                        string v = ws.Cell(r, c).GetString();
                        string? tag = MapHeaderToTag(header);
                        if (!string.IsNullOrEmpty(tag))   // only add if mapped
                        { 
                            dict[tag] = v;
                            if (string.Equals(tag.Trim(), "Id.-No", StringComparison.OrdinalIgnoreCase))
                                IdNo = v; 

                        }
                            
                    }

                    results.Add(new CableRow { RowId = IdNo, Values = dict });
                }
            }
            return results;
        }
    }

    public sealed class CableRow
    {
        public string RowId { get; set; }
        public Dictionary<string, string> Values { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

}
