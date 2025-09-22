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

                    // 3) Call the GetRowsForSwitchboardId as Object CablRow
                    var rowsObjects = CableLookup.GetRowsForSwitchboardId(xlsxPath, switchboardId);

                    var rows = Tools.SortByRowId(rowsObjects, r => r.RowId);

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

                        // 2) Ensure block "board" definition exists
                        if (!bt.Has("board"))
                        {
                            ed.WriteMessage("\nBlock 'board' not found in drawing.");
                            return;
                        }

                        // 3) Create a reference block (a Board Ref block)
                        BlockReference BoardBlockRef = new BlockReference(insPt, bt["board"]);

                        ms.AppendEntity(BoardBlockRef);
                        tr.AddNewlyCreatedDBObject(BoardBlockRef, true);

                        if (BoardBlockRef.IsDynamicBlock)
                        {
                            // Set Dynamic Prop Distance1 to ...
                            foreach (DynamicBlockReferenceProperty prop in BoardBlockRef.DynamicBlockReferencePropertyCollection)
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

                        // Insert the block ref SWITCHE_s at pmidl
                        BlockReference SwitchBlockRef = new BlockReference(pmidl, bt[switchBlockName])
                        {
                            ScaleFactors = new Scale3d(40.5)
                        };

                        ms.AppendEntity(SwitchBlockRef);
                        tr.AddNewlyCreatedDBObject(SwitchBlockRef, true);

                        if (SwitchBlockRef.IsDynamicBlock)    // set Visibility1
                        {
                            foreach (DynamicBlockReferenceProperty prop in SwitchBlockRef.DynamicBlockReferencePropertyCollection)
                            {
                                if (string.Equals(prop.PropertyName, "Visibility1", StringComparison.OrdinalIgnoreCase))
                                {
                                    prop.Value = "Switch-Disconnector";    // .ToString();
                                    break;
                                }
                            }
                        }

                        ed.Command("_.ATTSYNC", "_N", "Switches_x", "_Y");

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
                            BlockReference aSwitchBlockRef = new BlockReference(p1, bt[switchBlockName])
                            {
                                ScaleFactors = new Scale3d(40.5)   // uniform scaling on X, Y, Z
                            };

                            ms.AppendEntity(aSwitchBlockRef);
                            tr.AddNewlyCreatedDBObject(aSwitchBlockRef, true);

                            if (aSwitchBlockRef.IsDynamicBlock)
                            {
                                foreach (DynamicBlockReferenceProperty prop in aSwitchBlockRef.DynamicBlockReferencePropertyCollection)
                                {
                                    if (string.Equals(prop.PropertyName, "Visibility1", StringComparison.OrdinalIgnoreCase))
                                    {
                                        prop.Value = "Isolator";    // .ToString();
                                        break;
                                    }
                                }
                            }

                            

                            // This is for Fiedl + 
                            List<string> msvdTerms = new List<string>();

                            // Fill attributes from row.Values (keys are your exact attribute TAGs)
                            // We create AttrebutesDef from THE Block Definition (I said THE because it's 1 unique in dwg)

                            BlockTableRecord SwitchBlockDef = (BlockTableRecord)tr.GetObject(bt[switchBlockName], OpenMode.ForRead);

                            List<ObjectId> msvdFieldIds = new List<ObjectId>();

                            foreach (ObjectId entId in SwitchBlockDef)
                            {
                                var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                                if (ent is AttributeDefinition attDef && !attDef.Constant)
                                {
                                    // Create an AttributeReference based on the ATTDEF
                                    var ar = new AttributeReference();
                                    ar.SetAttributeFromBlock(attDef, aSwitchBlockRef.BlockTransform);

                                    // Look up value by TAG (case-insensitive because your dictionary uses OrdinalIgnoreCase)
                                    if (!row.Values.TryGetValue(attDef.Tag, out string val) || val == null)
                                        val = string.Empty;

                                    ar.TextString = val;

                                    // Attach attribute to the inserted block reference
                                    aSwitchBlockRef.AttributeCollection.AppendAttribute(ar);
                                    tr.AddNewlyCreatedDBObject(ar, true);


                                    // Inside your foreach over attributes of each Switches_x
                                    if (string.Equals(attDef.Tag, "MSVD-Max-Diversified-Load-(A)", StringComparison.OrdinalIgnoreCase))
                                    {
                                        long objId = ar.ObjectId.OldIdPtr.ToInt64();

                                        string fieldPart = $"%<\\AcObjProp Object(%<\\_ObjId {objId}>%).Textstring>%";
                                        fieldParts.Add(fieldPart);

                                    }
                                }
                            } // foreach (ObjectId entId in SwitchBlockDef)

                            // CHECK CONSUMER UNIT TAB. 
                            // CHECK CONNECTION(COLUMN G)
                            // IF SUB-MAIN IS CONNECTED, THEN SHOW THIS SYMBOL ON TOP, WHERE:
                            // // Id-NO = "DB"[COLUMN B]
                            // No of Circuits = "NO-OF-WAYS"[COLUMN Q + COLUMN U, WHERE U IS NOT EQUAL TO N / A]

                            string SwitchTypeName = Tools.GetSWITCHType(xlsxPath, row.RowId);

                            ed.WriteMessage($"\nType for this SWITCHE {row.RowId} row is : {Tools.GetSWITCHType(xlsxPath , row.RowId)}");

                            switch (SwitchTypeName)
                            {
                                case "Consumer Unit":
                                    List<string> SubMainConsumerUnitData = Tools.FindConsumerUnitData(xlsxPath, row.RowId);

                                    if (SubMainConsumerUnitData.Count > 0)
                                    {
                                        p1 = new Point3d(p1.X, p1.Y + 7234.2, insPt.Z);

                                        if (!bt.Has("Dist Board"))
                                        {
                                            ed.WriteMessage($"\nBlock Dist Board not found in drawing.");

                                        }

                                        // Insert the block Dist Boar at P1
                                        BlockReference DistBoardBlockRef = new BlockReference(p1, bt["Dist Board"])
                                        {
                                            ScaleFactors = new Scale3d(1.6)   // uniform scaling on X, Y, Z
                                        };

                                        ms.AppendEntity(DistBoardBlockRef);
                                        tr.AddNewlyCreatedDBObject(DistBoardBlockRef, true);

                                        // Fill attributes from row.Values (keys are your exact attribute TAGs)
                                        BlockTableRecord DistBoardBlockDef = (BlockTableRecord)tr.GetObject(bt["Dist Board"], OpenMode.ForRead);
                                        DistBoardBlockDef.UpdateAnonymousBlocks();

                                        // Set Visibility

                                        if (DistBoardBlockRef.IsDynamicBlock)
                                        {
                                            foreach (DynamicBlockReferenceProperty prop in DistBoardBlockRef.DynamicBlockReferencePropertyCollection)
                                            {
                                                if (string.Equals(prop.PropertyName, "Visibility1", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    prop.Value = "Consumer Unit";    // .ToString();
                                                    break;
                                                }
                                            }
                                        }

                                        // Set attributes
                                        foreach (ObjectId entId in DistBoardBlockDef)
                                        {
                                            var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                                            if (ent is AttributeDefinition attDef && !attDef.Constant)
                                            {
                                                // Create an AttributeReference based on the ATTDEF
                                                var ar = new AttributeReference();
                                                ar.SetAttributeFromBlock(attDef, DistBoardBlockRef.BlockTransform);


                                                // Set DO to "test"
                                                if (string.Equals(attDef.Tag, "DB", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    ar.TextString = SubMainConsumerUnitData[0];
                                                    ar.AdjustAlignment(DistBoardBlockRef.Database);
                                                }

                                                // Set DO to "test"
                                                if (SubMainConsumerUnitData[2] != "N/A" && string.Equals(attDef.Tag, "NO-OF-WAYS", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    ar.TextString = (float.Parse(SubMainConsumerUnitData[1]) + float.Parse(SubMainConsumerUnitData[2])).ToString();
                                                    ar.AdjustAlignment(DistBoardBlockRef.Database);
                                                }

                                                // Attach attribute to the inserted block reference
                                                DistBoardBlockRef.AttributeCollection.AppendAttribute(ar);
                                                tr.AddNewlyCreatedDBObject(ar, true);

                                            }
                                        }

                                        DistBoardBlockRef.RecordGraphicsModified(true);
                                        //ed.Regen();

                                        //ed.Command("_.ATTSYNC", "_N", "Dist Board");

                                    }

                                    break;

                                case "Distribution Board":
                                    List<string> SubMainDistributionBoardData = Tools.FindDistributionBoardData(xlsxPath, row.RowId);

                                    if (SubMainDistributionBoardData.Count > 0)
                                    {
                                        p1 = new Point3d(p1.X, p1.Y + 7234.2, insPt.Z);

                                        if (!bt.Has("Dist Board"))
                                        {
                                            ed.WriteMessage($"\nBlock Dist Board not found in drawing.");

                                        }

                                        // Insert the block Dist Boar at P1
                                        BlockReference DistBoardBlockRef = new BlockReference(p1, bt["Dist Board"])
                                        {
                                            ScaleFactors = new Scale3d(1.6)   // uniform scaling on X, Y, Z
                                        };

                                        ms.AppendEntity(DistBoardBlockRef);
                                        tr.AddNewlyCreatedDBObject(DistBoardBlockRef, true);

                                        // Fill attributes from row.Values (keys are your exact attribute TAGs)
                                        BlockTableRecord DistBoardBlockDef = (BlockTableRecord)tr.GetObject(bt["Dist Board"], OpenMode.ForRead);
                                        DistBoardBlockDef.UpdateAnonymousBlocks();

                                        // Set Visibility

                                        if (DistBoardBlockRef.IsDynamicBlock)
                                        {
                                            foreach (DynamicBlockReferenceProperty prop in DistBoardBlockRef.DynamicBlockReferencePropertyCollection)
                                            {
                                                if (string.Equals(prop.PropertyName, "Visibility1", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    prop.Value = "Distribution Board";    // .ToString();
                                                    break;
                                                }
                                            }
                                        }

                                        // Set attributes
                                        foreach (ObjectId entId in DistBoardBlockDef)
                                        {
                                            var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                                            if (ent is AttributeDefinition attDef && !attDef.Constant)
                                            {
                                                // Create an AttributeReference based on the ATTDEF
                                                var ar = new AttributeReference();
                                                ar.SetAttributeFromBlock(attDef, DistBoardBlockRef.BlockTransform);


                                                // Set DO to "test"
                                                if (string.Equals(attDef.Tag, "DB", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    ar.TextString = SubMainDistributionBoardData[0];
                                                    ar.AdjustAlignment(DistBoardBlockRef.Database);
                                                }

                                                // Set DO to "test"
                                                if (SubMainDistributionBoardData[1] != "N/A" && string.Equals(attDef.Tag, "NO-OF-WAYS", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    ar.TextString = SubMainDistributionBoardData[1];
                                                    ar.AdjustAlignment(DistBoardBlockRef.Database);
                                                }

                                                // Attach attribute to the inserted block reference
                                                DistBoardBlockRef.AttributeCollection.AppendAttribute(ar);
                                                tr.AddNewlyCreatedDBObject(ar, true);

                                            }
                                        }

                                        DistBoardBlockRef.RecordGraphicsModified(true);
                                        //ed.Regen();

                                        //ed.Command("_.ATTSYNC", "_N", "Dist Board");

                                    }

                                    break;

                                case "FAP":
                                    
                                    p1 = new Point3d(p1.X, p1.Y + 7234.2, insPt.Z);

                                        if (!bt.Has("FAP"))
                                        {
                                            ed.WriteMessage($"\nBlock FAP not found in drawing.");

                                        }

                                        // Insert the block Dist Boar at P1
                                        BlockReference FAPBlockRef = new BlockReference(p1, bt["FAP"])
                                        {
                                            ScaleFactors = new Scale3d(1)   // uniform scaling on X, Y, Z
                                        };

                                        ms.AppendEntity(FAPBlockRef);
                                        tr.AddNewlyCreatedDBObject(FAPBlockRef, true);

                                        // Fill attributes from row.Values (keys are your exact attribute TAGs)
                                        BlockTableRecord FAPBlockDef = (BlockTableRecord)tr.GetObject(bt["FAP"], OpenMode.ForRead);
                                        FAPBlockDef.UpdateAnonymousBlocks();

                                        FAPBlockRef.RecordGraphicsModified(true);

                                    ed.Command("_.ATTSYNC", "_N", "FAP");

                                    break;

                                case "Load1" or "Load3":

                                    p1 = new Point3d(p1.X, p1.Y + 7234.2, insPt.Z);

                                    if (!bt.Has("Isolator_Disconnector"))
                                    {
                                        ed.WriteMessage($"\nBlock Isolator_Disconnector not found in drawing.");

                                    }

                                    // Insert the block Dist Boar at P1
                                    BlockReference LoadBlockRef = new BlockReference(p1, bt["Isolator_Disconnector"])
                                    {
                                        ScaleFactors = new Scale3d(1)   // uniform scaling on X, Y, Z
                                    };

                                    ms.AppendEntity(LoadBlockRef);
                                    tr.AddNewlyCreatedDBObject(LoadBlockRef, true);

                                    // Fill attributes ...
                                    BlockTableRecord LaodBlockDef = (BlockTableRecord)tr.GetObject(bt["FAP"], OpenMode.ForRead);
                                    LaodBlockDef.UpdateAnonymousBlocks();

                                    // Set Visibility

                                    if (LoadBlockRef.IsDynamicBlock)
                                    {
                                        foreach (DynamicBlockReferenceProperty prop in LoadBlockRef.DynamicBlockReferencePropertyCollection)
                                        {
                                            if (string.Equals(prop.PropertyName, "Visibility1", StringComparison.OrdinalIgnoreCase))
                                            {
                                                if (string.Equals("Load3", SwitchTypeName, StringComparison.OrdinalIgnoreCase))
                                                    prop.Value = "3 PHASE";
                                                else if (string.Equals("Load1", SwitchTypeName, StringComparison.OrdinalIgnoreCase))
                                                    prop.Value = "SINGLE PHASE";

                                                break;
                                            }
                                        }

                                    }

                                    // Set attributes
                                    foreach (ObjectId entId in LaodBlockDef)
                                    {
                                        var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                                        if (ent is AttributeDefinition attDef && !attDef.Constant)
                                        {
                                            // Create an AttributeReference based on the ATTDEF
                                            var ar = new AttributeReference();
                                            ar.SetAttributeFromBlock(attDef, LoadBlockRef.BlockTransform);


                                            // Set BREAKER-SIZE to Rating-(A) from cable sheet:
                                            if (string.Equals(attDef.Tag, "BREAKER-SIZE", StringComparison.OrdinalIgnoreCase))
                                            {
                                                if (row.Values.TryGetValue("Rating-(A)", out var ratingText))
                                                    ar.TextString = ratingText;
                                                    ar.AdjustAlignment(LoadBlockRef.Database);
                                            }

                                            // Attach attribute to the inserted block reference
                                            LoadBlockRef.AttributeCollection.AppendAttribute(ar);
                                            tr.AddNewlyCreatedDBObject(ar, true);

                                        }
                                    }

                                    ed.Command("_.ATTSYNC", "_N", "Isolator_Disconnector", "_Y");

                                    break;

                            } // End swith cases for ... 

                            dx = dx + 1100;
                            countRow = countRow - 1;
                        } // for each (var row in rows)


                        sb.Append(string.Join(" + ", fieldParts));
                        sb.Append(" \\f \"%lu2%pr3\">%");
                        string NewFieldExpression = sb.ToString();

 
                        foreach (ObjectId attId in BoardBlockRef.AttributeCollection) // 'br' is your board BlockReference
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


                        tr.Commit();
                        // ed.WriteMessage($"\n Full expression === > {FieldExpress}");

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
