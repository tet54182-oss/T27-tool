using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

namespace Civil3DCsharp
{
    /// <summary>
    /// Commands for getting excavation and filling volume information in cross-section
    /// </summary>
    public class CTSV_KhoiLuongCatNgang_Commands
    {
        /// <summary>
        /// Command to get excavation and filling volume information from cross-section
        /// Input: Alignment → MaterialList → MaterialListItem → MaterialQuantity
        /// Output: Table format information showing volume data
        /// </summary>
        [CommandMethod("CTSV_KhoiLuongCatNgang")]
        public void CTSV_KhoiLuongCatNgang()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            try
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    // Prompt user to select an Alignment
                    var alignmentPrompt = new PromptEntityOptions("\nSelect an Alignment: ");
                    alignmentPrompt.SetRejectMessage("Selected object is not an Alignment.");
                    alignmentPrompt.AddAllowedClass(typeof(Alignment), true);
                    
                    var alignmentResult = ed.GetEntity(alignmentPrompt);
                    if (alignmentResult.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\nCommand cancelled.");
                        return;
                    }

                    var alignment = trans.GetObject(alignmentResult.ObjectId, OpenMode.ForRead) as Alignment;
                    if (alignment == null)
                    {
                        ed.WriteMessage("\nFailed to get Alignment object.");
                        return;
                    }

                    // Get Material Lists from the Alignment
                    var materialLists = GetMaterialLists(alignment, trans, ed);
                    if (materialLists.Count == 0)
                    {
                        ed.WriteMessage("\nNo Material Lists found for the selected Alignment.");
                        return;
                    }

                    // Create and display volume information table
                    var volumeData = ExtractVolumeData(materialLists, trans);
                    DisplayVolumeTable(volumeData, ed);

                    trans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }

        /// <summary>
        /// Get Material Lists associated with the Alignment
        /// </summary>
        private List<MaterialList> GetMaterialLists(Alignment alignment, Transaction trans, Editor ed)
        {
            var materialLists = new List<MaterialList>();

            try
            {
                // Get all Material Lists in the drawing
                var materialListIds = CivilApplication.ActiveDocument.GetMaterialListIds();
                
                foreach (ObjectId materialListId in materialListIds)
                {
                    var materialList = trans.GetObject(materialListId, OpenMode.ForRead) as MaterialList;
                    if (materialList != null)
                    {
                        // Check if this Material List is associated with our Alignment
                        if (materialList.AlignmentId == alignment.ObjectId)
                        {
                            materialLists.Add(materialList);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError getting Material Lists: {ex.Message}");
            }

            return materialLists;
        }

        /// <summary>
        /// Extract volume data from Material Lists and their items
        /// </summary>
        private DataTable ExtractVolumeData(List<MaterialList> materialLists, Transaction trans)
        {
            var dataTable = new DataTable();
            
            // Setup table columns
            dataTable.Columns.Add("Material List", typeof(string));
            dataTable.Columns.Add("Material Name", typeof(string));
            dataTable.Columns.Add("Station Start", typeof(string));
            dataTable.Columns.Add("Station End", typeof(string));
            dataTable.Columns.Add("Cut Volume (m³)", typeof(double));
            dataTable.Columns.Add("Fill Volume (m³)", typeof(double));
            dataTable.Columns.Add("Net Volume (m³)", typeof(double));
            dataTable.Columns.Add("Cumulative Cut (m³)", typeof(double));
            dataTable.Columns.Add("Cumulative Fill (m³)", typeof(double));

            double cumulativeCut = 0.0;
            double cumulativeFill = 0.0;

            foreach (var materialList in materialLists)
            {
                try
                {
                    // Get Material List Items
                    var materialListItemIds = materialList.GetMaterialListItemIds();
                    
                    foreach (ObjectId itemId in materialListItemIds)
                    {
                        var materialListItem = trans.GetObject(itemId, OpenMode.ForRead) as MaterialListItem;
                        if (materialListItem != null)
                        {
                            // Get Material Quantities for this item
                            var quantities = GetMaterialQuantities(materialListItem, trans);
                            
                            foreach (var quantity in quantities)
                            {
                                cumulativeCut += quantity.CutVolume;
                                cumulativeFill += quantity.FillVolume;
                                
                                var row = dataTable.NewRow();
                                row["Material List"] = materialList.Name;
                                row["Material Name"] = materialListItem.Name;
                                row["Station Start"] = quantity.StationStart.ToString("F3");
                                row["Station End"] = quantity.StationEnd.ToString("F3");
                                row["Cut Volume (m³)"] = quantity.CutVolume;
                                row["Fill Volume (m³)"] = quantity.FillVolume;
                                row["Net Volume (m³)"] = quantity.CutVolume - quantity.FillVolume;
                                row["Cumulative Cut (m³)"] = cumulativeCut;
                                row["Cumulative Fill (m³)"] = cumulativeFill;
                                
                                dataTable.Rows.Add(row);
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    // Continue processing other material lists if one fails
                    continue;
                }
            }

            return dataTable;
        }

        /// <summary>
        /// Get Material Quantities from a Material List Item
        /// </summary>
        private List<MaterialQuantityInfo> GetMaterialQuantities(MaterialListItem materialListItem, Transaction trans)
        {
            var quantities = new List<MaterialQuantityInfo>();

            try
            {
                // Access material quantity data from the material list item
                // This is a simplified approach - actual implementation may vary based on Civil 3D API
                var quantityCollection = materialListItem.MaterialQuantities;
                
                foreach (MaterialQuantity quantity in quantityCollection)
                {
                    var quantityInfo = new MaterialQuantityInfo
                    {
                        StationStart = quantity.StartStation,
                        StationEnd = quantity.EndStation,
                        CutVolume = quantity.CutVolume,
                        FillVolume = quantity.FillVolume
                    };
                    
                    quantities.Add(quantityInfo);
                }
            }
            catch (System.Exception)
            {
                // Return empty list if unable to get quantities
            }

            return quantities;
        }

        /// <summary>
        /// Display volume data in table format
        /// </summary>
        private void DisplayVolumeTable(DataTable volumeData, Editor ed)
        {
            if (volumeData.Rows.Count == 0)
            {
                ed.WriteMessage("\nNo volume data found.");
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine("\n" + new string('=', 120));
            sb.AppendLine("EXCAVATION AND FILLING VOLUME INFORMATION - CROSS SECTION");
            sb.AppendLine(new string('=', 120));

            // Header row
            sb.AppendFormat("{0,-20} {1,-15} {2,-12} {3,-12} {4,-12} {5,-12} {6,-12} {7,-12} {8,-12}\n",
                "Material List", "Material", "Start Stn", "End Stn", "Cut Vol", "Fill Vol", "Net Vol", "Cum Cut", "Cum Fill");
            sb.AppendLine(new string('-', 120));

            // Data rows
            double totalCut = 0, totalFill = 0;
            foreach (DataRow row in volumeData.Rows)
            {
                totalCut += Convert.ToDouble(row["Cut Volume (m³)"]);
                totalFill += Convert.ToDouble(row["Fill Volume (m³)"]);

                sb.AppendFormat("{0,-20} {1,-15} {2,-12} {3,-12} {4,-12:F2} {5,-12:F2} {6,-12:F2} {7,-12:F2} {8,-12:F2}\n",
                    row["Material List"].ToString().Substring(0, Math.Min(20, row["Material List"].ToString().Length)),
                    row["Material Name"].ToString().Substring(0, Math.Min(15, row["Material Name"].ToString().Length)),
                    row["Station Start"],
                    row["Station End"],
                    row["Cut Volume (m³)"],
                    row["Fill Volume (m³)"],
                    row["Net Volume (m³)"],
                    row["Cumulative Cut (m³)"],
                    row["Cumulative Fill (m³)"]);
            }

            // Summary
            sb.AppendLine(new string('-', 120));
            sb.AppendFormat("SUMMARY: Total Cut Volume: {0:F2} m³, Total Fill Volume: {1:F2} m³, Net Volume: {2:F2} m³\n",
                totalCut, totalFill, totalCut - totalFill);
            sb.AppendLine(new string('=', 120));

            ed.WriteMessage(sb.ToString());
        }
    }

    /// <summary>
    /// Helper class to store material quantity information
    /// </summary>
    public class MaterialQuantityInfo
    {
        public double StationStart { get; set; }
        public double StationEnd { get; set; }
        public double CutVolume { get; set; }
        public double FillVolume { get; set; }
    }
}