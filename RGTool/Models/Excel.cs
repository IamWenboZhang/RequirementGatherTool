using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace RGTool.Models
{
    public class ExcelItem
    {
        public string ID { get; set; }
        public string RequirementID { get; set; }
        public string SectionID { get; set; }
        public string Itemcontent { get; set; }
        public string InformaltiveorNormaltive { get; set; }
        public string ServerorClient { get; set; }
    }
    public class ExcelUtil
    {
        public static void ApplyTemplate(string templateName, string jsonStr, string newDocName, string configJson)
        {
            //Deserialize JsonObeject from Json string
            List<ExcelItem> items = JsonConvert.DeserializeObject<List<ExcelItem>>(jsonStr);
            Config config = JsonConvert.DeserializeObject<Config>(configJson);
            File.Copy(templateName, newDocName);
            ExcelUtil.applyTemplate(newDocName, items, config);
            //CreatePackage(newDocName,items,templateName);
            //ExcelUtil.CreateSpreadsheetWorkbook(newDocName);
            //ExcelUtil.InsertText(newDocName, items);
        }

        public static void applyTemplate(string fileName, List<ExcelItem> items, Config config)
        {
            using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(fileName, true))
            {
                uint i = 20;
                ChangeTextFromCell(mySpreadsheet, "Requirements", "A", 1, config.TDShortName);
                ChangeTextFromCell(mySpreadsheet, "Requirements", "A", 2, config.TDShortName);
                ChangeTextFromCell(mySpreadsheet, "Requirements", "C", 3, config.TDVersion);
                foreach (var item in items)
                {
                    if (i == 34)
                    {
                        Console.Write("Come on");
                    }
                    ChangeTextFromCell(mySpreadsheet, "Requirements", "A", i, item.RequirementID);
                    ChangeTextFromCell(mySpreadsheet, "Requirements", "B", i, item.SectionID);
                    ChangeTextFromCell(mySpreadsheet, "Requirements", "C", i, item.Itemcontent);
                    ChangeTextFromCell(mySpreadsheet, "Requirements", "F", i, item.ServerorClient);
                    ChangeTextFromCell(mySpreadsheet, "Requirements", "G", i, item.InformaltiveorNormaltive);
                    ChangeTextFromCell(mySpreadsheet, "Requirements", "J", i, item.ID);
                    i++;
                }
            } 
        }

        public static void ChangeTextFromCell(SpreadsheetDocument mySpreadsheet, string sheetName, string colName, uint rowIndex, string text)
        {
            DeleteTextFromCell(mySpreadsheet, sheetName, colName, rowIndex);
            IEnumerable<Sheet> sheets = mySpreadsheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)mySpreadsheet.WorkbookPart.GetPartById(relationshipId);
            Cell cell = InsertCellInWorksheet(colName, rowIndex, worksheetPart);
            // Set the value of cell A1.
            cell.CellValue = new CellValue(text);
            cell.DataType = new EnumValue<CellValues>(CellValues.String);

            // Save the new worksheet.
            worksheetPart.Worksheet.Save();

        }

        // Given a document, a worksheet name, a column name, and a one-based row index,
        // deletes the text from the cell at the specified column and row on the specified worksheet.
        public static void DeleteTextFromCell(SpreadsheetDocument document, string sheetName, string colName, uint rowIndex)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);

            // Get the cell at the specified column and row.
            Cell cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex);
            if (cell == null)
            {
                // The specified cell does not exist.
                return;
            }

            cell.Remove();
            worksheetPart.Worksheet.Save();
        }

        // Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
        private static Cell GetSpreadsheetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex);
            if (rows.Count() == 0)
            {
                // A cell does not exist at the specified row.
                return null;
            }

            IEnumerable<Cell> cells = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
            if (cells.Count() == 0)
            {
                // A cell does not exist at the specified column, in the specified row.
                return null;
            }

            return cells.First();
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }


        //// Given a document, a worksheet name, a column name, and a one-based row index,
        //// deletes the text from the cell at the specified column and row on the specified worksheet.
        //public static void ChangeTextFromCell(SpreadsheetDocument document, string sheetName, string columnName, uint rowIndex, string text)
        //{
        //    if(columnName == "A" && rowIndex ==36)
        //    {
        //        Console.Write("What's the hell!");
        //    }
        //    try
        //    {
        //        IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
        //        if (sheets.Count() == 0)
        //        {
        //            // The specified worksheet does not exist.
        //            return;
        //        }
        //        string relationshipId = sheets.First().Id.Value;
        //        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
        //        SheetData sheetdata = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        //        if (columnName == "A" && rowIndex >= 36)
        //        {
        //            Row newrow = new Row();
        //            newrow.RowIndex = rowIndex;
        //            sheetdata.Append(newrow);
        //        }
        //        //Row currentRow = GetRow(rowIndex);
        //        // Get the cell at the specified column and row.
        //        Cell cell = GetSpreadsheetCell(worksheetPart.Worksheet, columnName, rowIndex);
        //        if (cell == null)
        //        {
        //            // The specified cell does not exist.
        //            cell = new Cell() { CellReference = columnName + rowIndex.ToString(), DataType = CellValues.InlineString };
        //            InlineString inlineStrnull = CreateInlineString(text);
        //            cell.Append(inlineStrnull);
        //            sheetdata.Append(cell);
        //        }
        //        else
        //        {
        //            cell.Remove();
        //            InlineString inlineStr = CreateInlineString(text);
        //            cell.Append(inlineStr);
        //        }

        //        worksheetPart.Worksheet.Save();
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.Write(ex.Message);
        //    }
        //}

        //public static Row GetRow( uint rowIndex)
        //{
        //    SheetData sheetdata = worksheet.GetFirstChild<SheetData>();
        //    IEnumerable<Row> rows = sheetdata.Elements<Row>().Where(r => r.RowIndex == rowIndex);
        //    if (rows.Count() == 0)
        //    {
        //        // A cell does not exist at the specified row.
        //        Row newrow = new Row();
        //        newrow.RowIndex = rowIndex;
        //        sheetdata.Append(newrow);
        //    }
        //}

        //// Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
        //private static Cell GetSpreadsheetCell(Worksheet worksheet, string columnName, uint rowIndex)
        //{
        //    SheetData sheetdata = worksheet.GetFirstChild<SheetData>();
        //    IEnumerable<Row> rows = sheetdata.Elements<Row>().Where(r => r.RowIndex == rowIndex);
        //    if (rows.Count() == 0)
        //    {
        //        // A cell does not exist at the specified row.
        //        Row newrow = new Row();
        //        newrow.RowIndex = rowIndex;
        //        sheetdata.Append(newrow);
        //    }

        //    IEnumerable<Cell> cells = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
        //    if (cells.Count() == 0)
        //    {
        //        // A cell does not exist at the specified column, in the specified row.
        //        return null;
        //    }

        //    return cells.First();
        //}

        //// Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
        //// reference the specified SharedStringItem and removes the item.
        //private static void RemoveSharedStringItem(int shareStringId, SpreadsheetDocument document)
        //{
        //    bool remove = true;

        //    foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
        //    {
        //        Worksheet worksheet = part.Worksheet;
        //        foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
        //        {
        //            // Verify if other cells in the document reference the item.
        //            if (cell.DataType != null &&
        //                cell.DataType.Value == CellValues.SharedString &&
        //                cell.CellValue.Text == shareStringId.ToString())
        //            {
        //                // Other cells in the document still reference the item. Do not remove the item.
        //                remove = false;
        //                break;
        //            }
        //        }

        //        if (!remove)
        //        {
        //            break;
        //        }
        //    }

        //    // Other cells in the document do not reference the item. Remove the item.
        //    if (remove)
        //    {
        //        SharedStringTablePart shareStringTablePart = document.WorkbookPart.SharedStringTablePart;
        //        if (shareStringTablePart == null)
        //        {
        //            return;
        //        }

        //        SharedStringItem item = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(shareStringId);
        //        if (item != null)
        //        {
        //            item.Remove();

        //            // Refresh all the shared string references.
        //            foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
        //            {
        //                Worksheet worksheet = part.Worksheet;
        //                foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
        //                {
        //                    if (cell.DataType != null &&
        //                        cell.DataType.Value == CellValues.SharedString)
        //                    {
        //                        int itemIndex = int.Parse(cell.CellValue.Text);
        //                        if (itemIndex > shareStringId)
        //                        {
        //                            cell.CellValue.Text = (itemIndex - 1).ToString();
        //                        }
        //                    }
        //                }
        //                worksheet.Save();
        //            }

        //            document.WorkbookPart.SharedStringTablePart.SharedStringTable.Save();
        //        }
        //    }
        //}



        //public static void GetSheetInfo(string fileName)
        //{
        //    // Open file as read-only.
        //    using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(fileName, false))
        //    {
        //        Sheets sheets = mySpreadsheet.WorkbookPart.Workbook.Sheets;

        //        // For each sheet, display the sheet information.
        //        foreach (Sheet sheet in sheets)
        //        {
        //            foreach (var attr in sheet.GetAttributes())
        //            {
        //                Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value);
        //            }
        //        }
        //    }
        //}

        //// Given a document name and text, 
        //// inserts a new worksheet and writes the text to cell "A1" of the new worksheet.
        //public static void InsertText(string docName, List<ExcelItem> items)
        //{
        //    // Open the document for editing.
        //    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
        //    {
        //        // Get the SharedStringTablePart. If it does not exist, create a new one.
        //        SharedStringTablePart shareStringPart;
        //        if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
        //        {
        //            shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
        //        }
        //        else
        //        {
        //            shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
        //        }

        //        //// Insert the text into the SharedStringTablePart.
        //        //int index = InsertSharedStringItem(text, shareStringPart);

        //        // Insert a new worksheet.
        //        //WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);
        //        WorksheetPart worksheetPart = spreadSheet.WorkbookPart.WorksheetParts.First<WorksheetPart>();
        //        uint i = 34;
        //        foreach (var item in items)
        //        {
        //            // Insert the text into the SharedStringTablePart.
        //            int indexID = InsertSharedStringItem(item.ID, shareStringPart);
        //            int indexRequirementID = InsertSharedStringItem(item.RequirementID, shareStringPart);
        //            int indexSectionID = InsertSharedStringItem(item.SectionID, shareStringPart);
        //            int indexContent = InsertSharedStringItem(item.Itemcontent, shareStringPart);
        //            int indexInformaltiveorNormaltive = InsertSharedStringItem(item.InformaltiveorNormaltive, shareStringPart);
        //            int indexServerorClient = InsertSharedStringItem(item.ServerorClient, shareStringPart);

        //            // Insert cell A1 into the new worksheet.
        //            Cell cellRequirementID = InsertCellInWorksheet("A", i, worksheetPart);
        //            Cell cellSectionID = InsertCellInWorksheet("B", i, worksheetPart);
        //            Cell cellContent = InsertCellInWorksheet("C", i, worksheetPart);
        //            Cell cellInformaltiveorNormaltive = InsertCellInWorksheet("D", i, worksheetPart);
        //            Cell cellServerorClient = InsertCellInWorksheet("E", i, worksheetPart);
        //            Cell cellID = InsertCellInWorksheet("F", i, worksheetPart);
        //            // Set the value of cell A1.
        //            cellRequirementID.CellValue = new CellValue(indexRequirementID.ToString());
        //            cellRequirementID.DataType = new EnumValue<CellValues>(CellValues.SharedString);

        //            cellSectionID.CellValue = new CellValue(indexSectionID.ToString());
        //            cellSectionID.DataType = new EnumValue<CellValues>(CellValues.SharedString);

        //            cellContent.CellValue = new CellValue(indexContent.ToString());
        //            cellContent.DataType = new EnumValue<CellValues>(CellValues.SharedString);

        //            cellInformaltiveorNormaltive.CellValue = new CellValue(indexInformaltiveorNormaltive.ToString());
        //            cellInformaltiveorNormaltive.DataType = new EnumValue<CellValues>(CellValues.SharedString);

        //            cellServerorClient.CellValue = new CellValue(indexServerorClient.ToString());
        //            cellServerorClient.DataType = new EnumValue<CellValues>(CellValues.SharedString);

        //            cellID.CellValue = new CellValue(indexID.ToString());
        //            cellID.DataType = new EnumValue<CellValues>(CellValues.SharedString);

        //            i++;
        //        }

        //        //// Insert cell A1 into the new worksheet.
        //        //Cell cell = InsertCellInWorksheet("A", 1, worksheetPart);

        //        //// Set the value of cell A1.
        //        //cell.CellValue = new CellValue(index.ToString());
        //        //cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

        //        // Save the new worksheet.
        //        worksheetPart.Worksheet.Save();
        //    }
        //}


        //public static void CreatePackage(string filePath, List<ExcelItem> items, string templateName)
        //{
        //    using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        //    {
        //        CreateParts(package,items, templateName);
        //    }
        //}
        //private static void CreateParts(SpreadsheetDocument document, List<ExcelItem> items, string templateName)
        //{
        //    WorkbookPart workbookPart1 = document.AddWorkbookPart();
        //    GenerateWorkbookPart1Content(workbookPart1);

        //    WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
        //    GenerateWorksheetPart1Content(worksheetPart1, items, templateName);
        //}
        //private static void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
        //{
        //    Workbook workbook1 = new Workbook();
        //    workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        //    Sheets sheets1 = new Sheets();
        //    Sheet sheet1 = new Sheet() { Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };
        //    sheets1.Append(sheet1);

        //    workbook1.Append(sheets1);
        //    workbookPart1.Workbook = workbook1;
        //}
        //private static void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1, List<ExcelItem> items, string templateName)
        //{
        //    Worksheet worksheet1 = new Worksheet();
        //    SheetData sheetData1 = new SheetData();
        //    //Add Template
        //    GetSheetInfo(templateName);
        //    for(int i = 0;i<items.Count;i++)
        //    {
        //        var row = CreateRow(items[i], i+20);
        //        sheetData1.Append(row);
        //    }

        //    worksheet1.Append(sheetData1);
        //    worksheetPart1.Worksheet = worksheet1;
        //}

        //    private static Row CreateRow(ExcelItem item, int rowindex)
        //{
        //    Row row = new Row();
        //    Cell cellRequirementID = new Cell() { CellReference = "A" + rowindex.ToString(), DataType = CellValues.InlineString };
        //    Cell cellSectionID = new Cell() { CellReference = "B" + rowindex.ToString(), DataType = CellValues.InlineString };
        //    Cell cellContent = new Cell() { CellReference = "C" + rowindex.ToString(), DataType = CellValues.InlineString };
        //    Cell cellInformaltiveorNormaltive = new Cell() { CellReference = "D" + rowindex.ToString(), DataType = CellValues.InlineString };
        //    Cell cellServerorClient = new Cell() { CellReference = "E" + rowindex.ToString(), DataType = CellValues.InlineString };
        //    Cell cellID = new Cell() { CellReference = "F" + rowindex.ToString(), DataType = CellValues.InlineString };

        //    InlineString inlineStrRequirementID = CreateInlineString(item.RequirementID);
        //    InlineString inlineStrSectionID = CreateInlineString(item.SectionID);
        //    InlineString inlineStrContent = CreateInlineString(item.Itemcontent);
        //    InlineString inlineStrInformaltiveorNormaltive = CreateInlineString(item.InformaltiveorNormaltive);
        //    InlineString inlineStrServerorClient = CreateInlineString(item.ServerorClient);
        //    InlineString inlineStrID = CreateInlineString(item.ID);


        //    cellRequirementID.Append(inlineStrRequirementID);
        //    cellSectionID.Append(inlineStrSectionID);
        //    cellContent.Append(inlineStrContent);
        //    cellInformaltiveorNormaltive.Append(inlineStrInformaltiveorNormaltive);
        //    cellServerorClient.Append(inlineStrServerorClient);
        //    cellID.Append(inlineStrID);

        //    row.Append(cellRequirementID);
        //    row.Append(cellSectionID);
        //    row.Append(cellContent);
        //    row.Append(cellInformaltiveorNormaltive);
        //    row.Append(cellServerorClient);
        //    row.Append(cellID);

        //    return row;
        //}

        //private static InlineString CreateInlineString(string content)
        //{
        //    InlineString inlineString = new InlineString();
        //    Text text = new Text();
        //    text.Text = content;
        //    inlineString.Append(text);
        //    return inlineString;
        //}
        //// Given a document name and text, 
        //// inserts a new worksheet and writes the text to cell "A1" of the new worksheet.
        //public static void InsertText(string docName, string text, string newDocName)
        //{
        //    List<ExcelItem> items = JsonConvert.DeserializeObject<List<ExcelItem>>(text);
        //    var newSpreadsheetDocument = SpreadsheetDocument.Create(newDocName, SpreadsheetDocumentType.Workbook);
        //    // Open the document for editing.
        //    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
        //    {
        //        // Get the SharedStringTablePart. If it does not exist, create a new one.
        //        SharedStringTablePart shareStringPart;
        //        if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
        //        {
        //            shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
        //        }
        //        else
        //        {
        //            shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
        //        }
        //        uint i = 0;
        //        // Insert a new worksheet.
        //        WorksheetPart worksheetPart = InsertWorksheet(newSpreadsheetDocument.WorkbookPart);
        //        foreach (var item in items)
        //        {
        //            // Insert the text into the SharedStringTablePart.
        //            int index = InsertSharedStringItem(text, shareStringPart);

        //            // Insert cell A1 into the new worksheet.
        //            Cell cellRequirementID = InsertCellInWorksheet("A", i, worksheetPart);
        //            Cell cellSectionID = InsertCellInWorksheet("B", i, worksheetPart);
        //            Cell cellContent = InsertCellInWorksheet("C", i, worksheetPart);
        //            Cell cellInformaltiveorNormaltive = InsertCellInWorksheet("D", i, worksheetPart);
        //            Cell cellServerorClient = InsertCellInWorksheet("E", i, worksheetPart);
        //            Cell cellID = InsertCellInWorksheet("F", i, worksheetPart);
        //            // Set the value of cell A1.
        //            cellRequirementID.CellValue = new CellValue(item.RequirementID);
        //            cellRequirementID.DataType = new EnumValue<CellValues>(CellValues.String);

        //            cellSectionID.CellValue = new CellValue(item.SectionID);
        //            cellSectionID.DataType = new EnumValue<CellValues>(CellValues.String);

        //            cellContent.CellValue = new CellValue(item.Itemcontent);
        //            cellContent.DataType = new EnumValue<CellValues>(CellValues.String);

        //            cellInformaltiveorNormaltive.CellValue = new CellValue(item.InformaltiveorNormaltive);
        //            cellInformaltiveorNormaltive.DataType = new EnumValue<CellValues>(CellValues.String);

        //            cellServerorClient.CellValue = new CellValue(item.ServerorClient);
        //            cellServerorClient.DataType = new EnumValue<CellValues>(CellValues.String);

        //            cellID.CellValue = new CellValue(item.ID);
        //            cellID.DataType = new EnumValue<CellValues>(CellValues.String);

        //            i++;
        //        }
        //        // Save the new worksheet.
        //        worksheetPart.Worksheet.Save();
        //    }
        //}

        //// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        //// and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        //private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        //{
        //    // If the part does not contain a SharedStringTable, create one.
        //    if (shareStringPart.SharedStringTable == null)
        //    {
        //        shareStringPart.SharedStringTable = new SharedStringTable();
        //    }

        //    int i = 0;

        //    // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        //    foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        //    {
        //        if (item.InnerText == text)
        //        {
        //            return i;
        //        }

        //        i++;
        //    }

        //    // The text does not exist in the part. Create the SharedStringItem and return its index.
        //    shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        //    shareStringPart.SharedStringTable.Save();

        //    return i;
        //}

        //// Given a WorkbookPart, inserts a new worksheet.
        //private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        //{
        //    // Add a new worksheet part to the workbook.
        //    WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        //    newWorksheetPart.Worksheet = new Worksheet(new SheetData());
        //    newWorksheetPart.Worksheet.Save();

        //    Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        //    string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

        //    // Get a unique ID for the new sheet.
        //    uint sheetId = 1;
        //    if (sheets.Elements<Sheet>().Count() > 0)
        //    {
        //        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        //    }

        //    string sheetName = "Sheet" + sheetId;

        //    // Append the new worksheet and associate it with the workbook.
        //    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        //    sheets.Append(sheet);
        //    workbookPart.Workbook.Save();

        //    return newWorksheetPart;
        //}
        //// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        //// If the cell already exists, returns it. 
        //private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        //{
        //    Worksheet worksheet = worksheetPart.Worksheet;
        //    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        //    string cellReference = columnName + rowIndex;

        //    // If the worksheet does not contain a row with the specified row index, insert one.
        //    Row row;
        //    if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
        //    {
        //        row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        //    }
        //    else
        //    {
        //        row = new Row() { RowIndex = rowIndex };
        //        sheetData.Append(row);
        //    }

        //    // If there is not a cell with the specified column name, insert one.  
        //    if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
        //    {
        //        return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
        //    }
        //    else
        //    {
        //        // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
        //        Cell refCell = null;
        //        foreach (Cell cell in row.Elements<Cell>())
        //        {
        //            if (cell.CellReference.Value.Length == cellReference.Length)
        //            {
        //                if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
        //                {
        //                    refCell = cell;
        //                    break;
        //                }
        //            }
        //        }

        //        Cell newCell = new Cell() { CellReference = cellReference };
        //        row.InsertBefore(newCell, refCell);

        //        worksheet.Save();
        //        return newCell;
        //    }
        //}

        //public static void CreateSpreadsheetWorkbook(string filepath)
        //{
        //    // Create a spreadsheet document by supplying the filepath.
        //    // By default, AutoSave = true, Editable = true, and Type = xlsx.
        //    SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
        //        Create(filepath, SpreadsheetDocumentType.Workbook);

        //    // Add a WorkbookPart to the document.
        //    WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
        //    workbookpart.Workbook = new Workbook();

        //    // Add a WorksheetPart to the WorkbookPart.
        //    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        //    worksheetPart.Worksheet = new Worksheet(new SheetData());

        //    // Add Sheets to the Workbook.
        //    Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
        //        AppendChild<Sheets>(new Sheets());

        //    // Append a new worksheet and associate it with the workbook.
        //    Sheet sheet = new Sheet()
        //    {
        //        Id = spreadsheetDocument.WorkbookPart.
        //        GetIdOfPart(worksheetPart),
        //        SheetId = 1,
        //        Name = "mySheet"
        //    };
        //    sheets.Append(sheet);

        //    workbookpart.Workbook.Save();

        //    // Close the document.
        //    spreadsheetDocument.Close();
        //}
    }
}