using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using ExcelDataReader;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.IO;
using System.Linq;

namespace Csharp_WriteDataToExcel
{
    public class WriteDataToExcel
    {

        //To Find the Excel==========>Get the Excelfilepath with the Excelname
        public static string ExcelFilePath()
        {
            string currentDirectoryPath = Environment.CurrentDirectory;
            string actualPath = currentDirectoryPath.Substring(0, currentDirectoryPath.LastIndexOf("bin"));
            string projectPath = new Uri(actualPath).LocalPath;
            string ExcelfilePath = projectPath + "\\Excelsheets\\TestData.xlsx";
            return ExcelfilePath;
        }
        // To Write Data in Excel
        //Update Sheet
        public static void UpdateSheet(string docName, string strSheetName, string text, int rowIndex, string columnName)
        {
            try
            {
                UpdateCell(docName, strSheetName, text, rowIndex, columnName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in Updatesheet " + ex.Message);
            }

        }

        //Update Cell
        public static void UpdateCell(string docName, string strSheetName, string text, int rowIndex, string columnName)
        {
            //Open the Spreadsheet Excel document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {

                //Go to the Sheet
                WorksheetPart worksheetPart = GetWorksheetPartByName(spreadSheet, strSheetName);

                if (worksheetPart != null)
                {
                    //Insert a Cell in the sheet
                    InsertCellInWorksheet(columnName, (UInt32)rowIndex, worksheetPart);

                    //Goto the Cell
                    Cell cell = GetCell(worksheetPart.Worksheet, columnName, rowIndex);
                    cell.CellValue = new CellValue(text);
                    cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
                    cell.InlineString = new InlineString() { Text = new Text(text) };

                    //Save the WorkSheet
                    worksheetPart.Worksheet.Save();
                }
                //close the Excel Spreadsheet==>(Note .close() will be deprecated soon use Dispose())
                spreadSheet.Dispose();

            }
        }

        //Inserting New Call in Worksheet
        public static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;
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

            Cell refCell = row.Descendants<Cell>().LastOrDefault();
            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertAfter(newCell, refCell);

            worksheet.Save();
            return newCell;
        }

        //
        private static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist . 
                return null;
            }
            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;
        }

        //Given -> Worksheet, ColumnName, Rowindex
        //Gets the cell at the specified Column
        private static Cell GetCell(Worksheet worksheet, string columnName, int rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);
            if (row == null)
            {
                return null;
            }
            string strCellAddress = columnName + rowIndex;
            var response = row.Elements<Cell>().Where(c => c.CellReference == strCellAddress).First();
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == strCellAddress).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == strCellAddress).First();
            }
            else
            {
                Cell cellRef = new Cell();
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, strCellAddress, true) > 0)
                    {
                        cellRef = cell;
                        break;
                    }
                }
                Cell newCell = new Cell()
                {
                    CellReference = strCellAddress,
                    StyleIndex = (UInt32Value)1U
                };
                row.InsertBefore(newCell, cellRef);
                worksheet.Save();
                return newCell;
            }
        }
        //Given Worksheet and RowIndex => Return the Row
        private static Row GetRow(Worksheet worksheet, int rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
    }
}
