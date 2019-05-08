using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Text;

/// <summary>
/// Thanks to @Jess
/// https://stackoverflow.com/questions/657131/how-to-read-data-of-an-excel-file-using-c/19065266#19065266
/// https://github.com/OfficeDev/Open-XML-SDK
/// </summary>
namespace MercyShipsTimeEntry
{
    public class XLSXHelper
    {
        /// <summary>
        /// Got code from: https://msdn.microsoft.com/en-us/library/office/gg575571.aspx
        /// </summary>
        public string GetCroppedExcelFile(
            string fileName = "ExcelFiles\\File_With_Many_Tabs.xlsx",
            string sheetName = "Submission Form",
            string cropStartWhenFindThisValue = "RowWithSpecialCellValue_Start",
            string cropStopWhenFindThisValue = "RowWithSpecialCellValue_Stop",
            string outputDelim = "|",
            bool crop = true
            )
        {
            StringBuilder sb = new StringBuilder();
            using (var document = SpreadsheetDocument.Open(fileName, isEditable: false))
            {
                WorkbookPart workbookPart;
                SheetData sheetData = GetSheet(sheetName, document, out workbookPart);
                LoopThroughCellValues(
                    cropStartWhenFindThisValue, 
                    cropStopWhenFindThisValue, 
                    outputDelim, crop, sb, 
                    workbookPart, sheetData);
            }
            return sb.ToString().Replace("\n\n", "\n");
        }
        public string GetCroppedExcelFile(string xlsFileName, TimeSheetPerson cropSinglePerson, string delim = "|", bool crop = true)
        {
            return GetCroppedExcelFile(
                xlsFileName,
                DateTime.Now.ToString("MMM"),
                cropSinglePerson.Name,
                "Total",
                delim,
                crop);
        }

        private void LoopThroughCellValues(string cropStartWhenFindThisValue, string cropStopWhenFindThisValue, string outputDelim, bool crop, StringBuilder sb, WorkbookPart workbookPart, SheetData sheetData)
        {
            bool foundSpecialToken = !crop;
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellValue = GetCellValue(cell, workbookPart);
                    if (crop)
                    {
                        foundSpecialToken = (foundSpecialToken || (cellValue == cropStartWhenFindThisValue)) && (cellValue != cropStopWhenFindThisValue);
                    }
                }
                if (foundSpecialToken)
                {
                    sb.AppendLine(String.Join(outputDelim, row.Elements<Cell>()
                        .Select(cell =>
                            GetCellValue(cell, workbookPart))));
                }
            }
        }

        private static SheetData GetSheet(string sheetName, SpreadsheetDocument document, out WorkbookPart workbookPart)
        {
            workbookPart = document.WorkbookPart;
            var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name.ToString().Trim() == sheetName);
            if (sheet == null)
            {
                throw new EntryPointNotFoundException(String.Format(@"Sheetname ""{0}"" not found", sheetName));
            }
            var worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
            return worksheetPart.Worksheet.Elements<SheetData>().First();
        }

        

            /// <summary>
            /// Got code from: https://msdn.microsoft.com/en-us/library/office/hh298534.aspx
            /// </summary>
            /// <param name="cell"></param>
            /// <param name="workbookPart"></param>
            /// <returns></returns>
            private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell == null)
            {
                return null;
            }

            var value = cell.CellFormula != null
                ? cell.CellValue.InnerText
                : cell.InnerText.Trim();

            // If the cell represents an integer number, you are done. 
            // For dates, this code returns the serialized value that 
            // represents the date. The code handles strings and 
            // Booleans individually. For shared strings, the code 
            // looks up the corresponding value in the shared string 
            // table. For Booleans, the code converts the value into 
            // the words TRUE or FALSE.
            if (cell.DataType == null)
            {
                return value;
            }
            switch (cell.DataType.Value)
            {
                case CellValues.SharedString:

                    // For shared strings, look up the value in the
                    // shared strings table.
                    var stringTable =
                        workbookPart.GetPartsOfType<SharedStringTablePart>()
                            .FirstOrDefault();

                    // If the shared string table is missing, something 
                    // is wrong. Return the index that is in
                    // the cell. Otherwise, look up the correct text in 
                    // the table.
                    if (stringTable != null)
                    {
                        value =
                            stringTable.SharedStringTable
                                .ElementAt(int.Parse(value)).InnerText;
                    }
                    break;

                case CellValues.Boolean:
                    switch (value)
                    {
                        case "0":
                            value = "FALSE";
                            break;
                        default:
                            value = "TRUE";
                            break;
                    }
                    break;
            }
            return value;
        }
    }
}
