using System;
using DocToExcelNoteParser.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using NPOI.HPSF;
namespace DocToExcelNoteParser.Workers
{
	public class ExcelCreator
	{
		public void GenerateExcel(IEnumerable<FootNoteToken> footNotes)
		{
            var fileName = GetDateFileName();

            using var spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
            var workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sheet1"
            };
            sheets.Append(sheet);

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            AddColumnTitles(sheetData);

            foreach (var footNote in footNotes)
                AddRow(sheetData, footNote.FootNoteName ?? "", footNote.FootNoteContent ?? "");

            workbookPart.Workbook.Save();
        }

        private string GetDateFileName()
        {
            var now = DateTime.Now;
            var timeString = now.ToString("HH:mm");
            var fileName = $"excel_{now:dd.MM.yyyy}_{timeString}.xlsx";

            return fileName;
        }

        private void AddColumnTitles(SheetData sheetData)
        {
            var titleRow = new Row();

            titleRow.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Слово")
            });

            titleRow.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Текст")
            });

            sheetData.Append(titleRow);
        }

        private void AddRow(SheetData sheetData, string cellNameValue, string cellContentValue)
        {
            var row = new Row();

            row.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(cellNameValue)
            });

            row.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(cellContentValue)
            });

            sheetData.Append(row);
        }
	}
}

