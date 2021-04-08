using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlEx;
using OpenXmlEx.Styles;

namespace OpenXmlExTests
{
    internal static class BaseTestData
    {
        #region Базовые данные для теста DRY

        /// <summary> Список номеров имён для тестовых данных </summary>
        internal static readonly Stack<int> Names = new(Enumerable.Range(1, 100));

        internal static string NewName => $"Test_{Names.Pop()}.xlsx";
        /// <summary> Генерирует базовые тестовые данные </summary>
        internal static WorksheetPart GetBaseSpreadsheetDocument()
        {
            var spread_sheet = SpreadsheetDocument.Create(NewName, SpreadsheetDocumentType.Workbook);
            // create the workbook
            var workbook_part = spread_sheet.AddWorkbookPart();
            var workbook = workbook_part.Workbook = new Workbook();
            var sheets = workbook.AppendChild(new Sheets());
            var worksheet_part = workbook_part.AddNewPart<WorksheetPart>();
            var sheet = new Sheet { Id = workbook_part.GetIdOfPart(worksheet_part), SheetId = 1, Name = "first_sheet" };
            // ReSharper disable once PossiblyMistakenUseOfParamsMethod
            sheets.Append(sheet);
            return worksheet_part;
        }
        /// <summary> Генерирует базовые тестовые данные </summary>
        internal static OpenXmlWriterEx GetBaseXmlWriterTestData() => new(GetBaseSpreadsheetDocument(), new OpenXmlExStyles());

        /// <summary> Генерирует данные для работы с листом </summary>
        internal static OpenXmlWriterEx GetXmlWriterSheetTestData()
        {
            var writer = GetBaseXmlWriterTestData();
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());
            return writer;
        }

        #endregion
    }
}
