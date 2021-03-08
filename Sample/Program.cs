using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32.SafeHandles;
using OpenXmlEx;
using OpenXmlEx.Styles;
using OpenXmlEx.Styles.Base;
using Color = System.Drawing.Color;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            var FileName = "test.xlsx";
            var fonts = new[] { "Times New Roman", "Calibri", "Arial" };
            var fills = new[] { System.Drawing.Color.BlueViolet, System.Drawing.Color.Crimson };
            var sizes = new[] { 8U, 10U, 12U, 14U, 16U };

            using var document = SpreadsheetDocument.Create(FileName, SpreadsheetDocumentType.Workbook);
            // create the workbook
            var workbook_part = document.AddWorkbookPart();
            var wbsp = workbook_part.AddNewPart<WorkbookStylesPart>();

            #region Styles

            var styles = OpenXmlEx.OpenXmlEx.GetStyles(
                new List<OpenXmlExStyle>()
                {
                    new OpenXmlExStyle() {FontColor = Color.Crimson, IsBoldFont = true},
                    new OpenXmlExStyle() {FontSize = 20, FontName = "Calibri", BorderColor = Color.Red}
                });

            wbsp.Stylesheet = styles.Styles;
            wbsp.Stylesheet.Save();

            #endregion

            #region document start

            var workbook = workbook_part.Workbook = new Workbook();
            var sheets = workbook.AppendChild(new Sheets());

            #endregion


            #region Sheet DATA

            var sheet_name = "Test_sheet_name";
            var ws_part = workbook_part.AddNewPart<WorksheetPart>();

            // sheet
            var sheet_1 = new Sheet { Id = workbook_part.GetIdOfPart(ws_part), SheetId = 1, Name = sheet_name };

            // ReSharper disable once PossiblyMistakenUseOfParamsMethod
            sheets.Append(sheet_1);
            #endregion

            using var writer = new OpenXmlEx.OpenXmlEx(ws_part, new OpenXmlExStyles());

            writer.WriteStartElement(new Worksheet());

            #region Надстройка страницы - кнопки группировки сверху

            writer.SetGrouping(false, false);

            #endregion

            #region Установка ширины колонок

            //Установка размеров колонок
            var width_setting = new List<(uint First, uint Last, double width)>
            {
                (1, 2, 7),
                (3, 3, 11),
                (4, 12, 9.5),
                (13, 13, 17),
                (14, 14, 40),
                (15, 16, 15),
                (18, 20, 15)
            };
            writer.SetWidth(width_setting);

            #endregion

            writer.WriteStartElement(new SheetData());

            var mer_list = new List<MergeCell>();

            #region 1 лист


            var (key, value) = writer.FindStyleOrDefault(
                new OpenXmlExStyle()
                {
                    FontColor = Color.Crimson,
                    //FontSize = 20,
                    //IsBoldFont = true,
                    //LeftBorderStyle =  BorderStyleValues.Dashed,
                    //RightBorderStyle = BorderStyleValues.Dashed
                });
            writer.AddRow(3);
            writer.AddCell("Test",3,3, key);
            writer.CloseRow(3);
            writer.AddRow(4);
            writer.AddCell("Test",4,4, key);
            writer.AddCell("Test",4,5, key);
            writer.CloseRow(4);
            writer.WriteEndElement(); //end of SheetData

            #region Секция настроек

            //Секция с фильтром (нужно утвердить на листе)
            writer.SetFilter(sheet_name, 1, 20, 3, 3);

            //Секция с объединенными ячейками должна быть в конце перед закрытием секции WorkSheet
            if (mer_list.Count > 0)
                writer.SetMergedList(mer_list);

            #endregion

            writer.WriteEndElement(); //end of worksheet

            #endregion
        }

        private static void Test(string FileName)
        {
            using var spread_sheet = SpreadsheetDocument.Create(FileName, SpreadsheetDocumentType.Workbook);
            // create the workbook
            var workbook_part = spread_sheet.AddWorkbookPart();

            var wbsp = workbook_part.AddNewPart<WorkbookStylesPart>();
            //wbsp.Stylesheet = helper.GenerateStyleSheet();
            wbsp.Stylesheet.Save();


            var workbook = workbook_part.Workbook = new Workbook();
            var sheets = workbook.AppendChild(new Sheets());




            // create worksheet 1
            var first_sheet_name = "Faults";
            var worksheet_part_1 = workbook_part.AddNewPart<WorksheetPart>();
            var sheet_1 = new Sheet { Id = workbook_part.GetIdOfPart(worksheet_part_1), SheetId = 1, Name = first_sheet_name };
            // ReSharper disable once PossiblyMistakenUseOfParamsMethod
            sheets.Append(sheet_1);

            var mer_list = new List<MergeCell>();
            using var writer = OpenXmlWriter.Create(worksheet_part_1);
        }
    }
}
