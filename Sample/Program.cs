using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32.SafeHandles;
using OpenXmlEx;
using OpenXmlEx.Extensions;
using OpenXmlEx.Styles;
using OpenXmlEx.Styles.Base;
using OpenXmlEx.SubClasses;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;

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
            
            Test(FileName);
            new Process {StartInfo = new ProcessStartInfo(FileName) {UseShellExecute = true}}.Start();
        }

        static void Test(string FileName)
        {
            using var document = SpreadsheetDocument.Create(FileName, SpreadsheetDocumentType.Workbook);
            // create the workbook
            var workbook_part = document.AddWorkbookPart();
            var wbsp = workbook_part.AddNewPart<WorkbookStylesPart>();

            #region Styles

            var styles = OpenXmlWriterEx.GetStyles(
                new List<BaseOpenXmlExStyle>()
                {
                    new BaseOpenXmlExStyle() {FontColor = System.Drawing.Color.Crimson, IsBoldFont = true},
                    new BaseOpenXmlExStyle() {FontSize = 20, FontName = "Calibri", BorderColor = System.Drawing.Color.Red}
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

            using var writer = new OpenXmlWriterEx(ws_part, styles);

            writer.WriteStartElement(new Worksheet());

            #region Надстройка страницы - кнопки группировки сверху

            writer.SetGrouping(false, false);

            #endregion

            #region Установка ширины колонок

            //Установка размеров колонок
            var width_setting = new List<WidthOpenXmlEx>
            {
               new (1, 2, 7),
               new (3, 3, 11),
               new (4, 12, 9.5),
               new (13, 13, 17),
               new (14, 14, 40),
               new (15, 16, 15),
               new (18, 20, 15)
            };
            writer.SetWidth(width_setting);

            #endregion

            writer.WriteStartElement(new SheetData());

            #region 1 лист


            var (key, value) = writer.FindStyleOrDefault(
                new BaseOpenXmlExStyle()
                {
                    FontColor = System.Drawing.Color.Crimson,
                    //FontSize = 20,
                    //IsBoldFont = true,
                    //LeftBorderStyle =  BorderStyleValues.Dashed,
                    //RightBorderStyle = BorderStyleValues.Dashed
                });


            #region Get cell addres

            var c = new Cell
            {
                CellReference = StringValue.FromString($"{OpenXmlExHelper.GetColumnName(850)}{82}"),
                CellValue = new CellValue("text"),
            };
            var t = OpenXmlExHelper.GetCellAddress(c);

            #endregion

            writer.AddRow(3,0,false,true);

            writer.AddCell("Test", 1, 3, 0);
            writer.AddCell("Test", 7, 3, 0);
            writer.CloseRow(3);
            writer.AddRow(4, 0, false, true);
            writer.AddCell("Test", 4, 4, 1);
            writer.AddCell("Test", 5, 4, 2);
            writer.AddCell("Test", 6, 4, 3);

            writer.AddCell("Test", 7, 4, 3);
            writer.MergeCells(6, 3, 10, 5);
            writer.CloseRow(4);
            writer.WriteEndElement(); //end of SheetData

            #region Секция настроек

            //Секция с фильтром (нужно утвердить на листе)
            writer.SetFilter(sheet_name, 1, 20, 3, 30);

            //Секция с объединенными ячейками должна быть в конце перед закрытием секции WorkSheet
            writer.SetMergedList();

            #endregion

            writer.WriteEndElement(); //end of worksheet

            #endregion

        }
    }

}
