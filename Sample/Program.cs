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
            //using var stream = File.OpenWrite(FileName);
            // create the workbook
            var workbook_part = document.AddWorkbookPart();
            var wbsp = workbook_part.AddNewPart<WorkbookStylesPart>();
            var styles = OpenXmlEx.OpenXmlEx.GetStyles(
                new List<OpenXmlExStyle>()
                {
                    new OpenXmlExStyle() {FontColor = Color.Crimson, IsBoldFont = true},
                    new OpenXmlExStyle() {FontSize = 20, FontName = "Calibri", BorderColor = Color.Red}
                });

            wbsp.Stylesheet = styles.Styles;
            wbsp.Stylesheet.Save();

            var writer = new OpenXmlEx.OpenXmlEx(wbsp, styles);

            writer.WriteStartElement(new Workbook());
            writer.WriteStartElement(new Sheets());
            var first_sheet_name = "Faults";
            var worksheet_part_1 = workbook_part.AddNewPart<WorksheetPart>();
            writer.WriteElement(new Sheet { Id = workbook_part.GetIdOfPart(worksheet_part_1), SheetId = 1, Name = first_sheet_name });
            
            var mer_list = new List<MergeCell>();

            #region 1 лист

            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());


            var test_cell_style = writer.FindStyleOrDefault(
                new OpenXmlExStyle()
                {
                    FontColor = Color.Crimson,
                    //FontSize = 20,
                    //IsBoldFont = true,
                    //LeftBorderStyle =  BorderStyleValues.Dashed,
                    //RightBorderStyle = BorderStyleValues.Dashed
                });

            writer.Add("Test",3,3, test_cell_style);

            writer.WriteEndElement(); //end of SheetData
            writer.WriteEndElement(); //end of worksheet
            writer.WriteEndElement(); //end of Sheets
            writer.WriteEndElement(); //end of Workbook



            writer.Close();
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
