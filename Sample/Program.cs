using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net.Http;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32.SafeHandles;
using OpenXmlEx;
namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            var FileName = @"C:\Users\Evril\Downloads\Test\test.xlsx";

            using var spread_sheet = SpreadsheetDocument.Create(FileName, SpreadsheetDocumentType.Workbook);
            //using var stream = File.OpenWrite(FileName);
            // create the workbook
            var workbook_part = spread_sheet.AddWorkbookPart();

            var wbsp = workbook_part.AddNewPart<WorkbookStylesPart>();
            var writer = new OpenXmlEx.OpenXmlEx(wbsp);
            writer.InitStyles(new[] {"Times New Roman", "Calibri", "Arial"}, new []{System.Drawing.Color.BlueViolet, System.Drawing.Color.Crimson });


            var styles = OpenXmlEx.Styles.OpenXmlExStyles.GetStylesheet();
            writer.WriteElement(styles);

            writer.WriteStartElement(new Workbook());
            writer.WriteStartElement(new Sheets());
            var first_sheet_name = "Faults";
            var worksheet_part_1 = workbook_part.AddNewPart<WorksheetPart>();
            writer.WriteElement(new Sheet { Id = workbook_part.GetIdOfPart(worksheet_part_1), SheetId = 1, Name = first_sheet_name });


            var mer_list = new List<MergeCell>();

            #region 1 лист

            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());


            writer.WriteEndElement(); //end of SheetData
            writer.WriteEndElement(); //end of worksheet
            writer.WriteEndElement(); //end of Sheets
            writer.WriteEndElement(); //end of Workbook



            writer.Close();
            #endregion
        }

        private void Test(string FileName)
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
            using (var writer = OpenXmlWriter.Create(worksheet_part_1))
            {

            }


        }
    }
}
