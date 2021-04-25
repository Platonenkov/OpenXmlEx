using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlEx;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlEx.Errors;
using OpenXmlEx.Errors.Actions;
using OpenXmlEx.Errors.Cells;
using OpenXmlEx.Errors.Rows;
using OpenXmlEx.Styles;
using OpenXmlEx.Styles.Base;
using OpenXmlEx.SubClasses;
using OpenXmlExTests;
using Color = System.Drawing.Color;

namespace OpenXmlEx.Tests
{
    [TestClass]
    public class OpenXmlWriterExTests
    {

        [TestMethod()]
        public void SetGrouping_Secondary_Test() =>
            Assert.ThrowsException<GroupingException>(() =>
            {
                using var writer = BaseTestData.GetBaseXmlWriterTestData();
                writer.WriteStartElement(new Worksheet());
                writer.SetGrouping();
                writer.SetGrouping();
            }, "Ошибка валидации повторного добавления группировки");

        [TestMethod]
        public void SetGrouping_BeforeStartNewSheet_Test() =>
            Assert.ThrowsException<GroupingException>(() =>
            {
                using var writer = BaseTestData.GetBaseXmlWriterTestData();
                writer.SetGrouping();
            }, "Ошибка валидации добавления группировки до начала листа");

        [TestMethod()]
        public void SetWidth_Secondary_Test() =>
            Assert.ThrowsException<SetWidthException>(() =>
            {
                using var writer = BaseTestData.GetBaseXmlWriterTestData();
                writer.WriteStartElement(new Worksheet());
                writer.SetWidth(Enumerable.Empty<WidthOpenXmlEx>());
                writer.SetWidth(Enumerable.Empty<WidthOpenXmlEx>());
            }, "Ошибка валидации повторного добавления размера колонок");

        [TestMethod()]
        public void SetWidth_BeforeStartNewSheet_Test() =>
            Assert.ThrowsException<SetWidthException>(() =>
            {
                using var writer = BaseTestData.GetBaseXmlWriterTestData();
                writer.SetWidth(Enumerable.Empty<WidthOpenXmlEx>());
            }, "Ошибка валидации повторного добавления размера колонок");

        [TestMethod()]
        public void AddCell_AddressValueZeroError_Test() =>
            Assert.ThrowsException<ArgumentException>(() =>
            {
                using var writer = BaseTestData.GetXmlWriterSheetTestData();
                writer.AddRow(1);
                writer.AddCell("", 0, 1);
            });

        [TestMethod()]
        public void AddCell_ReWrite_Test() =>
            Assert.ThrowsException<CellException>(() =>
            {
                using var writer = BaseTestData.GetXmlWriterSheetTestData();
                writer.AddRow(1);
                writer.AddCell("", 1, 1);
                writer.AddCell("", 1, 1, CanReWrite: false);
            });

        [TestMethod()]
        public void AddCell_RowClosed_Test() =>
            Assert.ThrowsException<RowNotOpenException>(() =>
            {
                using var writer = BaseTestData.GetXmlWriterSheetTestData();
                writer.AddRow(1);
                writer.CloseRow(1);
                writer.AddCell("", 1, 1);
            });

        [TestMethod()]
        public void AddCell_WriteToAboveCell_Test()
        {
            Assert.ThrowsException<CellException>(() =>
            {
                using var writer = BaseTestData.GetXmlWriterSheetTestData();
                writer.AddRow(1);
                writer.AddCell("", 2, 1);
                writer.AddCell("", 1, 1);
            });
        }
        [TestMethod()]
        public void AddCell_RowNotAdded_Test() =>
            Assert.ThrowsException<RowException>(() =>
            {
                using var writer = BaseTestData.GetXmlWriterSheetTestData();
                writer.AddCell("", 1, 1);
            });

        [TestMethod()]
        public void AddRow_OrderByRows_Test()
        {
            Assert.ThrowsException<RowException>(() =>
            {
                using var writer = BaseTestData.GetXmlWriterSheetTestData();
                writer.AddRow(2);
            });
        }
        [TestMethod()]
        public void AddRow_ClosePrevious_Test()
        {
            Assert.ThrowsException<RowNotClosedException>(() =>
            {
                using var writer = BaseTestData.GetXmlWriterSheetTestData();
                writer.AddRow(1);
                writer.AddRow(2);
            });
        }

        [TestMethod()]
        public void CloseRow_NotAdded_Test()
        {
            Assert.ThrowsException<RowException>(() =>
            {
                using var writer = BaseTestData.GetXmlWriterSheetTestData();
                writer.CloseRow(1);
            });
        }
        [TestMethod()]
        public void CloseRow_IsClosedTest()
        {
            Assert.ThrowsException<RowNotOpenException>(() =>
            {
                using var writer = BaseTestData.GetXmlWriterSheetTestData();
                writer.AddRow(1);
                writer.CloseRow(1);
                writer.CloseRow(1);
            });
        }

        [TestMethod()]
        public void SetFilter_NoSheetName_Test()
        {
            Assert.ThrowsException<FilterException>(() =>
            {
                using var writer = BaseTestData.GetXmlWriterSheetTestData();
                writer.SetFilter(1,1,1,1);
            });
        }
        [TestMethod()]
        public void SetFilter_SecondarySet_Test()
        {
            Assert.ThrowsException<FilterException>(() =>
            {
                using var writer = BaseTestData.GetXmlWriterSheetTestData();
                writer.SetFilter(1,1,1,1,"test_sheet");
                writer.SetFilter(1, 2, 1, 2, "test_sheet");
            });
        }

        [TestMethod()]
        public void MergeCellsTest()
        {
            Assert.ThrowsException<MergeCellException>(() =>
            {
                using var writer = BaseTestData.GetXmlWriterSheetTestData();
                writer.MergeCells(1,1,3,3);
                writer.MergeCells(1,2,4,4);
            });
        }

        [TestMethod()]
        public void FindStyleOrDefaultTest()
        {
            var style = new BaseOpenXmlExStyle()
            {
                FontColor = Color.Red,
                BorderColor = Color.Blue,
                BottomBorderStyle = BorderStyleValues.DashDot,
                FillColor = Color.Gray,
                FillPattern = PatternValues.DarkDown,
                FontName = "Tahoma",
                FontSize = 14,
                HorizontalAlignment = HorizontalAlignmentValues.Center,
                IsBoldFont = true,
                IsItalicFont = true,
                LeftBorderStyle = BorderStyleValues.Dashed,
                RightBorderStyle = BorderStyleValues.Double,
                TopBorderStyle = BorderStyleValues.Medium,
                VerticalAlignment = VerticalAlignmentValues.Top,
                WrapText = true
            };

            using var writer = new OpenXmlWriterEx(BaseTestData.GetBaseSpreadsheetDocument(), new OpenXmlExStyles(new[] {style}));
            Assert.AreEqual(2U, writer.FindStyleOrDefault(new BaseOpenXmlExStyle()
            {
                FontColor = Color.Red,
                BorderColor = Color.Blue,
                BottomBorderStyle = BorderStyleValues.DashDot,
                FillColor = Color.Gray,
                FillPattern = PatternValues.DarkDown,
                FontName = "Tahoma",
                FontSize = 14,
                HorizontalAlignment = HorizontalAlignmentValues.Center,
                IsBoldFont = true,
                IsItalicFont = true,
                LeftBorderStyle = BorderStyleValues.Dashed,
                RightBorderStyle = BorderStyleValues.Double,
                TopBorderStyle = BorderStyleValues.Medium,
                VerticalAlignment = VerticalAlignmentValues.Top,
                WrapText = true
            }).Key);
        }
    }
}