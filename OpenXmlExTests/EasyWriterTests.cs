using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlEx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlEx.Errors;
using OpenXmlEx.Errors.Actions;
using OpenXmlEx.Errors.Sheets;
using OpenXmlEx.Styles;
using OpenXmlEx.Styles.Base;
using OpenXmlEx.SubClasses;
using OpenXmlExTests;
using Color = System.Drawing.Color;

namespace OpenXmlEx.Tests
{
    [TestClass()]
    public class EasyWriterTests
    {
        [TestMethod()]
        public void AddNewSheetTest() =>
            Assert.ThrowsException<SheetException>(
                () =>
                {
                    using var writer = new EasyWriter(BaseTestData.NewName);
                    writer.AddNewSheet("new");
                    writer.AddNewSheet("new");
                });

        [TestMethod()]
        public void SetGrouping_Secondary_Test() =>
            Assert.ThrowsException<GroupingException>(() =>
            {
                using var writer = new EasyWriter(BaseTestData.NewName);
                writer.AddNewSheet();
                writer.SetGrouping();
                writer.SetGrouping();
            }, "Ошибка вилидации повторного добавления группировки");

        [TestMethod]
        public void SetGrouping_BeforeStartNewSheet_Test() =>
            Assert.ThrowsException<WriterException>(() =>
            {
                using var writer = new EasyWriter(BaseTestData.NewName);
                writer.SetGrouping();
            }, "Ошибка вилидации добавления группировки до начала листа");

        [TestMethod()]
        public void SetWidth_Secondary_Test() =>
            Assert.ThrowsException<SetWidthException>(() =>
            {
                using var writer = new EasyWriter(BaseTestData.NewName);
                writer.AddNewSheet();
                writer.SetWidth(Enumerable.Empty<WidthOpenXmlEx>());
                writer.SetWidth(Enumerable.Empty<WidthOpenXmlEx>());
            }, "Ошибка валидации повторного добавления размера колонок");
        [TestMethod()]
        public void SetWidth_BeforeStartNewSheet_Test() =>
            Assert.ThrowsException<WriterException>(() =>
            {
                using var writer = new EasyWriter(BaseTestData.NewName);
                writer.SetWidth(Enumerable.Empty<WidthOpenXmlEx>());
            });

        [TestMethod()]
        public void AddRow_BeforeStartNewSheet_Test() =>
            Assert.ThrowsException<WriterException>(() =>
            {
                using var writer = new EasyWriter(BaseTestData.NewName);
                writer.AddRow(1);
            });

        [TestMethod()]
        public void CloseRow_BeforeStartNewSheet_Test() =>
            Assert.ThrowsException<WriterException>(() =>
            {
                using var writer = new EasyWriter(BaseTestData.NewName);
                writer.CloseRow(1);
            });

        [TestMethod()]
        public void AddCell_BeforeStartNewSheet_Test() =>
            Assert.ThrowsException<WriterException>(() =>
            {
                using var writer = new EasyWriter(BaseTestData.NewName);
                writer.AddCell("",1,1);
            });

        [TestMethod()]
        public void MergeCellsTest() =>
            Assert.ThrowsException<MergeCellException>(() =>
            {
                using var writer = new EasyWriter(BaseTestData.NewName);
                writer.MergeCells(1, 1, 1,2);
            });

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

            using var writer = new EasyWriter(BaseTestData.NewName, new OpenXmlExStyles(new[] { style }));
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

        [TestMethod()]
        public void CloseTest() =>
            Assert.ThrowsException<ObjectDisposedException>(() =>
            {
                var writer = new EasyWriter(BaseTestData.NewName);
                writer.Close();
                writer.Close();
            });
    }
}