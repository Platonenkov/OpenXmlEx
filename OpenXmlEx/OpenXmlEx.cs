using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Bold = DocumentFormat.OpenXml.Spreadsheet.Bold;
using Border = DocumentFormat.OpenXml.Spreadsheet.Border;
using BottomBorder = DocumentFormat.OpenXml.Spreadsheet.BottomBorder;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using Column = DocumentFormat.OpenXml.Spreadsheet.Column;
using Columns = DocumentFormat.OpenXml.Spreadsheet.Columns;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;
using Fonts = DocumentFormat.OpenXml.Spreadsheet.Fonts;
using FontSize = DocumentFormat.OpenXml.Spreadsheet.FontSize;
using HorizontalAlignmentValues = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues;
using LeftBorder = DocumentFormat.OpenXml.Spreadsheet.LeftBorder;
using RightBorder = DocumentFormat.OpenXml.Spreadsheet.RightBorder;
using TopBorder = DocumentFormat.OpenXml.Spreadsheet.TopBorder;
using VerticalAlignmentValues = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues;
using System.Drawing.Text;
using OpenXmlEx.Styles;
using OpenXmlEx.Styles.Base;

namespace OpenXmlEx
{
    public class OpenXmlEx : OpenXmlPartWriter
    {
        /// <inheritdoc />
        public OpenXmlEx(
            OpenXmlPart OpenXmlPart,
            IEnumerable<string> FontNames, IEnumerable<uint> FontSizes, IEnumerable<System.Drawing.Color> Colors)
            : base(OpenXmlPart) => InitStyles(FontNames, FontSizes, Colors);

        /// <inheritdoc />
        public OpenXmlEx(OpenXmlPart OpenXmlPart, Encoding encoding,
            IEnumerable<string> FontNames, IEnumerable<uint> FontSizes, IEnumerable<System.Drawing.Color> Colors)
            : base(OpenXmlPart, encoding) => InitStyles(FontNames, FontSizes, Colors);

        /// <inheritdoc />
        public OpenXmlEx(Stream PartStream,
            IEnumerable<string> FontNames, IEnumerable<uint> FontSizes, IEnumerable<System.Drawing.Color> Colors)
            : base(PartStream) => InitStyles(FontNames, FontSizes, Colors);

        /// <inheritdoc />
        public OpenXmlEx(Stream PartStream, Encoding encoding,
            IEnumerable<string> FontNames, IEnumerable<uint> FontSizes, IEnumerable<System.Drawing.Color> Colors)
            : base(PartStream, encoding) => InitStyles(FontNames, FontSizes, Colors);

        public OpenXmlExStyles Style { get; private set; }
        private void InitStyles(IEnumerable<string> FontNames, IEnumerable<uint> FontSizes, IEnumerable<System.Drawing.Color> Colors)
            => Style = new OpenXmlExStyles(FontNames, FontSizes, Colors);


        #region Extensions

        /// <summary> Устанавливает тип группировки для строк и столбцов </summary>
        /// <param name="SummaryBelow">группировать сверху (false - сверху, true - снизу)</param>
        /// <param name="SummaryRight">группировать справа (false - справа, true - слева)</param>
        public void SetGrouping(bool SummaryBelow = false, bool SummaryRight = false)
        {
            #region Надстройка страницы - кнопки группировки сверху

            WriteStartElement(new SheetProperties());
            WriteElement(new OutlineProperties { SummaryBelow = SummaryBelow, SummaryRight = SummaryRight });
            WriteEndElement();

            #endregion
        }

        /// <summary> Устанавливает параметры столбцов </summary>
        /// <param name="Settings">список надстроек для листа</param>
        public void SetWidth(IEnumerable<(uint First, uint Last, double width)> Settings)
        {

            #region Установка ширины колонок

            WriteStartElement(new Columns());
            foreach (var (first, last, width) in Settings)
                WriteElement(new Column { Min = first, Max = last, Width = width });
            WriteEndElement();

            #endregion

        }

        /// <summary> Устанавливает параметры столбцов </summary>
        /// <param name="first">начальная колонка</param>
        /// <param name="last">конечная колонка</param>
        /// <param name="width">ширина по умолчанию</param>
        public void SetWidth(uint first, uint last, double width = 12)
        {

            #region Установка ширины колонок

            WriteStartElement(new Columns());
            WriteElement(new Column { Min = first, Max = last, Width = width });
            WriteEndElement();

            #endregion

        }

        /// <summary> Добавляет значение в ячейку документа </summary>
        /// <param name="text">текст для записи</param>
        /// <param name="CellNum">номер колонки</param>
        /// <param name="RowNum">номер строки</param>
        /// <param name="StyleIndex">силь</param>
        /// <param name="Type">тип данных</param>
        public void Add(string text, int CellNum, uint RowNum, uint StyleIndex = 0, CellValues Type = CellValues.String) =>
            WriteElement(
                new Cell
                {
                    CellReference = StringValue.FromString($"{GetColumnName(CellNum)}{RowNum}"),
                    CellValue = new CellValue(text),
                    DataType = Type,
                    StyleIndex = StyleIndex
                });

        /// <summary> Печатает ячейки с одинаковым значением и стилем со столбца по столбец в одной и той же строке</summary>
        /// <param name="FirstColumn">колонка с которой начали печать</param>
        /// <param name="LastPrintColumn">последняя напечатанная колонка</param>
        /// <param name="RowNumber">строка в которой идёт печать</param>
        /// <param name="Style">стиль ячейки</param>
        public void PrintEmptyCells(int FirstColumn, int LastPrintColumn, uint RowNumber, uint Style = 0) =>
            PrintCells(FirstColumn, LastPrintColumn, RowNumber, string.Empty, Style);

        /// <summary> Печатает ячейки с одинаковым значением и стилем со столбца по столбец в одной и той же строке</summary>
        /// <param name="FirstColumn">колонка с которой начали печать</param>
        /// <param name="LastPrintColumn">последняя напечатанная колонка</param>
        /// <param name="RowNumber">строка в которой идёт печать</param>
        /// <param name="Value">значение для печати (по умолчанию string.Empty)</param>
        /// <param name="Style">стиль ячейки</param>
        public void PrintCells(int FirstColumn, int LastPrintColumn, uint RowNumber, string Value, uint Style)
        {
            foreach (var i in Enumerable.Range(FirstColumn, LastPrintColumn - FirstColumn + 1))
                WriteElement(
                    new Cell
                    {
                        CellReference = StringValue.FromString($"{GetColumnName(i)}{RowNumber}"),
                        CellValue = new CellValue(Value),
                        DataType = CellValues.String,
                        StyleIndex = Style
                    });
        }

        /// <summary> Устанавливает фильтр на колонки (ставить в конце листа перед закрытием)</summary>
        /// Позиционировать обязательно в конце страницы после закрытия блока SheetData
        /// перед закрытием блока WorkSheet и MergedList
        /// <param name="ListName">Имя листа</param>
        /// <param name="FirstColumn">первая колонка</param>
        /// <param name="LastColumn">последняя колонка</param>
        /// <param name="FirstRow">первая строка</param>
        /// <param name="LastRow">последня строка</param>
        public void SetFilter(string ListName, uint FirstColumn, uint LastColumn, uint FirstRow, uint? LastRow = null)
        {
            WriteElement(new AutoFilter { Reference = $"{GetColumnName(FirstColumn)}{FirstRow}:{GetColumnName(LastColumn)}{LastRow ?? FirstRow}" });
            // не забыть в конце листа утвердить в конце листа
            ApprovalFilter(ListName, FirstColumn, LastColumn, FirstRow, LastRow ?? FirstRow);
        }

        /// <summary> Утверждение секции фильтра на листе </summary>
        /// <param name="ListName">Имя листа</param>
        /// <param name="FirstColumn">первая колонка</param>
        /// <param name="LastColumn">последняя колонка</param>
        /// <param name="FirstRow">первая строка</param>
        /// <param name="LastRow">последня строка</param>
        private void ApprovalFilter(string ListName, uint FirstColumn, uint LastColumn, uint FirstRow, uint LastRow)
        {
            //Секция с фильтром часть-2 - подтвердение принадлежности к листу
            WriteStartElement(new DefinedNames());
            WriteElement(
                new DefinedName
                {
                    Name = "_xlnm._FilterDatabase",
                    LocalSheetId = 0U,
                    Hidden = true,
                    Text = $"{ListName}!${GetColumnName(FirstColumn)}${FirstRow}:${GetColumnName(LastColumn)}${LastRow}"
                });
            WriteEndElement(); //Filter
        }

        /// <summary>
        /// Устанавливает объединенные ячейки на листе
        /// Позиционировать обязательно в конце страницы после закрытия блока SheetData
        /// после блока фильтров но до закрытия блока WorkSheet
        /// </summary>
        /// <param name="MergedCells">перечень объединенных ячеек</param>
        public void SetMergedList(IEnumerable<MergeCell> MergedCells)
        {
            WriteStartElement(new MergeCells());
            foreach (var mer in MergedCells) WriteElement(mer);
            WriteEndElement();
        }


        #endregion

        #region Helper
        /// <summary> Метод генерирует стили для ячеек </summary>
        /// <returns></returns>
        public Stylesheet GenerateStyleSheet() =>
            new Stylesheet(
                new Fonts(
                    new Font( // Стиль под номером 0 - Шрифт по умолчанию.
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Times New Roman" }),
                    new Font( // Стиль под номером 1 - Жирный шрифт Times New Roman.
                        new Bold(),
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Times New Roman" }),
                    new Font( // Стиль под номером 2 - Шрифт Times New Roman размером 14.
                        new FontSize() { Val = 14 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Times New Roman" }),
                    new Font( // Стиль под номером 3 - Calibri Шрифт по умолчанию.
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font( // Стиль под номером 4 - Жирный шрифт Calibri.
                        new Bold(),
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font( // Стиль под номером 5 - Шрифт Calibri размером 14.
                        new FontSize() { Val = 14 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font( // Стиль под номером 6 - Жирный шрифт Calibri 11.
                        new Bold(),
                        new FontSize() { Val = 10 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" })
                ),
                new Fills(
                    new Fill( // Стиль под номером 0 - Заполнение ячейки по умолчанию.
                        new PatternFill() { PatternType = PatternValues.None }),
                    // Стиль под номером 1 - Заполнение ячейки серыми точками (хз как но 1 стиль всегда серые точки не важно от настроек
                    new Fill(
                        new PatternFill(
                                new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFA500" } }
                            )
                        { PatternType = PatternValues.Solid }),
                    new Fill( // Стиль под номером 2 - Заполнение ячейки Оранжеваым цветом
                        new PatternFill(
                                new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FB793D" } }
                            )
                        { PatternType = PatternValues.Solid }),
                    new Fill( // Стиль под номером 3 - Заполнение ячейки серым
                        new PatternFill(
                                new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "CCC2A6" } }
                            )
                        { PatternType = PatternValues.Solid }),
                    new Fill( // Стиль под номером 4 - Заполнение ячейки синим 
                        new PatternFill(
                                new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "8FBCE6" } }
                            )
                        { PatternType = PatternValues.Solid }),
                    new Fill( // Стиль под номером 5 - Заполнение ячейки светло зелёным 
                        new PatternFill(
                                new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "10B23C" } }
                            )
                        { PatternType = PatternValues.Solid })
                )
               ,
                new Borders(
                    new Border( // Стиль под номером 0 - Грани.
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()),
                    new Border( // Стиль под номером 1 - Грани
                        new LeftBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Medium },
                        new RightBorder(
                                new Color() { Indexed = (UInt32Value)64U }
                            )
                        { Style = BorderStyleValues.Medium },
                        new TopBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Medium },
                        new BottomBorder(
                                new Color() { Indexed = (UInt32Value)64U }
                            )
                        { Style = BorderStyleValues.Medium },
                        new DiagonalBorder()),
                    new Border( // Стиль под номером 2 - Грани.
                        new LeftBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Thin },
                        new RightBorder(
                                new Color() { Indexed = (UInt32Value)64U }
                            )
                        { Style = BorderStyleValues.Thin },
                        new TopBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Thin },
                        new BottomBorder(
                                new Color() { Indexed = (UInt32Value)64U }
                            )
                        { Style = BorderStyleValues.Thin },
                        new DiagonalBorder()),
                    new Border( // Стиль под номером 3 - Dotted|Thin.
                        new LeftBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Thin },
                        new RightBorder(
                                new Color() { Indexed = (UInt32Value)64U }
                            )
                        { Style = BorderStyleValues.Thin },
                        new TopBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Dotted },
                        new BottomBorder(
                                new Color() { Indexed = (UInt32Value)64U }
                            )
                        { Style = BorderStyleValues.Dotted },
                        new DiagonalBorder()),
                    new Border( // Стиль под номером 4 - Dotted.
                        new LeftBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Dotted },
                        new RightBorder(
                                new Color() { Indexed = (UInt32Value)64U }
                            )
                        { Style = BorderStyleValues.Dotted },
                        new TopBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Dotted },
                        new BottomBorder(
                                new Color() { Indexed = (UInt32Value)64U }
                            )
                        { Style = BorderStyleValues.Dotted },
                        new DiagonalBorder())
                ),
                new CellFormats(
                    // Стиль под номером 0 - The default cell style.  (по умолчанию)
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 },
                    // Стиль под номером 1 - заголовки по центру оранж в рамке 
                    new CellFormat(
                            new Alignment()
                            {
                                Horizontal = HorizontalAlignmentValues.Center,
                                Vertical = VerticalAlignmentValues.Center,
                                WrapText = true
                            })
                    { FontId = 1, FillId = 2, BorderId = 1, ApplyFont = true },
                    // Стиль под номером 2 - заголовки по центру серые в рамке
                    new CellFormat(
                            new Alignment()
                            {
                                Horizontal = HorizontalAlignmentValues.Center,
                                Vertical = VerticalAlignmentValues.Center,
                                WrapText = true
                            })
                    { FontId = 1, FillId = 3, BorderId = 1, ApplyFont = true },
                    // Стиль под номером 3 - Оранж с фильром
                    new CellFormat(
                            new Alignment()
                            {
                                Horizontal = HorizontalAlignmentValues.Center,
                                Vertical = VerticalAlignmentValues.Center,
                                WrapText = true
                            })
                    { FontId = 1, FillId = 2, BorderId = 1, ApplyFont = true },
                    // Стиль под номером 4 - Серый с фильтром
                    new CellFormat(
                            new Alignment()
                            {
                                Horizontal = HorizontalAlignmentValues.Center,
                                Vertical = VerticalAlignmentValues.Center,
                                WrapText = true
                            })
                    { FontId = 1, FillId = 3, BorderId = 1, ApplyFont = true },
                    // Стиль под номером 5 - заголовки по центру синие в рамке
                    new CellFormat(
                            new Alignment()
                            {
                                Horizontal = HorizontalAlignmentValues.Center,
                                Vertical = VerticalAlignmentValues.Center,
                                WrapText = true
                            })
                    { FontId = 1, FillId = 4, BorderId = 1, ApplyFont = true },
                    // Стиль под номером 6 - Синий с фильтром
                    new CellFormat(
                            new Alignment()
                            {
                                Horizontal = HorizontalAlignmentValues.Center,
                                Vertical = VerticalAlignmentValues.Center,
                                WrapText = true
                            })
                    { FontId = 1, FillId = 4, BorderId = 1, ApplyFont = true },
                    // Стиль под номером 7 - рамка.
                    new CellFormat(new Alignment() { WrapText = true })
                    { FontId = 0, FillId = 0, BorderId = 2, ApplyFont = true },
                    // Стиль под номером 8 - рамка и текст по центру
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 3, ApplyFont = true },
                    // Стиль под номером 9 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 3, ApplyFont = true },
                    // Стиль под номером 10 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 3, ApplyFont = true },
                    // Стиль под номером 11 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Top, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 3, ApplyFont = true },
                    // Стиль под номером 12 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 3, ApplyFont = true },
                    // Стиль под номером 13 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Top, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 3, ApplyFont = true },
                    // Стиль под номером 14 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 3, ApplyFont = true },
                    // Стиль под номером 15 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Bottom, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 3, ApplyFont = true },
                    // Стиль под номером 16 - рамка и текст по центру
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Bottom, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 3, ApplyFont = true },
                    // Стиль под номером 17 - рамка и текст по центру
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 4, ApplyFont = true },
                    // Стиль под номером 18 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 4, ApplyFont = true },
                    // Стиль под номером 19 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 4, ApplyFont = true },
                    // Стиль под номером 20 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Top, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 4, ApplyFont = true },
                    // Стиль под номером 21 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 4, ApplyFont = true },
                    // Стиль под номером 22 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Top, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 4, ApplyFont = true },
                    // Стиль под номером 23 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 4, ApplyFont = true },
                    // Стиль под номером 24 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Bottom, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 4, ApplyFont = true },
                    // Стиль под номером 25 - рамка.
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Bottom, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 4, ApplyFont = true },
                    // Стиль под номером 26 - заголовки по центру синие в рамке Dash
                    new CellFormat(
                            new Alignment()
                            {
                                Horizontal = HorizontalAlignmentValues.Center,
                                Vertical = VerticalAlignmentValues.Center,
                                WrapText = true
                            })
                    { FontId = 6, FillId = 4, BorderId = 3, ApplyFont = true },
                    // Стиль под номером 27 - заголовки по центру серые в рамке Dash
                    new CellFormat(
                            new Alignment()
                            {
                                Horizontal = HorizontalAlignmentValues.Center,
                                Vertical = VerticalAlignmentValues.Center,
                                WrapText = true
                            })
                    { FontId = 6, FillId = 3, BorderId = 3, ApplyFont = true },
                    // Стиль под номером 28 - заголовки по центру Оранж в рамке Dash
                    new CellFormat(
                            new Alignment()
                            {
                                Horizontal = HorizontalAlignmentValues.Center,
                                Vertical = VerticalAlignmentValues.Center,
                                WrapText = true
                            })
                    { FontId = 6, FillId = 2, BorderId = 3, ApplyFont = true },
                    // Стиль под номером 29 - заголовки по центру Зелёный в рамке Dash
                    new CellFormat(
                            new Alignment()
                            {
                                Horizontal = HorizontalAlignmentValues.Center,
                                Vertical = VerticalAlignmentValues.Center,
                                WrapText = true
                            })
                    { FontId = 6, FillId = 5, BorderId = 3, ApplyFont = true },
                    // Стиль под номером 30 - заголовки по центру чёрный в точку в рамке Dash
                    new CellFormat(
                            new Alignment()
                            {
                                Horizontal = HorizontalAlignmentValues.Center,
                                Vertical = VerticalAlignmentValues.Center,
                                WrapText = true
                            })
                    { FontId = 6, FillId = 1, BorderId = 3, ApplyFont = true }

                )
            );

        /// <summary> Словарь имен колонок excel </summary>
        private readonly Dictionary<int, string> _Columns = new(676);


        /// <summary> Возвращает строковое имя колонки по номеру (1 - А, 2 - В) </summary>
        /// <param name="index">номер колонки</param>
        /// <returns></returns>
        public string GetColumnName(uint index) => GetColumnName((int)index);

        /// <summary> Возвращает строковое имя колонки по номеру (1 - А, 2 - В) </summary>
        /// <param name="index">номер колонки</param>
        /// <returns></returns>
        public string GetColumnName(int index)
        {
            var int_col = index - 1;
            if (_Columns.ContainsKey(int_col)) return _Columns[int_col];
            var int_first_letter = ((int_col) / 676) + 64;
            var int_second_letter = ((int_col % 676) / 26) + 64;
            var int_third_letter = (int_col % 26) + 65;
            var FirstLetter = (int_first_letter > 64) ? (char)int_first_letter : ' ';
            var SecondLetter = (int_second_letter > 64) ? (char)int_second_letter : ' ';
            var ThirdLetter = (char)int_third_letter;
            var s = string.Concat(FirstLetter, SecondLetter, ThirdLetter).Trim();
            _Columns.Add(int_col, s);
            return s;
        }

        #region MergedCell

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то таже что и начало)</param>
        /// <returns></returns>
        public MergeCell MergeCells(int StartCell, int StartRow, int EndCell, int? EndRow = null)
            => new() { Reference = new StringValue($"{GetColumnName(StartCell)}{StartRow}:{GetColumnName(EndCell)}{EndRow ?? StartRow}") };

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то таже что и начало)</param>
        /// <returns></returns>
        public MergeCell MergeCells(int StartCell, uint StartRow, int EndCell, uint? EndRow = null)
            => new() { Reference = new StringValue($"{GetColumnName(StartCell)}{StartRow}:{GetColumnName(EndCell)}{EndRow ?? StartRow}") };

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то таже что и начало)</param>
        /// <returns></returns>
        public MergeCell MergeCells(uint StartCell, uint StartRow, uint EndCell, uint? EndRow = null)
            => new() { Reference = new StringValue($"{GetColumnName(StartCell)}{StartRow}:{GetColumnName(EndCell)}{EndRow ?? StartRow}") };

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то таже что и начало)</param>
        /// <returns></returns>
        public MergeCell MergeCells(uint StartCell, int StartRow, uint EndCell, int? EndRow = null)
            => new() { Reference = new StringValue($"{GetColumnName(StartCell)}{StartRow}:{GetColumnName(EndCell)}{EndRow ?? StartRow}") };

        #endregion

        /// <summary>
        /// Создаёт запись о группировке для writer
        /// </summary>
        /// <param name="lvl">уровень группы</param>
        /// <returns></returns>
        public static OpenXmlAttribute[] GetCollapsedAttributes(int lvl = 1) => lvl == 0
            ? Array.Empty<OpenXmlAttribute>()
            : new[] { new OpenXmlAttribute("outlineLevel", string.Empty, $"{lvl}"), new OpenXmlAttribute("hidden", string.Empty, $"{lvl}") };


        #endregion

        #region Style Comparer

        public uint FindStyleOrDefault(OpenXmlExStyle style)
        {
            IEnumerable<KeyValuePair<uint, OpenXmlExStyleCell>> values = Array.Empty<KeyValuePair<uint, OpenXmlExStyleCell>>();

            #region Заливка

            if(style.FillColor!=null)
                values = Style.CellsStyles.Where(
                    s=> s.Value.FillStyle.Value.FillColor.Key == style.FillColor.Value).AsEnumerable();
            if(style.FillPattern!=null)
                values = values.Where(
                    s => s.Value.FillStyle.Value.FillPattern == style.FillPattern).AsEnumerable();

            #endregion

            #region Borders

            if (style.BorderColor != null)
                values = values.Where(
                    s => s.Value.BorderStyle.Value.BorderColor.Key == style.BorderColor).AsEnumerable();
            if (style.LeftBorderStyle != null)
                values = values.Where(
                    s => s.Value.BorderStyle.Value.LeftBorder.BorderStyle == style.LeftBorderStyle).AsEnumerable();
            if (style.TopBorderStyle != null)
                values = values.Where(
                    s => s.Value.BorderStyle.Value.TopBorder.BorderStyle == style.TopBorderStyle).AsEnumerable();
            if (style.RightBorderStyle != null)
                values = values.Where(
                    s => s.Value.BorderStyle.Value.RightBorder.BorderStyle == style.RightBorderStyle).AsEnumerable();
            if (style.BottomBorderStyle != null)
                values = values.Where(
                    s => s.Value.BorderStyle.Value.BottomBorder.BorderStyle == style.BottomBorderStyle).AsEnumerable();


            #endregion

            #region Шрифт

            if (style.FontSize!=null)
                values = values.Where(
                    s => s.Value.FontStyle.Value.FontSize == style.FontSize).AsEnumerable();
            if(style.FontColor != null)
                values = values.Where(
                    s => s.Value.FontStyle.Value.FontColor.Key == style.FontColor).AsEnumerable();
            if(string.IsNullOrWhiteSpace(style.FontName))
                values = values.Where(
                    s => s.Value.FontStyle.Value.FontName == style.FontName).AsEnumerable();
            if(style.IsBoldFont != null)
                values = values.Where(
                    s => s.Value.FontStyle.Value.IsBoldFont == style.IsBoldFont).AsEnumerable();
            if(style.IsItalicFont != null)
                values = values.Where(
                    s => s.Value.FontStyle.Value.IsItalicFont == style.IsItalicFont).AsEnumerable();

            #endregion

            #region Выравнивание

            if (style.WrapText != null)
                values = values.Where(
                    s => s.Value.WrapText == style.WrapText).AsEnumerable();
            if (style.HorizontalAlignment != null)
                values = values.Where(
                    s => s.Value.HorizontalAlignment == style.HorizontalAlignment).AsEnumerable();
            if (style.VerticalAlignment != null)
                values = values.Where(
                    s => s.Value.VerticalAlignment == style.VerticalAlignment).AsEnumerable();

            #endregion

            return values.FirstOrDefault().Key;
        }

        #endregion
    }
}
