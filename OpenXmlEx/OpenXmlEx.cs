using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlEx.Errors;
using Column = DocumentFormat.OpenXml.Spreadsheet.Column;
using Columns = DocumentFormat.OpenXml.Spreadsheet.Columns;
using OpenXmlEx.Styles;
using OpenXmlEx.Styles.Base;

namespace OpenXmlEx
{
    public class OpenXmlEx : OpenXmlPartWriter
    {
        public static OpenXmlExStyles GetStyles(IEnumerable<OpenXmlExStyle> styles) => new OpenXmlExStyles(styles);

        public OpenXmlExStyles Style { get; private set; }

        #region Конструкторы

        #region приоритет 1

        /// <inheritdoc />
        public OpenXmlEx(
            OpenXmlPart OpenXmlPart,
            OpenXmlExStyles styles)
            : base(OpenXmlPart) => InitStyles(styles);

        /// <inheritdoc />
        public OpenXmlEx(OpenXmlPart OpenXmlPart, Encoding encoding,
            OpenXmlExStyles styles)
            : base(OpenXmlPart, encoding) => InitStyles(styles);

        /// <inheritdoc />
        public OpenXmlEx(Stream PartStream,
            OpenXmlExStyles styles)
            : base(PartStream) => InitStyles(styles);

        /// <inheritdoc />
        public OpenXmlEx(Stream PartStream, Encoding encoding,
            OpenXmlExStyles styles)
            : base(PartStream, encoding) => InitStyles(styles);

        private void InitStyles(OpenXmlExStyles styles)
            => Style = styles;

        #endregion

        #region приоритет 2

        /// <inheritdoc />
        public OpenXmlEx(
            OpenXmlPart OpenXmlPart,
            IEnumerable<OpenXmlExStyle> styles)
            : base(OpenXmlPart) => InitStyles(styles);

        /// <inheritdoc />
        public OpenXmlEx(OpenXmlPart OpenXmlPart, Encoding encoding,
            IEnumerable<OpenXmlExStyle> styles)
            : base(OpenXmlPart, encoding) => InitStyles(styles);

        /// <inheritdoc />
        public OpenXmlEx(Stream PartStream,
            IEnumerable<OpenXmlExStyle> styles)
            : base(PartStream) => InitStyles(styles);

        /// <inheritdoc />
        public OpenXmlEx(Stream PartStream, Encoding encoding,
            IEnumerable<OpenXmlExStyle> styles)
            : base(PartStream, encoding) => InitStyles(styles);

        private void InitStyles(IEnumerable<OpenXmlExStyle> styles)
            => Style = new OpenXmlExStyles(styles);

        #endregion

        #endregion


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

        /// <summary>
        /// Список записанных ячеек, со статусом (false - open, true - close)
        /// ключ - (номер строки, номер ячейки)
        /// </summary>
        private readonly Dictionary<(uint row, uint cell), bool> _Cells = new();

        /// <summary> Добавляет значение в ячейку документа </summary>
        /// <param name="text">текст для записи</param>
        /// <param name="RowNum">номер строки</param>
        /// <param name="CellNum">номер колонки</param>
        /// <param name="StyleIndex">индекс стиля</param>
        /// <param name="Type">тип данных</param>
        /// <param name="CanReWrite">разрешить перезапись данных (иначе при повторной записи в ячейку будет генерирование ошибки)</param>
        public void AddCell(string text, uint RowNum, uint CellNum, uint StyleIndex = 0, CellValues Type = CellValues.String, bool CanReWrite = false)
        {
            //Проверка валидности номера строки или столбца (должны быть больше 0)
            if (CellNum == 0 || RowNum == 0)
            {
                throw new ArgumentException($"Address must be greater that 0, Row({RowNum}) and Cell({CellNum})");
            }
            (uint row, uint cell) key = (RowNum, CellNum);

            //Проверка на перезапись данных
            if (_Cells.TryGetValue(key, out var _) && !CanReWrite)
            {
                throw new CellException("Re-writing data to a cell", RowNum, CellNum, GetColumnName(CellNum));
            }
            // проверка на то что пишем в правильную строку
            if (_Rows.TryGetValue(RowNum, out var row_is_closed))
            {
                //Если строка закрыта
                if (row_is_closed)
                    throw new CellException("Row was closed, but you try write to cell", RowNum, CellNum, GetColumnName(CellNum));

                //Если запись в ячейку выше (левее) текущей
                var last_cell = _Cells.Keys.Where(k => k.row == RowNum).Select(s => s.cell).LastOrDefault(c => c > CellNum);
                if (last_cell != default)
                    throw new CellException($"Record in cell number {CellNum}, that above last recorded cell with number {last_cell}- not available", RowNum, CellNum, GetColumnName(CellNum));
            }
            else
                throw new CellException("Row not added to document, before writing to cell", RowNum, CellNum, GetColumnName(CellNum));

            WriteElement(
                new Cell
                {
                    CellReference = StringValue.FromString($"{GetColumnName(CellNum)}{RowNum}"),
                    CellValue = new CellValue(text),
                    DataType = Type,
                    StyleIndex = StyleIndex
                });
            _Cells.Add(key, true);
        }
        /// <summary> Печатает ячейки с одинаковым значением и стилем со столбца по столбец в одной и той же строке</summary>
        /// <param name="FirstColumn">колонка с которой начали печать</param>
        /// <param name="LastPrintColumn">последняя напечатанная колонка</param>
        /// <param name="RowNumber">строка в которой идёт печать</param>
        /// <param name="StyleIndex">стиль ячейки</param>
        public void PrintEmptyCells(int FirstColumn, int LastPrintColumn, uint RowNumber, uint StyleIndex = 0) =>
            PrintCells(FirstColumn, LastPrintColumn, RowNumber, string.Empty, CellValues.String, StyleIndex);

        /// <summary> Печатает ячейки с одинаковым значением и стилем со столбца по столбец в одной и той же строке</summary>
        /// <param name="FirstColumn">колонка с которой начали печать</param>
        /// <param name="LastPrintColumn">последняя напечатанная колонка</param>
        /// <param name="RowNumber">строка в которой идёт печать</param>
        /// <param name="Value">значение для печати</param>
        /// <param name="Type">Тип входных данных</param>
        /// <param name="StyleIndex">стиль ячейки</param>
        /// <param name="CanReWrite">разрешить перезапись данных (иначе при повторной записи в ячейку будет генерирование ошибки)</param>
        public void PrintCells(int FirstColumn, int LastPrintColumn, uint RowNumber, string Value, CellValues Type = CellValues.String, uint StyleIndex = 0, bool CanReWrite = false)
        {
            foreach (var i in Enumerable.Range(FirstColumn, LastPrintColumn - FirstColumn + 1))
            {
                AddCell(Value, (uint)i, RowNumber, StyleIndex, Type, CanReWrite);
            }
        }

        #region Rows

        /// <summary>
        /// Список записанных строк, со статусом (false - open, true - close)
        /// </summary>
        private readonly Dictionary<uint, bool> _Rows = new();
        /// <summary>
        /// Создаёт новую строку в документе
        /// Если предыдущая строка не закрыта - генерирует ошибку
        /// </summary>
        /// <param name="RowIndex">номер новой строки</param>
        /// <param name="CollapsedLvl">уровень группировки - 0 если без группировки</param>
        /// <param name="ClosePreviousIfOpen">задача закрыть предыдущую строку перед созданием новой</param>
        public void AddRow(uint RowIndex, uint CollapsedLvl = 0, bool ClosePreviousIfOpen = false)
        {
            switch (ClosePreviousIfOpen)
            {
                case true when _Rows.Count > 0:
                    {
                        var last_row = _Rows.Last().Key;
                        CloseRow(last_row);
                        break;
                    }
                case false when _Rows.Count > 0 && !_Rows.Last().Value:
                    throw new RowNotClosedException("You must close the previous line before writing a new one", RowIndex);
            }

            WriteStartElement(new Row { RowIndex = RowIndex }, GetCollapsedAttributes(CollapsedLvl));
            _Rows.Add(RowIndex, false);
        }
        /// <summary> Закрыть строку </summary>
        /// <param name="RowIndex">Номер строки</param>
        public void CloseRow(uint RowIndex)
        {
            if (_Rows.TryGetValue(RowIndex, out var row_is_closed))
            {
                if (row_is_closed)
                    throw new RowNotOpenException("Row was closed, but you agane try close", RowIndex);
            }
            else
                throw new RowException("Row not added to document, but you try close it", RowIndex);

            //if (_Rows.Last().Key is { } row_key && row_key != RowIndex)
            //    throw new RowException(
            //        $"The last row added does not match the one that should be closed, Last row number is {row_key}, but closed {RowIndex}", RowIndex);

            WriteEndElement(); //end of Row
            _Rows[RowIndex] = true;
        }

        #endregion


        /// <summary> Устанавливает фильтр на колонки (ставить в конце листа перед закрытием)</summary>
        /// Позиционировать обязательно в конце страницы после закрытия блока SheetData
        /// перед закрытием блока WorkSheet и MergedList
        /// <param name="ListName">Имя листа</param>
        /// <param name="FirstColumn">первая колонка</param>
        /// <param name="LastColumn">последняя колонка</param>
        /// <param name="FirstRow">первая строка</param>
        /// <param name="LastRow">последняя строка</param>
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
        /// <param name="LastRow">последняя строка</param>
        private void ApprovalFilter(string ListName, uint FirstColumn, uint LastColumn, uint FirstRow, uint LastRow)
        {
            //Секция с фильтром часть-2 - подтверждение принадлежности к листу
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
        /// <summary> Словарь имен колонок excel </summary>
        private static readonly Dictionary<int, string> _Columns = new(676);

        /// <summary> Возвращает строковое имя колонки по номеру (1 - А, 2 - В) </summary>
        /// <param name="index">номер колонки</param>
        /// <returns></returns>
        public static string GetColumnName(uint index) => GetColumnName((int)index);

        /// <summary> Возвращает строковое имя колонки по номеру (1 - А, 2 - В) </summary>
        /// <param name="index">номер колонки</param>
        /// <returns></returns>
        public static string GetColumnName(int index)
        {
            lock (_Columns)
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
        }

        #region MergedCell

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то также что и начало)</param>
        /// <returns></returns>
        public static MergeCell MergeCells(int StartCell, int StartRow, int EndCell, int? EndRow = null)
            => new() { Reference = new StringValue($"{GetColumnName(StartCell)}{StartRow}:{GetColumnName(EndCell)}{EndRow ?? StartRow}") };

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то также что и начало)</param>
        /// <returns></returns>
        public static MergeCell MergeCells(int StartCell, uint StartRow, int EndCell, uint? EndRow = null)
            => new() { Reference = new StringValue($"{GetColumnName(StartCell)}{StartRow}:{GetColumnName(EndCell)}{EndRow ?? StartRow}") };

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то также что и начало)</param>
        /// <returns></returns>
        public MergeCell MergeCells(uint StartCell, uint StartRow, uint EndCell, uint? EndRow = null)
            => new() { Reference = new StringValue($"{GetColumnName(StartCell)}{StartRow}:{GetColumnName(EndCell)}{EndRow ?? StartRow}") };

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то к что и начало)</param>
        /// <returns></returns>
        public MergeCell MergeCells(uint StartCell, int StartRow, uint EndCell, int? EndRow = null)
            => new() { Reference = new StringValue($"{GetColumnName(StartCell)}{StartRow}:{GetColumnName(EndCell)}{EndRow ?? StartRow}") };

        #endregion

        /// <summary> Создаёт запись о группировке для writer </summary>
        /// <param name="lvl">уровень группы</param>
        /// <returns></returns>
        public static OpenXmlAttribute[] GetCollapsedAttributes(uint lvl = 0) => lvl == 0
            ? Array.Empty<OpenXmlAttribute>()
            : new[] { new OpenXmlAttribute("outlineLevel", string.Empty, $"{lvl}"), new OpenXmlAttribute("hidden", string.Empty, $"{lvl}") };


        #endregion

        #region Style Comparer



        /// <summary>
        /// Получить номер стиля похожего на искомый
        /// </summary>
        /// <param name="style">искомый стиль</param>
        /// <returns></returns>
        public uint FirstOrDefault(OpenXmlExStyle style) => FindStyleOrDefault(style).Key;

        /// <summary>
        /// Получить стиль и его номер, похожего на искомый
        /// </summary>
        /// <param name="style">искомый стиль</param>
        /// <returns></returns>

        public KeyValuePair<uint, OpenXmlExStyleCell> FindStyleOrDefault(OpenXmlExStyle style)
        {
            if (style is null) return default;

            return Style.CellsStyles.FirstOrDefault(
                s =>

                #region Заливка

                    (style.FillColor is null || s.Value.FillStyle.Value.FillColor.Key.Equals(style.FillColor)) &&
                    (style.FillPattern is null || s.Value.FillStyle.Value.FillPattern == style.FillPattern) &&

                #endregion

                #region Borders

                    (style.BorderColor is null || s.Value.BorderStyle.Value.BorderColor.Key.Equals(style.BorderColor)) &&
                    (style.LeftBorderStyle is null || s.Value.BorderStyle.Value.LeftBorder.BorderStyle == style.LeftBorderStyle) &&
                    (style.TopBorderStyle is null || s.Value.BorderStyle.Value.TopBorder.BorderStyle == style.TopBorderStyle) &&
                    (style.RightBorderStyle is null || s.Value.BorderStyle.Value.RightBorder.BorderStyle == style.RightBorderStyle) &&
                    (style.BottomBorderStyle is null || s.Value.BorderStyle.Value.BottomBorder.BorderStyle == style.BottomBorderStyle) &&

                #endregion

                #region Шрифт

                    (style.FontSize is null || s.Value.FontStyle.Value.FontSize == style.FontSize) &&
                    (style.FontColor is null || s.Value.FontStyle.Value.FontColor.Key.Equals(style.FontColor)) &&
                    (string.IsNullOrWhiteSpace(style.FontName) || s.Value.FontStyle.Value.FontName == style.FontName) &&
                    (style.IsBoldFont is null || s.Value.FontStyle.Value.IsBoldFont == style.IsBoldFont) &&
                    (style.IsItalicFont is null || s.Value.FontStyle.Value.IsItalicFont == style.IsItalicFont) &&

                #endregion

                #region Выравнивание

                    (style.WrapText is null || s.Value.WrapText == style.WrapText) &&
                    (style.HorizontalAlignment is null || s.Value.HorizontalAlignment == style.HorizontalAlignment) &&
                    (style.VerticalAlignment is null || s.Value.VerticalAlignment == style.VerticalAlignment));

            #endregion



        }

        #endregion
    }
}
