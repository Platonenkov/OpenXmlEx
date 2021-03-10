using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlEx.Errors;
using OpenXmlEx.Errors.Cells;
using OpenXmlEx.Errors.Rows;
using OpenXmlEx.Extensions;
using Column = DocumentFormat.OpenXml.Spreadsheet.Column;
using Columns = DocumentFormat.OpenXml.Spreadsheet.Columns;
using OpenXmlEx.Styles;
using OpenXmlEx.Styles.Base;
using OpenXmlEx.SubClasses;

namespace OpenXmlEx
{
    public class OpenXmlWriterEx : OpenXmlPartWriter
    {
        public static OpenXmlExStyles GetStyles(IEnumerable<BaseOpenXmlExStyle> styles) => new(styles);

        public OpenXmlExStyles Style { get; private set; }

        #region Конструкторы

        #region приоритет 1

        /// <inheritdoc />
        public OpenXmlWriterEx(
            OpenXmlPart OpenXmlPart,
            OpenXmlExStyles styles)
            : base(OpenXmlPart) => InitStyles(styles);

        /// <inheritdoc />
        public OpenXmlWriterEx(OpenXmlPart OpenXmlPart, Encoding encoding,
            OpenXmlExStyles styles)
            : base(OpenXmlPart, encoding) => InitStyles(styles);

        /// <inheritdoc />
        public OpenXmlWriterEx(Stream PartStream,
            OpenXmlExStyles styles)
            : base(PartStream) => InitStyles(styles);

        /// <inheritdoc />
        public OpenXmlWriterEx(Stream PartStream, Encoding encoding,
            OpenXmlExStyles styles)
            : base(PartStream, encoding) => InitStyles(styles);

        private void InitStyles(OpenXmlExStyles styles)
            => Style = styles;

        #endregion

        #region приоритет 2

        /// <inheritdoc />
        public OpenXmlWriterEx(
            OpenXmlPart OpenXmlPart,
            IEnumerable<BaseOpenXmlExStyle> styles)
            : base(OpenXmlPart) => InitStyles(styles);

        /// <inheritdoc />
        public OpenXmlWriterEx(OpenXmlPart OpenXmlPart, Encoding encoding,
            IEnumerable<BaseOpenXmlExStyle> styles)
            : base(OpenXmlPart, encoding) => InitStyles(styles);

        /// <inheritdoc />
        public OpenXmlWriterEx(Stream PartStream,
            IEnumerable<BaseOpenXmlExStyle> styles)
            : base(PartStream) => InitStyles(styles);

        /// <inheritdoc />
        public OpenXmlWriterEx(Stream PartStream, Encoding encoding,
            IEnumerable<BaseOpenXmlExStyle> styles)
            : base(PartStream, encoding) => InitStyles(styles);

        private void InitStyles(IEnumerable<BaseOpenXmlExStyle> styles)
            => Style = new OpenXmlExStyles(styles);

        #endregion

        #endregion

        #region Overrides of OpenXmlPartWriter
        /// <summary>
        /// Добавляет записи о новых элементах в словари
        /// </summary>
        /// <param name="elementObject">новый элемент</param>
        /// <param name="closed">открытый или закрыт</param>
        void AddIndex(OpenXmlElement elementObject, bool closed)
        {
            switch (elementObject)
            {
                case Row row:
                    _Rows.Add(row.RowIndex, closed);
                    break;

                case Cell cell:
                    {
                        var address = OpenXmlExHelper.GetCellAddress(cell);
                        if (address.Equals(default))
                            return;
                        _Cells.Add((address.rowNum, address.collNum), closed);
                        break;
                    }
            }

        }

        public override void WriteStartElement(OpenXmlElement elementObject, IEnumerable<OpenXmlAttribute> attributes)
        {
            base.WriteStartElement(elementObject, attributes);
            AddIndex(elementObject, false);
        }
        public override void WriteStartElement(OpenXmlElement elementObject)
        {
            base.WriteStartElement(elementObject);
            AddIndex(elementObject, false);
        }
        public override void WriteElement(OpenXmlElement elementObject)
        {
            base.WriteElement(elementObject);
            AddIndex(elementObject, true);
        }

        public override void Close()
        {
            var (cell_key, cell_value) = _Cells.LastOrDefault();
            if (cell_key != default && !cell_value)
            {
                _Cells[cell_key] = true;
                WriteEndElement();
            }
            var (row_key, row_value) = _Rows.LastOrDefault();
            if (row_key != default && !row_value)
            {
                CloseRow(row_key);
            }

            base.Close();

        }

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
        public void SetWidth(IEnumerable<WidthOpenXmlEx> Settings)
        {

            #region Установка ширины колонок

            WriteStartElement(new Columns());
            foreach (var set in Settings)
                WriteElement(new Column { Min = set.First, Max = set.Last, Width = set.Width });
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
        /// <param name="CellNum">номер колонки</param>
        /// <param name="RowNum">номер строки</param>
        /// <param name="StyleIndex">индекс стиля</param>
        /// <param name="Type">тип данных</param>
        /// <param name="CanReWrite">разрешить перезапись данных (иначе при повторной записи в ячейку будет генерирование ошибки)</param>
        public void AddCell(string text, uint CellNum, uint RowNum, uint StyleIndex = 0, CellValues Type = CellValues.String, bool CanReWrite = false)
        {
            #region Проверки и ошибки

            //Проверка валидности номера строки или столбца (должны быть больше 0)
            if (CellNum == 0 || RowNum == 0)
            {
                throw new ArgumentException($"Address must be greater that 0, Row({RowNum}) and Cell({CellNum})");
            }
            (uint row, uint cell) key = (RowNum, CellNum);

            //Проверка на перезапись данных
            if (_Cells.TryGetValue(key, out var _) && !CanReWrite)
            {
                throw new CellException("Re-writing data to a cell", RowNum, CellNum, OpenXmlExHelper.GetColumnName(CellNum), nameof(AddCell));
            }
            // проверка на то что пишем в правильную строку
            if (_Rows.TryGetValue(RowNum, out var row_is_closed))
            {
                //Если строка закрыта
                if (row_is_closed)
                    throw new CellException("Row was closed, but you try write to cell", RowNum, CellNum, OpenXmlExHelper.GetColumnName(CellNum), nameof(AddCell));

                //Если запись в ячейку выше (левее) текущей
                var last_cell = _Cells.Keys.Where(k => k.row == RowNum).Select(s => s.cell).LastOrDefault(c => c > CellNum);
                if (last_cell != default)
                    throw new CellException(
                        $"Record in cell number {CellNum}, that above last recorded cell with number {last_cell} - not available", RowNum, CellNum,
                        OpenXmlExHelper.GetColumnName(CellNum), nameof(AddCell));
            }
            else
                throw new CellException("Row not added to document, before writing to cell", RowNum, CellNum, OpenXmlExHelper.GetColumnName(CellNum), nameof(AddCell));

            #endregion

            WriteElement(
                new Cell
                {
                    CellReference = StringValue.FromString($"{OpenXmlExHelper.GetColumnName(CellNum)}{RowNum}"),
                    CellValue = new CellValue(text),
                    DataType = Type,
                    StyleIndex = StyleIndex
                });
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
        public void PrintCells(
            int FirstColumn, int LastPrintColumn, uint RowNumber, string Value, CellValues Type = CellValues.String, uint StyleIndex = 0,
            bool CanReWrite = false)
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
        /// <param name="AddSkipedRows">Добавить пропущенные строки (если пишем 2-ю строку, а первую не записали - будет ошибка)</param>
        public void AddRow(uint RowIndex, uint CollapsedLvl = 0, bool ClosePreviousIfOpen = false, bool AddSkipedRows = false)
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
                    throw new RowNotClosedException("You must close the previous line before writing a new one", RowIndex, nameof(AddRow));
            }

            var previous = _Rows.Keys.LastOrDefault();
            if (previous > 0 && RowIndex - 1 != previous && !AddSkipedRows)
                throw new RowException($"Rows must go in order, Last used row was {previous}", RowIndex, nameof(AddRow));
            if (AddSkipedRows)
            {
                for (var r = previous + 1; r < RowIndex; r++)
                    WriteElement(new Row { RowIndex = r });
            }
            WriteStartElement(new Row { RowIndex = RowIndex }, OpenXmlExHelper.GetCollapsedAttributes(CollapsedLvl));
        }

        /// <summary> Закрыть строку </summary>
        /// <param name="RowIndex">Номер строки</param>
        public void CloseRow(uint RowIndex)
        {
            if (_Rows.TryGetValue(RowIndex, out var row_is_closed))
            {
                if (row_is_closed)
                    throw new RowNotOpenException($"Row was closed, but you agane try close, Row - {RowIndex}", RowIndex, nameof(CloseRow));
            }
            else
                throw new RowException("Row not added to document, but you try close it", RowIndex, nameof(CloseRow));

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
            WriteElement(new AutoFilter { Reference = $"{OpenXmlExHelper.GetColumnName(FirstColumn)}{FirstRow}:{OpenXmlExHelper.GetColumnName(LastColumn)}{LastRow ?? FirstRow}" });
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
                    Text = $"{ListName}!${OpenXmlExHelper.GetColumnName(FirstColumn)}${FirstRow}:${OpenXmlExHelper.GetColumnName(LastColumn)}${LastRow}"
                });
            WriteEndElement(); //Filter
        }

        /// <summary>
        /// Устанавливает объединенные ячейки на листе
        /// Позиционировать обязательно в конце страницы после закрытия блока SheetData
        /// после блока фильтров но до закрытия блока WorkSheet
        /// </summary>
        /// <param name="MergedCells">перечень объединенных ячеек</param>
        public void SetMergedList()
        {
            WriteStartElement(new MergeCells());
            foreach (var mer in _MergedCells) WriteElement(mer.Value);
            WriteEndElement();
        }


        #endregion

        #region MergedCell

        private Dictionary<OpenXmlMergedCellEx, MergeCell> _MergedCells { get; } = new();
        private void MergeCells(OpenXmlMergedCellEx merged)
        {
            if (_MergedCells.Keys.Any(
                k => k.Equals(merged) || merged.StartRow > k.StartRow && merged.EndRow < k.EndRow || merged.EndRow > k.StartRow && merged.EndRow < k.EndRow))
            {

            }
            _MergedCells.Add(merged, new() { Reference = new StringValue($"{OpenXmlExHelper.GetColumnName(merged.StartCell)}{merged.StartRow}:{OpenXmlExHelper.GetColumnName(merged.EndCell)}{merged.EndRow}") });
        }

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то также что и начало)</param>
        /// <returns></returns>
        public void MergeCells(int StartCell, int StartRow, int EndCell, int? EndRow = null)
            => MergeCells(new OpenXmlMergedCellEx((uint)StartCell, (uint)StartRow, (uint)EndCell, EndRow is null ? (uint)StartRow : (uint)EndRow));

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то также что и начало)</param>
        /// <returns></returns>
        public void MergeCells(int StartCell, uint StartRow, int EndCell, uint? EndRow = null)
            => MergeCells(new OpenXmlMergedCellEx((uint)StartCell, StartRow, (uint)EndCell, EndRow ?? StartRow));

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то также что и начало)</param>
        /// <returns></returns>
        public void MergeCells(uint StartCell, uint StartRow, uint EndCell, uint? EndRow = null)
            => MergeCells(new OpenXmlMergedCellEx(StartCell, StartRow, EndCell, EndRow ?? StartRow));

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то к что и начало)</param>
        /// <returns></returns>
        public void MergeCells(uint StartCell, int StartRow, uint EndCell, int? EndRow = null)
            => MergeCells(new OpenXmlMergedCellEx(StartCell, (uint)StartRow, EndCell, EndRow is null ? (uint)StartRow : (uint)EndRow));

        #endregion

        #region Style Comparer

        /// <summary>
        /// Получить номер стиля похожего на искомый
        /// </summary>
        /// <param name="style">искомый стиль</param>
        /// <returns></returns>
        public uint FirstOrDefault(BaseOpenXmlExStyle style) => FindStyleOrDefault(style).Key;

        /// <summary>
        /// Получить стиль и его номер, похожего на искомый
        /// </summary>
        /// <param name="style">искомый стиль</param>
        /// <returns></returns>

        public KeyValuePair<uint, OpenXmlExStyleCell> FindStyleOrDefault(BaseOpenXmlExStyle style)
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
