using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlEx.Errors;
using OpenXmlEx.Errors.Actions;
using OpenXmlEx.Errors.Sheets;
using OpenXmlEx.Styles;
using OpenXmlEx.Styles.Base;
using OpenXmlEx.SubClasses;

namespace OpenXmlEx
{
    public class EasyWriter : IBaseWriter
    {
        /// <summary>
        /// путь к файлу который записывается
        /// </summary>
        private readonly string _FilePath;

        private readonly Encoding _Encoding;
        #region Документ

        /// <summary> стили для документа </summary>
        private readonly OpenXmlExStyles _Styles;
        /// <summary> созданный файл-документ </summary>
        private SpreadsheetDocument _Document;
        /// <summary> инструмент записи данных </summary>
        private OpenXmlWriterEx _Writer;
        /// <summary> текущая часть записываемого документа </summary>
        private WorkbookPart _WorkbookPart;
        private Workbook _WorkBook;
        private Sheets _Sheets;

        #endregion

        /// <summary>
        /// список листов в документе (id, sheet), значение - открыта запись - false или закрыта - true
        /// </summary>
        private Dictionary<(uint id, Sheet sheet), bool> _SheetDic { get; } = new();

        #region Конструкторы
        public EasyWriter(string FilePath, Encoding encoding, OpenXmlExStyles styles)
        {
            _FilePath = FilePath;
            _Styles = new OpenXmlExStyles(styles.BaseStyles);
            _Encoding = encoding;
            InitializeDocumentBaseBody(FilePath);
        }

        public EasyWriter(string FilePath, OpenXmlExStyles styles) : this(FilePath, Encoding.UTF8, styles)
        {

        }
        public EasyWriter(string FilePath) : this(FilePath, Encoding.UTF8, new OpenXmlExStyles())
        {
        }

        public EasyWriter(string FilePath, Encoding encoding, IEnumerable<BaseOpenXmlExStyle> styles) : this(FilePath, encoding, new OpenXmlExStyles(styles))
        {

        }
        public EasyWriter(string FilePath, IEnumerable<BaseOpenXmlExStyle> styles) : this(FilePath, Encoding.UTF8, new OpenXmlExStyles(styles))
        {

        }

        #endregion

        /// <summary> Инициализация базовой структуры документа </summary>
        /// <param name="FilePath">путь к документу</param>
        private void InitializeDocumentBaseBody(string FilePath)
        {

            _Document = SpreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook);
            _WorkbookPart = _Document.AddWorkbookPart();


            #region styles to the document

            var wbsp = _WorkbookPart.AddNewPart<WorkbookStylesPart>();
            wbsp.Stylesheet = _Styles.Styles;
            wbsp.Stylesheet.Save();

            #endregion

            _WorkBook = _WorkbookPart.Workbook = new Workbook();
            _Sheets = _WorkBook.AppendChild(new Sheets());
        }
        /// <summary> Добавить новый лист в документ </summary>
        /// <param name="SheetName">имя листа</param>
        public void AddNewSheet(string SheetName = null)
        {
            var sheet_name = string.IsNullOrWhiteSpace(SheetName) ? $"Sheet_{_SheetDic.Count + 1}" : SheetName;
            if (_SheetDic.Keys.Any(k => k.sheet.Name == sheet_name))
                throw new SheetException($"Document allready have sheet with name {sheet_name}, impossible have 2 same name", sheet_name, nameof(AddNewSheet));

            var ws_part = _WorkbookPart.AddNewPart<WorksheetPart>();
            var sheet = new Sheet { Id = _WorkbookPart.GetIdOfPart(ws_part), SheetId = (uint)_SheetDic.Count + 1, Name = SheetName };
            _SheetDic.Add((sheet.SheetId, sheet), false);

            // ReSharper disable once PossiblyMistakenUseOfParamsMethod
            _Sheets.Append(sheet);
            CreateWriter(ws_part, sheet_name);
        }

        /// <summary> Создаёт новое перо для записи в документ </summary>
        /// <param name="wsPart">Часть документа для записи</param>
        /// <param name="SheetName">имя листа</param>
        private void CreateWriter(WorksheetPart wsPart, string SheetName)
        {
            #region Закрываем текущее перо
            //_Writer?.CloseSheet();

            _Writer?.Close();
            _Writer = null;

            #endregion

            _Writer = new OpenXmlWriterEx(wsPart, _Encoding, _Styles, SheetName);

            _Writer.WriteStartElement(new Worksheet());
        }

        #region Base Settings

        /// <summary> Устанавливает тип группировки для строк и столбцов </summary>
        /// <param name="SummaryBelow">группировать сверху (false - сверху, true - снизу)</param>
        /// <param name="SummaryRight">группировать справа (false - справа, true - слева)</param>
        public void SetGrouping(bool SummaryBelow = false, bool SummaryRight = false)
        {
            CheckForError(nameof(SetGrouping));
            _Writer.SetGrouping(SummaryBelow, SummaryRight);

        }

        /// <summary> Устанавливает параметры столбцов </summary>
        /// <param name="settings">список надстроек для листа</param>
        public void SetWidth(IEnumerable<WidthOpenXmlEx> settings)
        {
            CheckForError(nameof(SetWidth));
            _Writer.SetWidth(settings);
        }

        /// <summary> отложенная установка фильтра на колонки (ставить в конце листа перед закрытием)</summary>
        /// установит фильтр перед закрытием документа
        /// <param name="ListName">Имя листа</param>
        /// <param name="FirstColumn">первая колонка</param>
        /// <param name="LastColumn">последняя колонка</param>
        /// <param name="FirstRow">первая строка</param>
        /// <param name="LastRow">последняя строка</param>
        public void SetFilter(uint FirstColumn, uint LastColumn, uint FirstRow, uint? LastRow = null, string ListName = null) =>
            _Writer.SetFilter(FirstColumn, LastColumn, FirstRow, LastRow, ListName);
        #endregion

        #region Rows


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
            CheckForError(nameof(AddRow));
            _Writer.AddRow(RowIndex, CollapsedLvl, ClosePreviousIfOpen, AddSkipedRows);
        }

        /// <summary> Закрыть строку </summary>
        /// <param name="RowNumber">Номер строки</param>
        public void CloseRow(uint RowNumber)
        {
            CheckForError(nameof(AddRow));
            _Writer.CloseRow(RowNumber);
        }

        #endregion

        #region Cells

        /// <summary> Добавляет значение в ячейку документа </summary>
        /// <param name="text">текст для записи</param>
        /// <param name="CellNum">номер колонки</param>
        /// <param name="RowNum">номер строки</param>
        /// <param name="StyleIndex">индекс стиля</param>
        /// <param name="Type">тип данных</param>
        /// <param name="CanReWrite">разрешить перезапись данных (иначе при повторной записи в ячейку будет генерирование ошибки)</param>
        public void AddCell(string text, uint CellNum, uint RowNum, uint StyleIndex = 0, CellValues Type = CellValues.String, bool CanReWrite = false)
        {
            CheckForError(nameof(AddCell));
            _Writer.AddCell(text, CellNum, RowNum, StyleIndex, Type, CanReWrite);
        }
        /// <summary> Печатает ячейки с одинаковым значением и стилем со столбца по столбец в одной и той же строке</summary>
        /// <param name="FirstColumn">колонка с которой начали печать</param>
        /// <param name="LastPrintColumn">последняя напечатанная колонка</param>
        /// <param name="RowNumber">строка в которой идёт печать</param>
        /// <param name="StyleIndex">стиль ячейки</param>
        public void PrintEmptyCells(int FirstColumn, int LastPrintColumn, uint RowNumber, uint StyleIndex = 0) =>
            AddCellsSameData(FirstColumn, LastPrintColumn, RowNumber, string.Empty, CellValues.String, StyleIndex);

        /// <summary> Печатает ячейки с одинаковым значением и стилем со столбца по столбец в одной и той же строке</summary>
        /// <param name="FirstColumn">колонка с которой начали печать</param>
        /// <param name="LastPrintColumn">последняя напечатанная колонка</param>
        /// <param name="RowNumber">строка в которой идёт печать</param>
        /// <param name="Value">значение для печати</param>
        /// <param name="Type">Тип входных данных</param>
        /// <param name="StyleIndex">стиль ячейки</param>
        /// <param name="CanReWrite">разрешить перезапись данных (иначе при повторной записи в ячейку будет генерирование ошибки)</param>
        public void AddCellsSameData(
            int FirstColumn, int LastPrintColumn, uint RowNumber, string Value, CellValues Type = CellValues.String, uint StyleIndex = 0,
            bool CanReWrite = false)
        {
            foreach (var i in Enumerable.Range(FirstColumn, LastPrintColumn - FirstColumn + 1))
            {
                AddCell(Value, (uint)i, RowNumber, StyleIndex, Type, CanReWrite);
            }
        }

        #endregion

        #region MergeCells
        /// <summary> Формирует объединенную ячейку для документа </summary>
        /// <param name="new_range">новый диапазон для объединения</param>
        public void MergeCells(OpenXmlMergedCellEx new_range)
        {
            if (_Writer is null)
                throw new MergeCellException("Start new sheet before set merge", new_range, null, nameof(MergeCells));
            _Writer.MergeCells(new_range);
        }

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то также что и начало)</param>
        /// <returns></returns>
        public void MergeCells(int StartCell, int StartRow, int EndCell, int? EndRow = null) =>
            MergeCells(new OpenXmlMergedCellEx((uint)StartCell, (uint)StartRow, (uint)EndCell, EndRow is null ? (uint)StartRow : (uint)EndRow));

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то также что и начало)</param>
        /// <returns></returns>
        public void MergeCells(uint StartCell, uint StartRow, uint EndCell, uint? EndRow = null) =>
            MergeCells(new OpenXmlMergedCellEx(StartCell, StartRow, EndCell, EndRow ?? StartRow));

        #endregion

        #region Styles

        /// <summary>
        /// Получить стиль и его номер, похожего на искомый
        /// </summary>
        /// <param name="style">искомый стиль</param>
        /// <returns></returns>

        public KeyValuePair<uint, OpenXmlExStyleCell> FindStyleOrDefault(BaseOpenXmlExStyle style) => _Styles.FindStyleOrDefault(style);

        #endregion

        /// <summary>
        /// Проверка базовых ошибок перед записью
        /// </summary>
        /// <param name="methodName"></param>
        private void CheckForError(string methodName)
        {
            ThrowIfObjectDisposed();
            if (_Writer is null)
                throw new WriterException($"You have not Active Writer in document - {Path.GetFileName(_FilePath)}", null, methodName);
            var data = _SheetDic.LastOrDefault();
            if (data.Value)
                throw new WriterException("You sheet was closed, but you try write to it", data.Key.sheet.Name, methodName);
            if (data.Key == default)
                throw new WriterException("You have not sheets to set settings", null, methodName);

        }

        /// <summary>
        /// Вызывается для закрытия записи и освобождения документа
        /// </summary>
        public void Close()
        {
            _Writer?.Close();
            _Document?.Close();
        }

        #region Dispose

        private bool _disposed;

        /// <summary>
        /// Throw if object is disposed.
        /// </summary>
        protected virtual void ThrowIfObjectDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(GetType().Name);
            }
        }

        /// <summary>
        /// Closes the reader, and releases all resources.
        /// </summary>
        /// <param name="disposing">true to release both managed and unmanaged resources; false to release only unmanaged resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    Close();
                }

                _disposed = true;
            }
        }

        /// <summary>
        /// Closes the writer, and releases all resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion

    }
}
