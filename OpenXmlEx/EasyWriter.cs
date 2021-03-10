using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlEx.Errors.Actions;
using OpenXmlEx.Errors.Sheets;
using OpenXmlEx.Styles;
using OpenXmlEx.Styles.Base;
using OpenXmlEx.SubClasses;

namespace OpenXmlEx
{
    public class EasyWriter
    {
        /// <summary>
        /// путь к файлу который записывается
        /// </summary>
        private readonly string _FilePath;
        /// <summary>
        /// стили для документа
        /// </summary>
        private readonly OpenXmlExStyles _Styles;
        /// <summary>
        /// созданный файл-документ
        /// </summary>
        private SpreadsheetDocument _Document { get; set; }
        /// <summary>
        /// инструмент записи данных
        /// </summary>
        private OpenXmlWriterEx _Writer { get; set; }
        /// <summary>
        /// текущая часть записываемого документа
        /// </summary>
        private WorkbookPart _WorkbookPart { get; set; }

        private Workbook _WorkBook { get; set; }
        private Sheets _Sheets { get; set; }

        /// <summary>
        /// список листов в документе (id, sheet), значение - открыта запись или закрыта
        /// </summary>
        private Dictionary<(uint id, Sheet sheet), bool> _SheetDic { get; } = new();

        public EasyWriter(string FilePath)
        {
            _FilePath = FilePath;
            _Styles = new OpenXmlExStyles();
            InitializeDocumentBaseBody(FilePath);
        }
        public EasyWriter(string FilePath, OpenXmlExStyles styles)
        {
            _FilePath = FilePath;
            _Styles = styles;
            InitializeDocumentBaseBody(FilePath);
        }
        public EasyWriter(string FilePath, IEnumerable<BaseOpenXmlExStyle> styles)
        {
            _Styles = OpenXmlWriterEx.GetStyles(styles);
            _FilePath = FilePath;
            InitializeDocumentBaseBody(FilePath);
        }
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
        public void AddSheet(string SheetName = null)
        {
            var sheet_name = string.IsNullOrWhiteSpace(SheetName) ? $"Sheet_{_SheetDic.Count + 1}" : SheetName;
            if (_SheetDic.Keys.Any(k => k.sheet.Name == sheet_name))
                throw new SheetException($"Document allready have sheet with name {sheet_name}, impossible have 2 same name", sheet_name, nameof(AddSheet));

            var ws_part = _WorkbookPart.AddNewPart<WorksheetPart>();
            var sheet = new Sheet { Id = _WorkbookPart.GetIdOfPart(ws_part), SheetId = (uint)_SheetDic.Count + 1, Name = SheetName };
            _SheetDic.Add((sheet.SheetId, sheet), false);

            // ReSharper disable once PossiblyMistakenUseOfParamsMethod
            _Sheets.Append(sheet);
            CreateWriter(ws_part);
        }
        /// <summary>
        /// статус открыта ли новая секция документа
        /// </summary>
        private bool WorksheetIsOpen { get; set; }
        /// <summary>
        /// статус открыта ли лист для записи
        /// </summary>
        private bool SheetIsOpen { get; set; }

        /// <summary> Сбрасывает статусы флагов на листе </summary>
        void CloseWorkSheetFlags()
        {
            WorksheetIsOpen = false;
            SheetIsOpen = false;
        }
        /// <summary> Создаёт новое перо для записи в документ </summary>
        /// <param name="wsPart">Часть документа для записи</param>
        public void CreateWriter(WorksheetPart wsPart)
        {
            if (_Writer != null)
            {
                if (SheetIsOpen)
                    _Writer.WriteEndElement();
                if (WorksheetIsOpen)
                    _Writer.WriteEndElement();
                CloseWorkSheetFlags();

                _Writer?.Close();
            }

            _Writer = new OpenXmlWriterEx(wsPart, _Styles);

            NewWorkSheet();
        }
        /// <summary> Создаёт в структуре документа новый WorkSheet и помечает его как открытый </summary>
        private void NewWorkSheet()
        {
            _Writer.WriteStartElement(new Worksheet());
            WorksheetIsOpen = true;
        }
        /// <summary> Устанавливает тип группировки для строк и столбцов </summary>
        /// <param name="SummaryBelow">группировать сверху (false - сверху, true - снизу)</param>
        /// <param name="SummaryRight">группировать справа (false - справа, true - слева)</param>
        public void SetGrouping(bool SummaryBelow = false, bool SummaryRight = false)
        {
            if (_Writer is null)
                throw new GroupingException($"You have not Active Writer in document - {Path.GetFileName(_FilePath)}", null, nameof(SetGrouping));
            var (sheet, _) = _SheetDic.LastOrDefault();
            if (sheet == default)
                throw new GroupingException("You have not sheets to create grouping", null, nameof(SetGrouping));

            if (WorksheetIsOpen && !SheetIsOpen)
            {
                _Writer.SetGrouping(SummaryBelow, SummaryRight);
                return;
            }

            throw new GroupingException("Wrong location to set grouping, set before opening entry in sheet", sheet.sheet.Name, nameof(SetGrouping));
        }

        /// <summary> Устанавливает параметры столбцов </summary>
        /// <param name="settings">список надстроек для листа</param>
        public void SetWidth(IEnumerable<WidthOpenXmlEx> settings)
        {
            if (_Writer is null)
                throw new SetWidthException($"You have not Active Writer in document - {Path.GetFileName(_FilePath)}", null, nameof(SetWidth));
            var (sheet, _) = _SheetDic.LastOrDefault();
            if (sheet == default)
                throw new SetWidthException("You have not sheets to set Width settings for the cells", null, nameof(SetWidth));

            if (WorksheetIsOpen && !SheetIsOpen)
            {
                _Writer.SetWidth(settings);
                return;
            }

            throw new SetWidthException("Wrong location to set Width settings for the cells, set before opening entry in sheet", sheet.sheet.Name, nameof(SetWidth));

        }

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
            if (!SheetIsOpen)
            {
                _Writer.WriteStartElement(new SheetData());
                SheetIsOpen = true;
            }
            _Writer.AddRow(RowIndex,CollapsedLvl,ClosePreviousIfOpen,AddSkipedRows);
        }
    }
}
