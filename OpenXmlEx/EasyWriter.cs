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
using OpenXmlEx.Styles;
using OpenXmlEx.Styles.Base;
using OpenXmlEx.SubClasses;

namespace OpenXmlEx
{
    public class EasyWriter
    {
        private readonly string _FilePath;
        private OpenXmlExStyles _Styles;
        private SpreadsheetDocument _Document { get; set; }
        private OpenXmlWriterEx _Writer { get; set; }
        private WorkbookPart _WorkbookPart { get; set; }
        private WorkbookStylesPart _WorkbookStylesPart { get; set; }
        private Workbook _WorkBook { get; set; }
        private Sheets _Sheets { get; set; }

        /// <summary>
        /// список листов в документе (id, sheet), значение - открыта запись или закрыта
        /// </summary>
        private Dictionary<(uint id, Sheet sheet), bool> _SheetDic { get; } = new ();
        
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
            _WorkbookStylesPart = _WorkbookPart.AddNewPart<WorkbookStylesPart>();
            _WorkBook = _WorkbookPart.Workbook = new Workbook();
            _Sheets = _WorkBook.AppendChild(new Sheets());
        }

        public void AddSheet(string SheetName)
        {
            var ws_part = _WorkbookPart.AddNewPart<WorksheetPart>();
            var sheet = new Sheet { Id = _WorkbookPart.GetIdOfPart(ws_part), SheetId = (uint)_SheetDic.Count+1, Name = SheetName };
            _SheetDic.Add((sheet.SheetId,sheet),false);

            // ReSharper disable once PossiblyMistakenUseOfParamsMethod
            _Sheets.Append(sheet);
            CreateWriter(ws_part);
        }

        private bool WorksheetIsOpen { get; set; }
        private bool SheetIsOpen { get; set; }

        void CloseWorkSheet()
        {
            WorksheetIsOpen = false;
            SheetIsOpen = false;
        }
        public void CreateWriter(WorksheetPart wsPart)
        {
            if (_Writer != null)
            {
                if(SheetIsOpen)
                    _Writer.WriteEndElement();
                if(WorksheetIsOpen)
                    _Writer.WriteEndElement();
                CloseWorkSheet();

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
                throw new GroupingException($"You have not Active Writer in document - {Path.GetFileName(_FilePath)}", null);
            var (sheet, _) = _SheetDic.LastOrDefault();
            if (sheet == default)
                throw new GroupingException("You have not sheets to create grouping", null);

            if (WorksheetIsOpen && !SheetIsOpen)
            {
                _Writer.SetGrouping(SummaryBelow, SummaryRight);
                return;
            }
            
            throw new GroupingException("Wrong location to set grouping, set before opening entry in sheet", sheet.sheet.Name);
        }

        /// <summary> Устанавливает параметры столбцов </summary>
        /// <param name="settings">список надстроек для листа</param>
        public void SetWidth(IEnumerable<WidthOpenXmlEx> settings)
        {
            if(_Writer is null)
                throw new SetWidthException($"You have not Active Writer in document - {Path.GetFileName(_FilePath)}", null);
            var (sheet, _) = _SheetDic.LastOrDefault();
            if (sheet == default)
                throw new SetWidthException("You have not sheets to set Width settings for the cells", null);

            if (WorksheetIsOpen && !SheetIsOpen)
            {
                _Writer.SetWidth(settings);
                return;
            }

            throw new SetWidthException("Wrong location to set Width settings for the cells, set before opening entry in sheet", sheet.sheet.Name);

        }

    }
}
