using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlEx.Styles
{
    /// <summary> определяет стили </summary>
    public class OpenXmlExStyleCell
    {
        /// <summary> OpenXML стиль ячейки </summary>
        public CellFormat CellStyle => GetCellStyle();

        #region Свойства класса для поиска

        /// <summary> стиль шрифтов </summary>
        public KeyValuePair<uint, OpenXmlExStyleFont> FontStyle { get; set; }
        /// <summary> стиль заливки </summary>
        public KeyValuePair<uint, OpenXmlExStyleFill> FillStyle { get; set; }
        /// <summary> стиль рамки </summary>
        public KeyValuePair<uint, OpenXmlExStyleBorderGrand> BorderStyle { get; set; }

        /// <summary> будет ли перенос текста в ячейке </summary>
        public bool WrapText { get; set; }

        /// <summary> Горизонтальное выравнивание в ячейке </summary>
        public HorizontalAlignmentValues HorizontalAlignment { get; set; }
        /// <summary> Вертикальное выравнивание в ячейке </summary>
        public VerticalAlignmentValues VerticalAlignment { get; set; }

        public uint TextRotation { get; set; }
        #endregion

        #region Генераторы

        public OpenXmlExStyleCell()
        {
            
        }

        public OpenXmlExStyleCell(KeyValuePair<uint, OpenXmlExStyleFont> Font, 
            KeyValuePair<uint, OpenXmlExStyleFill> Fill,
            KeyValuePair<uint, OpenXmlExStyleBorderGrand> Border,
            bool Wrap, HorizontalAlignmentValues h_align, VerticalAlignmentValues v_align, uint textRotation)
        {
            FontStyle = Font;
            FillStyle = Fill;
            BorderStyle = Border;
            HorizontalAlignment = h_align;
            VerticalAlignment = v_align;
            WrapText = Wrap;
            TextRotation = textRotation;
        }

        /// <summary> Генерирует стиль на основании данных класса </summary>
        /// <returns></returns>
        private CellFormat GetCellStyle() => new(
                new Alignment() { Horizontal = HorizontalAlignment, Vertical = VerticalAlignment, WrapText = WrapText,TextRotation = TextRotation})
            { FontId = FontStyle.Key, FillId = FillStyle.Key, BorderId = BorderStyle.Key, ApplyFont = true };

        #endregion
    }
}