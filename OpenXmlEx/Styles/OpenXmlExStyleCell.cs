using System;
using System.Collections.Generic;
using System.Linq;
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

        #endregion

        #region Статические поля

        private static IEnumerable<HorizontalAlignmentValues> __HAlign;
        private static IEnumerable<HorizontalAlignmentValues> H_Align => __HAlign ??= Enum.GetValues<HorizontalAlignmentValues>();

        private static IEnumerable<VerticalAlignmentValues> __VAlign;

        private static IEnumerable<VerticalAlignmentValues> V_Align => __VAlign ??= Enum.GetValues<VerticalAlignmentValues>();
        
        #endregion

        #region Генераторы

        /// <summary>
        /// генерирует варианты комбинаций стиля ячейки на основе входных стилей
        /// </summary>
        /// <param name="Fills">стили заливки</param>
        /// <param name="Borders">стили рамок</param>
        /// <param name="Fonts">стили шрифтов</param>
        /// <returns></returns>
        public static IEnumerable<OpenXmlExStyleCell> GetStyles(Dictionary<uint, OpenXmlExStyleFill> Fills,
            Dictionary<uint, OpenXmlExStyleBorderGrand> Borders, Dictionary<uint, OpenXmlExStyleFont> Fonts) =>
            from fonts_value in Fonts
            from fills_value in Fills
            from borders_value in Borders
            from style in Generate(fonts_value, fills_value, borders_value)
            select style;

        /// <summary>
        /// генерирует варианты комбинаций стиля ячейки на основе входных номеров стилей
        /// </summary>
        /// <param name="Fill">стиль заливки</param>
        /// <param name="Border">стиль рамки</param>
        /// <param name="Font">стиль шрифта</param>
        /// <returns></returns>
        private static IEnumerable<OpenXmlExStyleCell> Generate(
            KeyValuePair<uint, OpenXmlExStyleFont> Font,
            KeyValuePair<uint, OpenXmlExStyleFill> Fill,
            KeyValuePair<uint, OpenXmlExStyleBorderGrand> Border)
        {

            foreach (var h_align in H_Align)
                foreach (var v_align in V_Align)
                    for (var wrap = 0; wrap < 2; wrap++)
                        yield return new OpenXmlExStyleCell()
                        {
                            FontStyle = Font,
                            FillStyle=Fill,
                            BorderStyle = Border,
                            HorizontalAlignment = h_align,
                            VerticalAlignment = v_align,
                            WrapText = wrap == 1
                        };
        }
        /// <summary> Генерирует стиль на основании данных класса </summary>
        /// <returns></returns>
        private CellFormat GetCellStyle() => new(
                new Alignment() { Horizontal = HorizontalAlignment, Vertical = VerticalAlignment, WrapText = WrapText })
            { FontId = FontStyle.Key, FillId = FillStyle.Key, BorderId = BorderStyle.Key, ApplyFont = true };

        #endregion
    }
}