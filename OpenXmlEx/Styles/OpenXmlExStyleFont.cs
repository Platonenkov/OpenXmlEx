using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;

namespace OpenXmlEx.Styles
{
    /// <summary> Описывает стиль текста </summary>
    public class OpenXmlExStyleFont
    {
        #region Конструктор

        public OpenXmlExStyleFont(string StyleFontName, double StyleFontSize, Color StyleFontColor, bool StyleIsBoldFont, bool StyleIsItalicFont)
        {
            FontName = string.IsNullOrWhiteSpace(StyleFontName) ? "Times New Roman" : StyleFontName;
            FontSize = StyleFontSize;
            FontColor = new KeyValuePair<Color, string>(StyleFontColor, StyleFontColor.ToHexConverter());
            IsBoldFont = StyleIsBoldFont;
            IsItalicFont = StyleIsItalicFont;
        }

        public OpenXmlExStyleFont()
        {

        }

        #endregion
        /// <summary> Font OpenXML </summary>
        public Font Font => GetStyle();

        #region Свойства стиля для поиска

        /// <summary> Размер шрифта </summary>
        public double FontSize { get; set; }
        /// <summary> цвет шрифта </summary>
        public KeyValuePair<Color, string> FontColor { get; set; }
        /// <summary> Имя шрифта </summary>
        public string FontName { get; set; }
        /// <summary> жирный или нет </summary>
        public bool IsBoldFont { get; set; }
        /// <summary> курсивный или нет </summary>
        public bool IsItalicFont { get; set; }

        #endregion

        /// <summary> Генерирует default стиль </summary>
        /// <returns></returns>
        public static OpenXmlExStyleFont GetDefault() => new() // Стиль под номером 0 (default)
        {
            FontSize = 11,
            FontColor = new KeyValuePair<Color, string>
            (
                Color.Black,
                Color.Black.ToHexConverter()
            ),
            FontName = "Times New Roman",
            IsBoldFont = false,
            IsItalicFont = false
        };

        #region Генераторы

        /// <summary> Генерирует стиль OpenXML для шрифта </summary>
        /// <returns></returns>
        private Font GetStyle() =>
            !IsBoldFont && !IsItalicFont
                ? new Font(
                    new FontSize() { Val = FontSize },
                    new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = new HexBinaryValue() { Value = FontColor.Value } },
                    new FontName() { Val = FontName })
                : IsBoldFont && IsItalicFont
                    ? new Font(
                        new Bold(),
                        new Italic(),
                        new FontSize() { Val = FontSize },
                        new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = new HexBinaryValue() { Value = FontColor.Value } },
                        new FontName() { Val = FontName })
                    : IsBoldFont
                        ? new Font(
                            new Bold(),
                            new FontSize() { Val = FontSize },
                            new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = new HexBinaryValue() { Value = FontColor.Value } },
                            new FontName() { Val = FontName })
                        : new Font(
                            new Italic(),
                            new FontSize() { Val = FontSize },
                            new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = new HexBinaryValue() { Value = FontColor.Value } },
                            new FontName() { Val = FontName });

        #endregion

    }
}
