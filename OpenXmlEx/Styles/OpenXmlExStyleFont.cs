using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;

namespace OpenXmlEx.Styles
{
    /// <summary> Описывает стиль текста </summary>
    public class OpenXmlExStyleFont
    {
        /// <summary> Font OpenXML </summary>
        public Font Font => GetStyle();

        #region Свойства стиля для поиска

        /// <summary> Размер шрифта </summary>
        public double FontSize { get; set; }
        /// <summary> цвет шрифта </summary>
        public KeyValuePair<System.Drawing.Color, string> FontColor { get; set; }
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

        /// <summary> Генерирует варианты стиля шрифта </summary>
        /// <param name="color">цвет</param>
        /// <param name="fontName">иям шрифта</param>
        /// <param name="FontSizes">Размерности шрифтов</param>
        /// <returns></returns>
        public static IEnumerable<OpenXmlExStyleFont> GetStyles(KeyValuePair<Color, string> color, string fontName, IEnumerable<uint> FontSizes)
        {
            foreach (var font_size in FontSizes)
                for (var bold = 0; bold < 2; bold++)
                    for (var italic = 0; italic < 2; italic++)
                        yield return new OpenXmlExStyleFont
                        {
                            FontColor = color,
                            FontName = fontName,
                            FontSize = font_size,
                            IsBoldFont = bold == 1,
                            IsItalicFont = italic == 1
                        };
        }

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
