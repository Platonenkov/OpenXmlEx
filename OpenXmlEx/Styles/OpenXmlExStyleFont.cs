using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlEx.Styles
{
    /// <summary> Описывает стиль текста </summary>
    public class OpenXmlExStyleFont
    {
        public static Dictionary<uint, Font> Fonts { get; } = new();

        /// <summary> Размер шрифта </summary>
        public double FontSize { get; set; }
        /// <summary> цвет шрифта </summary>
        public string FontColorHex { get; set; }
        /// <summary> Имя шрифта </summary>
        public string FontName { get; set; }
        /// <summary> жирный или нет </summary>
        public bool IsBoldFont { get; set; }
        /// <summary> курсивный или нет </summary>
        public bool IsItalicFont { get; set; }

        public static void GetStyles(string color, string fontName)
        {
            var count = (uint)Fonts.Count;
            foreach (var style in Generate(color, fontName))
            {
                Fonts.Add(count, style.GetStyle());
                count++;
            }
        }

        /// <summary> Генерирует варианты стиля шрифта </summary>
        /// <param name="color">цвет</param>
        /// <param name="fontName">иям шрифта</param>
        /// <returns></returns>
        private static IEnumerable<OpenXmlExStyleFont> Generate(string color, string fontName)
        {
            foreach (var font_size in Enumerable.Range(1, 409))
                for (var bold = 0; bold < 2; bold++)
                    for (var italic = 0; italic < 2; italic++)
                        yield return new OpenXmlExStyleFont { FontColorHex = color, FontName = fontName, FontSize = font_size, IsBoldFont = bold == 1, IsItalicFont = italic == 1 };
        }

        private Font GetStyle()
        {
            return !IsBoldFont && !IsItalicFont
                ? new Font(
                    new FontSize() { Val = FontSize },
                    new Color() { Rgb = new HexBinaryValue() { Value = FontColorHex } },
                    new FontName() { Val = FontName })
                : IsBoldFont && IsItalicFont
                    ? new Font(
                    new Bold(),
                    new Italic(),
                    new FontSize() { Val = FontSize },
                    new Color() { Rgb = new HexBinaryValue() { Value = FontColorHex } },
                    new FontName() { Val = FontName })
                : IsBoldFont
                    ? new Font(
                    new Bold(),
                    new FontSize() { Val = FontSize },
                    new Color() { Rgb = new HexBinaryValue() { Value = FontColorHex } },
                    new FontName() { Val = FontName })
                    : new Font(
                    new Italic(),
                    new FontSize() { Val = FontSize },
                    new Color() { Rgb = new HexBinaryValue() { Value = FontColorHex } },
                    new FontName() { Val = FontName });
        }
    }
}
