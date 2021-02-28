using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;

namespace OpenXmlEx.Styles
{
    public class OpenXmlExStyles
    {
        /// <summary> перечень всех комбинаций заливок стиля </summary>
        public Dictionary<uint, OpenXmlExStyleFill> Fills = new() { { 0, OpenXmlExStyleFill.GetDefault() } };

        /// <summary> перечень всех комбинаций рамок стиля </summary>
        public Dictionary<uint, OpenXmlExStyleBorderGrand> Borders = new() { { 0, OpenXmlExStyleBorderGrand.GetDefault() } };
        /// <summary> перечень всех генерированных стилей шрифта </summary>
        public Dictionary<uint, OpenXmlExStyleFont> Fonts { get; } = new() { { 0, OpenXmlExStyleFont.GetDefault() } };

        /// <summary> перечень всех генерированных стилей ячеек </summary>
        public Dictionary<uint, OpenXmlExStyleCell> CellsFormats = new()
        {
            {
                0,
                new OpenXmlExStyleCell()
                {
                    BorderStyleNum = 0,
                    FillStyleNum = 0,
                    FontStyleNum = 0,
                    WrapText = false,
                    HorizontalAlignment = HorizontalAlignmentValues.Left,
                    VerticalAlignment = VerticalAlignmentValues.Center
                }
            }
        };

        public Stylesheet Styles { get; set; }

        #region Конструкторы

        public OpenXmlExStyles(IEnumerable<string> FontNames, IEnumerable<uint> FontSizes, IEnumerable<Color> Colors)
        {
            GenerateStyles(FontNames, FontSizes, Colors);
        }
        public OpenXmlExStyles(IEnumerable<string> FontNames, IEnumerable<uint> FontSizes)
        {
            GenerateStyles(FontNames, FontSizes, new[] { Color.Black });
        }
        public OpenXmlExStyles(IEnumerable<string> FontNames)
        {
            GenerateStyles(FontNames, new[] { 11U }, new[] { Color.Black });
        }

        public OpenXmlExStyles()
        {
            GenerateStyles(new[] { "Times New Roman" }, new[] { 11U }, new[] { Color.Black });
        }

        #endregion

        /// <summary> Генератор стилей </summary>
        /// <param name="FontNames">список шрифтов</param>
        /// <param name="FontSizes">размерность шрифтов</param>
        /// <param name="Colors">список цветов</param>
        private void GenerateStyles(IEnumerable<string> FontNames, IEnumerable<uint> FontSizes, IEnumerable<Color> Colors)
        {
            var colors = Colors.ToArray();

            var color_dic = colors.ToDictionary(c => c, c => c.ToHexConverter());
            var fonts = FontNames.ToArray();

            foreach (var color in color_dic)
            {
                #region генератор стилей заливки

                var fills = OpenXmlExStyleFill.GetStyles(color);
                var fill_count = (uint)Fills.Count;
                foreach (var fill in fills)
                    Fills.Add(fill_count++, fill);

                #endregion

                #region генератор стилей рамки

                var borders = OpenXmlExStyleBorderGrand.GetStyles(color);
                var borders_count = (uint)Borders.Count;
                foreach (var border in borders)
                    Borders.Add(borders_count++, border);

                #endregion

                #region генератор стилей шрифтов
                var font_count = (uint)Fills.Count;
                foreach (var font_name in fonts)
                {
                    var generated_fonts = OpenXmlExStyleFont.GetStyles(color, font_name, FontSizes);
                    foreach (var font in generated_fonts)
                        Fonts.Add(font_count++, font);
                }


                #endregion
            }

            #region генератор стилей рамки

            var cells_formats = OpenXmlExStyleCell.GetStyles(Fills, Borders, Fonts);
            var cell_count = (uint)CellsFormats.Count;
            foreach (var cell in cells_formats)
                CellsFormats.Add(cell_count++, cell);

            #endregion

            Styles = GetStylesheet();
        }


        private Stylesheet GetStylesheet() =>
            new(
                new Fonts(Fonts.Values.Select(f => f.Font)),
                new Fills(Fills.Values.Select(f => f.Fill)),
                new Borders(Borders.Values.Select(b => b.Border)),
                new CellFormats(CellsFormats.Values.Select(c => c.CellStyle)));
    }
}
