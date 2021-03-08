using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlEx.Styles.Base;
using Color = System.Drawing.Color;

namespace OpenXmlEx.Styles
{
    public class OpenXmlExStyles
    {
        /// <summary> перечень всех комбинаций заливок стиля </summary>
        public Dictionary<uint, OpenXmlExStyleFill> Fills { get; } = new() { { 0, OpenXmlExStyleFill.GetDefault() } };

        /// <summary> перечень всех комбинаций рамок стиля </summary>
        public Dictionary<uint, OpenXmlExStyleBorderGrand> Borders { get; } = new() { { 0, OpenXmlExStyleBorderGrand.GetDefault() } };
        /// <summary> перечень всех генерированных стилей шрифта </summary>
        public Dictionary<uint, OpenXmlExStyleFont> Fonts { get; } = new() { { 0, OpenXmlExStyleFont.GetDefault() } };

        /// <summary> перечень всех генерированных стилей ячеек </summary>
        public Dictionary<uint, OpenXmlExStyleCell> CellsStyles { get; } = new()
        {
            {
                0,
                new OpenXmlExStyleCell()
                {
                    BorderStyle = new KeyValuePair<uint, OpenXmlExStyleBorderGrand>(0, OpenXmlExStyleBorderGrand.GetDefault()),
                    FillStyle = new KeyValuePair<uint, OpenXmlExStyleFill>(0, OpenXmlExStyleFill.GetDefault()),
                    FontStyle = new KeyValuePair<uint, OpenXmlExStyleFont>(0, OpenXmlExStyleFont.GetDefault()),
                    WrapText = false,
                    HorizontalAlignment = HorizontalAlignmentValues.Left,
                    VerticalAlignment = VerticalAlignmentValues.Center
                }
            }
        };

        public Stylesheet Styles { get; set; }

        #region Конструкторы

        public OpenXmlExStyles()
        {
        }

        public OpenXmlExStyles(IEnumerable<OpenXmlExStyle> styles)
        {
            GenerateStyles(styles);
        }

        #endregion

        /// <summary> Генератор стилей </summary>
        /// <param name="styles">стили заданные пользователем</param>
        private void GenerateStyles(IEnumerable<OpenXmlExStyle> styles)
        {
            foreach (var style in styles)
            {
                #region генератор стилей заливки

                var fill = new OpenXmlExStyleFill(style.FillColor ?? default, style.FillPattern ?? PatternValues.None);
                var fill_count = (uint)Fills.Count;
                Fills.Add(fill_count, fill);

                #endregion

                #region генератор стилей рамки

                var border = new OpenXmlExStyleBorderGrand(style.LeftBorderStyle ?? BorderStyleValues.None,
                    style.TopBorderStyle ?? BorderStyleValues.None,
                    style.RightBorderStyle ?? BorderStyleValues.None,
                    style.BottomBorderStyle ?? BorderStyleValues.None,
                    style.BorderColor ?? Color.Transparent);

                var borders_count = (uint)Borders.Count;
                Borders.Add(borders_count, border);

                #endregion

                #region генератор стилей шрифтов

                var font_count = (uint)Fonts.Count;
                var font = new OpenXmlExStyleFont(style.FontName, style.FontSize ?? 11, style.FontColor ?? Color.Black, style.IsBoldFont ?? false, style.IsItalicFont ?? false);
                Fonts.Add(font_count, font);

                #endregion

                #region генератор стилей рамки

                var cells_format = new OpenXmlExStyleCell(new KeyValuePair<uint, OpenXmlExStyleFont>(font_count, font),
                    new KeyValuePair<uint, OpenXmlExStyleFill>(fill_count, fill),
                    new KeyValuePair<uint, OpenXmlExStyleBorderGrand>(borders_count, border),
                    style.WrapText ?? false,
                    style.HorizontalAlignment ?? HorizontalAlignmentValues.Left,
                    style.VerticalAlignment ?? VerticalAlignmentValues.Center);

                var cell_count = (uint)CellsStyles.Count;
                CellsStyles.Add(cell_count, cells_format);

                #endregion



            }

            Styles = GetStylesheet();
        }



        private Stylesheet GetStylesheet() =>
            new(
                new Fonts(Fonts.Values.Select(f => f.Font)),
                new Fills(Fills.Values.Select(f => f.Fill)),
                new Borders(Borders.Values.Select(b => b.Border)),
                new CellFormats(CellsStyles.Values.Select(c => c.CellStyle)));
    }
}
