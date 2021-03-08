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
        #region стили  начиная с дефолтных (Обязательные 1 и 2 стили, второй всегда будет заливка Sepia)




        /// <summary> перечень всех комбинаций заливок стиля </summary>
        public Dictionary<uint, OpenXmlExStyleFill> Fills { get; } = new() { { 0, OpenXmlExStyleFill.GetDefault() },{ 1, OpenXmlExStyleFill.GetDefault() } };

        /// <summary> перечень всех комбинаций рамок стиля </summary>
        public Dictionary<uint, OpenXmlExStyleBorderGrand> Borders { get; } = new() { { 0, OpenXmlExStyleBorderGrand.GetDefault() },{ 1, OpenXmlExStyleBorderGrand.GetDefault() } };
        /// <summary> перечень всех генерированных стилей шрифта </summary>
        public Dictionary<uint, OpenXmlExStyleFont> Fonts { get; } = new() { { 0, OpenXmlExStyleFont.GetDefault() },{ 1, OpenXmlExStyleFont.GetDefault() } };

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
            },
            {
                1,
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

        #endregion

        private Stylesheet _Styles { get; }
        public Stylesheet Styles => _Styles;

        #region Конструкторы

        public OpenXmlExStyles()
        {
            _Styles = GetStylesheet();
        }

        public OpenXmlExStyles(IEnumerable<OpenXmlExStyle> styles)
        {
            GenerateStyles(styles);
            _Styles = GetStylesheet();

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
        }



        private Stylesheet GetStylesheet()
        {
            //var styles = new Stylesheet();
            //var fonts = new Fonts();
            //foreach (var font in Fonts.Values.Select(f => f.Font))
            //    fonts.Append(font);

            //var fills = new Fills();
            //foreach (var fill in Fills.Values.Select(f => f.Fill))
            //    fills.Append(fill);

            //var borders = new Borders();
            //foreach (var border in Borders.Values.Select(b => b.Border))
            //    borders.Append(border);


            //var cells = new CellFormats();
            //foreach (var cell in CellsStyles.Values.Select(c => c.CellStyle))
            //    cells.Append(cell);

            //styles.Append(fonts);
            //styles.Append(fills);
            //styles.Append(borders);
            //styles.Append(cells);
            //return styles;
            return new(
                new Fonts(Fonts.Values.Select(f => f.Font)),
                new Fills(Fills.Values.Select(f => f.Fill)),
                new Borders(Borders.Values.Select(b => b.Border)),
                new CellFormats(CellsStyles.Values.Select(c => c.CellStyle)));
        }
    }
}
