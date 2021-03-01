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

        public OpenXmlExStyles(IEnumerable<OpenXmlExStyle> styles)
        {
            GenerateStyles(styles);
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
            var sizes = FontSizes.ToArray();

#if DEBUG
            Console.WriteLine($"Создание базовых таблиц для {color_dic.Count} цветов");
#endif
            foreach (var color in color_dic)
            {
#if DEBUG
                Console.WriteLine($"работа с цветом {color.Key}");
#endif

                #region генератор стилей заливки

                var fills = OpenXmlExStyleFill.GetStyles(color);
                var fill_count = (uint)Fills.Count;
                foreach (var fill in fills)
                    Fills.Add(fill_count++, fill);
#if DEBUG
                Console.WriteLine($"Имеется {Fills.Count} FILLS");
#endif

                #endregion

                #region генератор стилей рамки

                var borders = OpenXmlExStyleBorderGrand.GetStyles(color);
                var borders_count = (uint)Borders.Count;
                foreach (var border in borders)
                    Borders.Add(borders_count++, border);
#if DEBUG
                Console.WriteLine($"Имеется {Borders.Count} BORDERS");
#endif

                #endregion

                #region генератор стилей шрифтов
                var font_count = (uint)Fonts.Count;
                foreach (var font_name in fonts)
                {
                    var generated_fonts = OpenXmlExStyleFont.GetStyles(color, font_name, sizes);
                    foreach (var font in generated_fonts)
                        Fonts.Add(font_count++, font);
                }
#if DEBUG
                Console.WriteLine($"Имеется {Fonts.Count} FONTS");
#endif

#if DEBUG
                Console.WriteLine($"Завершена работа с цветом {color.Key}");
#endif

                #endregion
            }
#if DEBUG
            Console.WriteLine($"Запуск генератора стилей ячеек из комбинаций");
#endif

            #region генератор стилей рамки

            var cells_formats = OpenXmlExStyleCell.GetStyles(Fills, Borders, Fonts);
            var cell_count = (uint)CellsStyles.Count;
            foreach (var cell in cells_formats)
                CellsStyles.Add(cell_count++, cell);

            #endregion
#if DEBUG
            Console.WriteLine($"Завершение работы генератора стилей ячеек\nДобавлено {CellsStyles.Count} стилей");
#endif

#if DEBUG
            Console.WriteLine($"Генерация стилей документа XML");
#endif

            Styles = GetStylesheet();
#if DEBUG
            Console.WriteLine($"Завершение работы генератора стилей");
#endif
        }

        private void Init(
            string FontName, double? FontSize, Color? FontColor, bool? IsBoldFont, bool? IsItalicFont, //Шрифт
            HorizontalAlignmentValues? HorizontalAlignment, VerticalAlignmentValues? VerticalAlignment, bool? WrapText, //выравнивание содержимого
            BorderStyleValues? LeftBorderStyle, BorderStyleValues? TopBorderStyle, BorderStyleValues? RightBorderStyle, BorderStyleValues? BottomBorderStyle, Color? BorderColor, // рамка
            Color? FillColor, PatternValues? FillPattern) //заливка
        {

        }
        /// <summary> Генератор стилей </summary>
        /// <param name="styles">стили заданные пользователем</param>
        private void GenerateStyles(IEnumerable<OpenXmlExStyle> styles)
        {
#if DEBUG
            Console.WriteLine($"Создание базовых таблиц стилей");
#endif

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
            }
#if DEBUG
            Console.WriteLine($"Запуск генератора стилей ячеек из комбинаций");
#endif

            #region генератор стилей рамки

            var cells_formats = OpenXmlExStyleCell.GetStyles(Fills, Borders, Fonts);
            var cell_count = (uint)CellsStyles.Count;
            foreach (var cell in cells_formats)
                CellsStyles.Add(cell_count++, cell);

            #endregion
#if DEBUG
            Console.WriteLine($"Завершение работы генератора стилей ячеек\nДобавлено {CellsStyles.Count} стилей");
#endif

#if DEBUG
            Console.WriteLine($"Генерация стилей документа XML");
#endif

            Styles = GetStylesheet();
#if DEBUG
            Console.WriteLine($"Завершение работы генератора стилей");
#endif
        }



        private Stylesheet GetStylesheet() =>
            new(
                new Fonts(Fonts.Values.Select(f => f.Font)),
                new Fills(Fills.Values.Select(f => f.Fill)),
                new Borders(Borders.Values.Select(b => b.Border)),
                new CellFormats(CellsStyles.Values.Select(c => c.CellStyle)));
    }
}
