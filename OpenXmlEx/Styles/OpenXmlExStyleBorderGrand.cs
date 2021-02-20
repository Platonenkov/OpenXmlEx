using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlEx.Styles
{
    /// <summary> Описывает стиль всех рамок у ячейки </summary>
    public class OpenXmlExStyleBorderGrand
    {
        public static Dictionary<uint, Border> Borders = new()
        {
            {
                0,
                new Border( // Стиль под номером 0 - Грани.
                    new LeftBorder(),
                    new RightBorder(),
                    new TopBorder(),
                    new BottomBorder(),
                    new DiagonalBorder())
            }
        };

        /// <summary> левая рамка </summary>
        public OpenXmlExStyleBorder LeftBorder { get; set; }
        /// <summary> верхняя рамка </summary>
        public OpenXmlExStyleBorder TopBorder { get; set; }
        /// <summary> правая рамка </summary>
        public OpenXmlExStyleBorder RightBorder { get; set; }
        /// <summary> нижняя рамка </summary>
        public OpenXmlExStyleBorder BottomBorder { get; set; }
        /// <summary> диагональ рамка </summary>
        public DiagonalBorder Diagonal { get; set; } = new();


        public static void GetStyles(string color)
        {
            var count = (uint)Borders.Count;
            foreach (var style in Generate(color))
            {
                Borders.Add(count, style.GetStyle());
                count++;
            }
        }

        /// <summary> Генерирует возможные комбинации стиля рамок </summary>
        /// <param name="color">цвет рамки</param>
        /// <returns></returns>
        public static IEnumerable<OpenXmlExStyleBorderGrand> Generate(string color)
        {
            var styles = OpenXmlExStyleBorder.GetStyles(color).ToArray();

            foreach (var left in styles)
                foreach (var top in styles)
                    foreach (var right in styles)
                        foreach (var bottom in styles)
                            yield return new OpenXmlExStyleBorderGrand() { LeftBorder = left, TopBorder = top, RightBorder = right, BottomBorder = bottom };
        }

        private Border GetStyle()
        {
            return new()
            {
                LeftBorder = LeftBorder.GetStyle<LeftBorder>(),
                TopBorder = TopBorder.GetStyle<TopBorder>(),
                RightBorder = RightBorder.GetStyle<RightBorder>(),
                BottomBorder = BottomBorder.GetStyle<BottomBorder>()
            };
        }

    }
}