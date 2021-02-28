using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;

namespace OpenXmlEx.Styles
{
    /// <summary> Описывает стиль всех рамок у ячейки </summary>
    public class OpenXmlExStyleBorderGrand
    {
        /// <summary> Рамка OpenXML </summary>
        public Border Border { get; set; }

        #region Свойства рамки для поиска стиля

        /// <summary> левая рамка </summary>
        public OpenXmlExStyleBorder LeftBorder { get; set; }
        /// <summary> верхняя рамка </summary>
        public OpenXmlExStyleBorder TopBorder { get; set; }
        /// <summary> правая рамка </summary>
        public OpenXmlExStyleBorder RightBorder { get; set; }
        /// <summary> нижняя рамка </summary>
        public OpenXmlExStyleBorder BottomBorder { get; set; }
        /// <summary> цвет рамки </summary>
        public KeyValuePair<Color, string> BorderColor { get; set; }

        #endregion

        #region Генераторы

        /// <summary> Генерирует default стиль рамки </summary>
        /// <returns></returns>
        public static OpenXmlExStyleBorderGrand GetDefault() => new()
        {
            Border = new Border( // Стиль под номером 0 - Грани.
                new LeftBorder(),
                new RightBorder(),
                new TopBorder(),
                new BottomBorder(),
                new DiagonalBorder()),
            BorderColor = new KeyValuePair<Color, string>(Color.Transparent, Color.Transparent.ToHexConverter()),
            LeftBorder = OpenXmlExStyleBorder.Default,
            TopBorder = OpenXmlExStyleBorder.Default,
            RightBorder = OpenXmlExStyleBorder.Default,
            BottomBorder = OpenXmlExStyleBorder.Default,

        };

        /// <summary> Создаёт стили рамок на основе комбинаций, в заданном цвете </summary>
        /// <param name="color">цвет рамки</param>
        public static IEnumerable<OpenXmlExStyleBorderGrand> GetStyles(KeyValuePair<Color, string> color)
        {
            var styles = OpenXmlExStyleBorder.GetStyles(color).ToArray();

            foreach (var border in styles)
                yield return new OpenXmlExStyleBorderGrand()
                {
                    LeftBorder = border,
                    TopBorder = border,
                    RightBorder = border,
                    BottomBorder = border,
                    Border = new Border(
                        new LeftBorder(border.BorderColorXML) { Style = border.BorderStyle },
                        new TopBorder(border.BorderColorXML) { Style = border.BorderStyle },
                        new RightBorder(border.BorderColorXML) { Style = border.BorderStyle },
                        new BottomBorder(border.BorderColorXML) { Style = border.BorderStyle },
                        new DiagonalBorder()
                    ),
                    BorderColor = color

                };
        }

        #endregion
    }
}