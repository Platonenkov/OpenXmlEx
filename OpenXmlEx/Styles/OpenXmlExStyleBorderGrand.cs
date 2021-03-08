﻿using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;

namespace OpenXmlEx.Styles
{
    /// <summary> Описывает стиль всех рамок у ячейки </summary>
    public class OpenXmlExStyleBorderGrand
    {
        #region Конструкторы

        public OpenXmlExStyleBorderGrand()
        {
            
        }
        public OpenXmlExStyleBorderGrand(
            BorderStyleValues LeftBorderStyle,
            BorderStyleValues TopBorderStyle,
            BorderStyleValues RightBorderStyle,
            BorderStyleValues BottomBorderStyle,
            Color StyleBorderColor)
        {
            LeftBorder = new OpenXmlExStyleBorder(StyleBorderColor, LeftBorderStyle);
            TopBorder = new OpenXmlExStyleBorder(StyleBorderColor, TopBorderStyle);
            RightBorder = new OpenXmlExStyleBorder(StyleBorderColor, RightBorderStyle);
            BottomBorder = new OpenXmlExStyleBorder(StyleBorderColor, BottomBorderStyle);

            BorderColor = new KeyValuePair<Color, string>(StyleBorderColor, StyleBorderColor.ToHexConverter());
        }

        #endregion

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

        #endregion
    }
}