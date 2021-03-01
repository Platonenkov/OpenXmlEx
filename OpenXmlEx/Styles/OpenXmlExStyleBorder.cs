using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlEx.Styles
{
    /// <summary> Описывает стиль рамки </summary>
    public class OpenXmlExStyleBorder
    {
        /// <summary> Генерирует default  стиль для рамки </summary>
        public static OpenXmlExStyleBorder Default => new()
        {
            BorderColor = new KeyValuePair<System.Drawing.Color, string>(System.Drawing.Color.Transparent, System.Drawing.Color.Transparent.ToHexConverter()),
            BorderStyle = BorderStyleValues.None
        };
        /// <summary> Стиль линии рамки </summary>
        public BorderStyleValues BorderStyle { get; set; }

        public Color BorderColorXML => new() {Rgb = BorderColor.Value};
        /// <summary> цвет рамки </summary>
        public KeyValuePair<System.Drawing.Color, string> BorderColor { get; set; }

        /// <summary> Стили рамок </summary>
        private static IEnumerable<BorderStyleValues> __BorderStyles;

        #region Конструкторы

        public OpenXmlExStyleBorder()
        {
            
        }
        public OpenXmlExStyleBorder(System.Drawing.Color StyleBorderColor, BorderStyleValues Style)
        {
            BorderColor = new KeyValuePair<System.Drawing.Color, string>(StyleBorderColor, StyleBorderColor.ToHexConverter());
            BorderStyle = Style;
        }


        #endregion
        /// <summary> Стили рамок </summary>
        private static IEnumerable<BorderStyleValues> BorderStyles => __BorderStyles ??= Enum.GetValues<BorderStyleValues>();

        /// <summary> Генерирует варианты стиля рамки </summary>
        /// <param name="color">цвет</param>
        /// <returns></returns>
        public static IEnumerable<OpenXmlExStyleBorder> GetStyles(KeyValuePair<System.Drawing.Color, string> color) =>
            BorderStyles.Select(border_style => new OpenXmlExStyleBorder()
            {
                BorderColor = color,
                BorderStyle = border_style
            });
    }
}