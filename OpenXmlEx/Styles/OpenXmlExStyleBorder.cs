using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlEx.Styles
{
    /// <summary> Описывает стиль рамки </summary>
    public class OpenXmlExStyleBorder
    {
        /// <summary> Генерирует default  стиль для рамки </summary>
        public static OpenXmlExStyleBorder Default => new()
        {
            BorderColor = new KeyValuePair<System.Drawing.Color, Color>(
                System.Drawing.Color.Transparent,
                new Color() {Rgb = System.Drawing.Color.Transparent.ToHexConverter()}),
            BorderStyle = BorderStyleValues.None
        };
        /// <summary> Стиль линии рамки </summary>
        public BorderStyleValues BorderStyle { get; set; }

        /// <summary> цвет рамки </summary>
        public KeyValuePair<System.Drawing.Color, Color> BorderColor { get; set; }

        /// <summary> Стили рамок </summary>
        private static IEnumerable<BorderStyleValues> __BorderStyles;
        /// <summary> Стили рамок </summary>
        private static IEnumerable<BorderStyleValues> BorderStyles => __BorderStyles ??= Enum.GetValues<BorderStyleValues>();

        /// <summary> Генерирует варианты стиля рамки </summary>
        /// <param name="color">цвет</param>
        /// <returns></returns>
        public static IEnumerable<OpenXmlExStyleBorder> GetStyles(KeyValuePair<System.Drawing.Color, string> color)
        {
            var (c, rgb) = color;

            foreach (var border_style in BorderStyles)
            {
                yield return new OpenXmlExStyleBorder()
                {
                    BorderColor = new KeyValuePair<System.Drawing.Color, Color>(c, new Color() { Rgb = rgb }),
                    BorderStyle = border_style
                };
            }
        }

        public T GetStyle<T>() where T : BorderPropertiesType, new()
        {
            return new T { Style = BorderStyle, Color = BorderColor.Value };
        }

    }
}