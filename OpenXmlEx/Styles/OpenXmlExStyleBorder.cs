using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlEx.Styles
{
    /// <summary> Описывает стиль рамки </summary>
    public class OpenXmlExStyleBorder
    {
        /// <summary> Стиль линии рамки </summary>
        public BorderStyleValues BorderStyle { get; set; }
        /// <summary> цвет рамки </summary>
        public Color BorderColor { get; set; }


        /// <summary> Генерирует варианты стиля рамки </summary>
        /// <param name="color">цвет</param>
        /// <returns></returns>
        public static IEnumerable<OpenXmlExStyleBorder> GetStyles(string color)
        {
            var border_styles = Enum.GetValues<BorderStyleValues>();
            foreach (var border_style in border_styles)
            {
                yield return new OpenXmlExStyleBorder() { BorderColor = new Color() { Rgb = color }, BorderStyle = border_style };
            }
        }

        public T GetStyle<T>() where T: BorderPropertiesType, new()
        {
            return new T {Style = BorderStyle, Color = BorderColor };
        }

    }
}