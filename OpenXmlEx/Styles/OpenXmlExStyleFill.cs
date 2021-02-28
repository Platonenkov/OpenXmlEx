using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;

namespace OpenXmlEx.Styles
{
    /// <summary> Описывает стиль заливки ячейки </summary>
    public class OpenXmlExStyleFill
    {
        #region OpenXML

        /// <summary> Заливка OpenXML </summary>
        public Fill Fill => GetFill();
        /// <summary> Возвращает стиль формата OpenXML </summary>
        /// <returns></returns>
        private Fill GetFill() => new(
            new PatternFill(
                    new ForegroundColor() { Rgb = new HexBinaryValue() { Value = FillColor.Value } }
                )
            { PatternType = Pattern });

        #endregion

        #region Свойства для поиска стиля

        /// <summary> Цвет заливки </summary>
        public KeyValuePair<Color, string> FillColor { get; set; }
        /// <summary> Стиль заливки </summary>
        public PatternValues Pattern { get; set; }

        #endregion

        #region Поля

        /// <summary> Стили заливки </summary>
        private static IEnumerable<PatternValues> __Patterns;
        /// <summary> Стили заливки </summary>
        private static IEnumerable<PatternValues> Patterns => __Patterns ??= Enum.GetValues<PatternValues>();

        #endregion

        #region Генераторы

        /// <summary> Генерирует варианты стиля на основе цвета </summary>
        /// <param name="color">цвет</param>
        public static IEnumerable<OpenXmlExStyleFill> GetStyles(KeyValuePair<Color, string> color) =>
            Patterns.Where(p => p != PatternValues.None).Select(pattern => new OpenXmlExStyleFill
            {
                FillColor = color,
                Pattern = pattern
            });

        /// <summary> Генерирует Default стиль заполнения </summary>
        /// <returns></returns>
        public static OpenXmlExStyleFill GetDefault() => new() // Стиль под номером 0 - Заполнение ячейки по умолчанию.
        {
            Pattern = PatternValues.None,
            FillColor = (new KeyValuePair<Color, string>(Color.Transparent, Color.Transparent.ToHexConverter()))
        };

        #endregion
    }
}