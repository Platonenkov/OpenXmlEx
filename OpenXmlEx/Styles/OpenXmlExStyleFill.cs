using System;
using System.Collections.Generic;
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
            { PatternType = FillPattern });

        #endregion

        #region Свойства для поиска стиля

        /// <summary> Цвет заливки </summary>
        public KeyValuePair<Color, string> FillColor { get; set; }
        /// <summary> Стиль заливки </summary>
        public PatternValues FillPattern { get; set; }

        #endregion

        #region Поля

        /// <summary> Стили заливки </summary>
        private static IEnumerable<PatternValues> __Patterns;
        /// <summary> Стили заливки </summary>
        private static IEnumerable<PatternValues> Patterns => __Patterns ??= Enum.GetValues<PatternValues>();

        #endregion

        #region Генераторы

        /// <summary> Генерирует Default стиль заполнения </summary>
        /// <returns></returns>
        public static OpenXmlExStyleFill GetDefault() => new() // Стиль под номером 0 - Заполнение ячейки по умолчанию.
        {
            FillPattern = PatternValues.None,
            FillColor = (new KeyValuePair<Color, string>(Color.Transparent, Color.Transparent.ToHexConverter()))
        };

        #endregion

        #region Конструкторы

        public OpenXmlExStyleFill()
        {

        }
        public OpenXmlExStyleFill(Color StyleFillColor, PatternValues StyleFillPattern)
        {
            FillColor = new KeyValuePair<Color, string>(StyleFillColor, StyleFillColor.ToHexConverter());
            FillPattern = StyleFillPattern;
        }


        #endregion

    }
}