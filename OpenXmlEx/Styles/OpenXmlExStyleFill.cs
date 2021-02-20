using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlEx.Styles
{
    /// <summary> Описывает стиль заливки ячейки </summary>
    public class OpenXmlExStyleFill
    {
        public static Dictionary<uint, Fill> Fills = new()
        {
            { 0, new Fill(new PatternFill() { PatternType = PatternValues.None }) }// Стиль под номером 0 - Заполнение ячейки по умолчанию.
        };

        /// <summary> Цвет заливки </summary>
        public string FillColorHex { get; set; }

        public Fill Fill { get; set; }

        public static void GetStyles(string color)
        {
            var count = (uint)Fills.Count;
            foreach (var style in Generate(color))
            {
                Fills.Add(count, style.Fill);
                count++;
            }
        }


        /// <summary> Генерирует варианты стилей заливки </summary>
        /// <param name="color">цвет</param>
        /// <returns></returns>
        public static IEnumerable<OpenXmlExStyleFill> Generate(string color)
        {
            var patterns = Enum.GetValues<PatternValues>();
            foreach (var pattern in patterns.Where(p => p != PatternValues.None))
            {
                yield return new OpenXmlExStyleFill
                {
                    FillColorHex = color,
                    Fill = new(
                        new PatternFill(
                                new ForegroundColor() { Rgb = new HexBinaryValue() { Value = color } }
                            )
                        { PatternType = pattern })
                };

            }
        }
    }
}