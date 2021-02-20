using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlEx.Styles
{
    /// <summary> определяет стили </summary>
    public class OpenXmlExStyleCell
    {
        public static Dictionary<uint, CellFormat> CellsFormats = new();

        /// <summary> Номер стиля шрифтов </summary>
        public uint FontStyleNum { get; set; }
        /// <summary> Номер стиля заливки </summary>
        public uint FillStyleNum { get; set; }
        /// <summary> Номер стиля рамки </summary>
        public uint BorderStyleNum { get; set; }
        /// <summary> будет ли перенос текста в ячейке </summary>
        public bool WrapText { get; set; }

        /// <summary> Горизонтальное выравнивание в ячейке </summary>
        public HorizontalAlignmentValues HorizontalAlignment { get; set; }
        /// <summary> Вертикальное выравнивание в ячейке </summary>
        public VerticalAlignmentValues VerticalAlignment { get; set; }

        private static IEnumerable<HorizontalAlignmentValues> H_Align { get; set; }
        private static IEnumerable<VerticalAlignmentValues> V_Align { get; set; }
        static OpenXmlExStyleCell()
        {
            H_Align = Enum.GetValues<HorizontalAlignmentValues>();
            V_Align = Enum.GetValues<VerticalAlignmentValues>();
        }

        public static void GetStyles()
        {
            var count = (uint)CellsFormats.Count;

            foreach (var fonts_value in OpenXmlExStyleFont.Fonts)
                foreach (var fills_value in OpenXmlExStyleFill.Fills)
                    foreach (var borders_value in OpenXmlExStyleBorderGrand.Borders)
                        foreach (var style in OpenXmlExStyleCell.Generate(fonts_value.Key, fills_value.Key, borders_value.Key))
                        {
                            CellsFormats.Add(count, style.GetCellFormat());
                            count++;
                        }

        }
        public static IEnumerable<OpenXmlExStyleCell> Generate(uint FontStyleNum, uint FillStyleNum, uint BorderStyleNum)
        {

            foreach (var h_align in H_Align)
                foreach (var v_align in V_Align)
                    for (var wrap = 0; wrap < 2; wrap++)
                        yield return new OpenXmlExStyleCell()
                        {
                            FontStyleNum = FontStyleNum,
                            FillStyleNum = FillStyleNum,
                            BorderStyleNum = BorderStyleNum,
                            HorizontalAlignment = h_align,
                            VerticalAlignment = v_align,
                            WrapText = wrap == 1
                        };
        }
        /// <summary> Генерирует стиль на основании данных класса </summary>
        /// <returns></returns>
        public CellFormat GetCellFormat() => new(
                new Alignment() { Horizontal = HorizontalAlignment, Vertical = VerticalAlignment, WrapText = WrapText })
            { FontId = FontStyleNum, FillId = FillStyleNum, BorderId = BorderStyleNum, ApplyFont = true };
    }
}