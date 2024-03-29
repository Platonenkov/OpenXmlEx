﻿using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;

namespace OpenXmlEx.Styles.Base
{
    public record BaseOpenXmlExStyle
    {
        #region Заливка

        /// <summary> Цвет заливки </summary>
        public Color? FillColor { get; set; }
        /// <summary> Стиль заливки </summary>
        public PatternValues? FillPattern { get; set; }

        #endregion

        #region Рамка

        /// <summary> левая рамка </summary>
        public BorderStyleValues? LeftBorderStyle { get; set; }
        /// <summary> верхняя рамка </summary>
        public BorderStyleValues? TopBorderStyle { get; set; }
        /// <summary> правая рамка </summary>
        public BorderStyleValues? RightBorderStyle { get; set; }
        /// <summary> нижняя рамка </summary>
        public BorderStyleValues? BottomBorderStyle { get; set; }
        /// <summary> цвет рамки </summary>
        public Color? BorderColor { get; set; }

        #endregion

        #region Шрифт

        /// <summary> Размер шрифта </summary>
        public double? FontSize { get; set; }
        /// <summary> цвет шрифта </summary>
        public Color? FontColor { get; set; }
        /// <summary> Имя шрифта </summary>
        public string FontName { get; set; }
        /// <summary> жирный или нет </summary>
        public bool? IsBoldFont { get; set; }
        /// <summary> курсивный или нет </summary>
        public bool? IsItalicFont { get; set; }

        #endregion

        #region Align

        /// <summary> будет ли перенос текста в ячейке </summary>
        public bool? WrapText { get; set; }

        /// <summary> Горизонтальное выравнивание в ячейке </summary>
        public HorizontalAlignmentValues? HorizontalAlignment { get; set; }
        /// <summary> Вертикальное выравнивание в ячейке </summary>
        public VerticalAlignmentValues? VerticalAlignment { get; set; }

        /// <summary> выравнивание текста </summary>
        public uint TextRotation { get; set; }

        #endregion


    }
}
