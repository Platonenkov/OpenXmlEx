using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXmlEx.SubClasses
{
    /// <summary>
    /// Класс для указания ширины ячеек таблицы
    /// </summary>
    public sealed record WidthOpenXmlEx
    {
        /// <summary> первая ячейка диапазона </summary>
        public uint First { get; set; }
        /// <summary> последняя ячейка диапазона </summary>
        public uint Last { get; set; }
        /// <summary> Ширина ячейки </summary>
        public double Width { get; set; }

        public WidthOpenXmlEx(uint first, uint last, double width)
        {
            First = first;
            Last = last;
            Width = width;
        }

        public WidthOpenXmlEx(uint first, double width) : this(first, first, width)
        {

        }

        public void Deconstruct(out uint first, out uint last, out double widt) => (first, last, widt) = (First, Last, Width);
    }
}
