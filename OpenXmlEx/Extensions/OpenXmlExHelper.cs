using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlEx.Extensions
{
    public static class OpenXmlExHelper
    {
        /// <summary> Словарь имен колонок excel </summary>
        private static readonly Dictionary<uint, string> _Columns = new(676);

        /// <summary> Возвращает строковое имя колонки по номеру (1 - А, 2 - В) </summary>
        /// <param name="index">номер колонки</param>
        /// <returns></returns>
        public static string GetColumnName(int index) => GetColumnName((uint)index);

        /// <summary> Возвращает строковое имя колонки по номеру (1 - А, 2 - В) </summary>
        /// <param name="index">номер колонки</param>
        /// <returns></returns>
        public static string GetColumnName(uint index) => GetColumnInfo(index).Value;
        /// <summary> Возвращает строковое имя колонки по номеру (1 - А, 2 - В) </summary>
        /// <param name="index">номер колонки</param>
        /// <returns></returns>
        public static KeyValuePair<uint, string> GetColumnInfo(uint index)
        {
            lock (_Columns)
            {
                var int_col = index - 1; //-1 так как в словаре индексы с 0, а в excel с 1
                if (_Columns.ContainsKey(int_col)) return _Columns.FirstOrDefault(c => c.Key == int_col);
                var int_first_letter = ((int_col) / 676) + 64;
                var int_second_letter = ((int_col % 676) / 26) + 64;
                var int_third_letter = (int_col % 26) + 65;
                var FirstLetter = (int_first_letter > 64) ? (char)int_first_letter : ' ';
                var SecondLetter = (int_second_letter > 64) ? (char)int_second_letter : ' ';
                var ThirdLetter = (char)int_third_letter;
                var s = string.Concat(FirstLetter, SecondLetter, ThirdLetter).Trim();
                var col = new KeyValuePair<uint, string>(int_col, s);
                _Columns.Add(col.Key, col.Value);
                return col;
            }
        }
        /// <summary> Получить адрес ячейки </summary>
        /// <param name="cell">ячейка</param>
        /// <returns></returns>
        public static (uint rowNum, uint collNum) GetCellAddress(Cell cell)
        {

            var cell_ref = cell.CellReference;
            var column_name = string.Empty;
            foreach (var symbol in cell_ref.Value)
            {
                if (!uint.TryParse(symbol.ToString(), out _))
                {
                    column_name += symbol;
                }
                else
                    break;
            }
            var can_get_num = uint.TryParse(cell_ref.Value.Split(column_name).LastOrDefault(), out var row_number);
            if (!can_get_num)
                return default;
            lock (_Columns)
            {
                var col_number = _Columns.FirstOrDefault(c => c.Value == column_name).Key + 1; //+1 так как в словаре индексы с 0, а в excel с 1
                return (row_number, col_number);
            }
        }
        /// <summary> Создаёт запись о группировке для writer </summary>
        /// <param name="lvl">уровень группы</param>
        /// <returns></returns>
        public static OpenXmlAttribute[] GetCollapsedAttributes(uint lvl = 0) => lvl == 0
            ? Array.Empty<OpenXmlAttribute>()
            : new[] { new OpenXmlAttribute("outlineLevel", string.Empty, $"{lvl}"), new OpenXmlAttribute("hidden", string.Empty, $"{lvl}") };

    }
}
