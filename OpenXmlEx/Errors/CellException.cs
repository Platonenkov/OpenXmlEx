using System;

namespace OpenXmlEx.Errors
{
    public class CellException : Exception
    {
        /// <summary> Номер строки вызвавшей ошибку </summary>
        public uint RowNumder { get; set; }
        /// <summary> Номер ячейки вызвавшей ошибку </summary>
        public uint CellNumder { get; set; }
        /// <summary> Строковое название колонки </summary>
        private string _columnName { get; }

        /// <summary> Адрес ячейки в excel </summary>
        public string ExcelAddress => $"{_columnName}{RowNumder}";

        public CellException(string message, uint rowNumder, uint cellNumber, string excelAddress)
            : base(message)
        {
            RowNumder = rowNumder;
            CellNumder = cellNumber;
            _columnName = excelAddress;
        }

        #region Overrides of Exception
        /// <summary>
        /// Вывод информации об ошибки
        /// </summary>
        /// <returns></returns>
        public override string ToString() =>
            $"Cell - {CellNumder}, in row number - {RowNumder}, Address - {ExcelAddress}\n{base.ToString()}, Cell - {CellNumder}, in row number - {RowNumder}";

        #endregion
    }
}