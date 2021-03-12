using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlEx.Extensions
{
    public static class OpenXmlExMerged
    {
        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то также что и начало)</param>
        /// <returns></returns>
        public static MergeCell MergeCells(int StartCell, int StartRow, int EndCell, int? EndRow = null)
            => new() { Reference = new StringValue($"{OpenXmlExHelper.GetColumnName(StartCell)}{StartRow}:{OpenXmlExHelper.GetColumnName(EndCell)}{EndRow ?? StartRow}") };

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то также что и начало)</param>
        /// <returns></returns>
        public static MergeCell MergeCells(int StartCell, uint StartRow, int EndCell, uint? EndRow = null)
            => new() { Reference = new StringValue($"{OpenXmlExHelper.GetColumnName(StartCell)}{StartRow}:{OpenXmlExHelper.GetColumnName(EndCell)}{EndRow ?? StartRow}") };

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то также что и начало)</param>
        /// <returns></returns>
        public static MergeCell MergeCells(uint StartCell, uint StartRow, uint EndCell, uint? EndRow = null)
            => new() { Reference = new StringValue($"{OpenXmlExHelper.GetColumnName(StartCell)}{StartRow}:{OpenXmlExHelper.GetColumnName(EndCell)}{EndRow ?? StartRow}") };

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то к что и начало)</param>
        /// <returns></returns>
        public static MergeCell MergeCells(uint StartCell, int StartRow, uint EndCell, int? EndRow = null)
            => new() { Reference = new StringValue($"{OpenXmlExHelper.GetColumnName(StartCell)}{StartRow}:{OpenXmlExHelper.GetColumnName(EndCell)}{EndRow ?? StartRow}") };

    }
}
