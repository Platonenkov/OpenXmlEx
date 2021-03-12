using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlEx.Styles;
using OpenXmlEx.Styles.Base;
using OpenXmlEx.SubClasses;

namespace OpenXmlEx
{
    public interface IBaseWriter : IDisposable
    {
        #region Base Settings

        /// <summary> Устанавливает тип группировки для строк и столбцов </summary>
        /// <param name="SummaryBelow">группировать сверху (false - сверху, true - снизу)</param>
        /// <param name="SummaryRight">группировать справа (false - справа, true - слева)</param>
        public void SetGrouping(bool SummaryBelow = false, bool SummaryRight = false);

        /// <summary> Устанавливает параметры столбцов </summary>
        /// <param name="settings">список надстроек для листа</param>
        public void SetWidth(IEnumerable<WidthOpenXmlEx> settings);

        /// <summary> отложенная установка фильтра на колонки (ставить в конце листа перед закрытием)</summary>
        /// установит фильтр перед закрытием документа
        /// <param name="ListName">Имя листа</param>
        /// <param name="FirstColumn">первая колонка</param>
        /// <param name="LastColumn">последняя колонка</param>
        /// <param name="FirstRow">первая строка</param>
        /// <param name="LastRow">последняя строка</param>
        public void SetFilter(uint FirstColumn, uint LastColumn, uint FirstRow, uint? LastRow = null, string ListName = null);

        #endregion
        #region Rows


        /// <summary>
        /// Создаёт новую строку в документе
        /// Если предыдущая строка не закрыта - генерирует ошибку
        /// </summary>
        /// <param name="RowIndex">номер новой строки</param>
        /// <param name="CollapsedLvl">уровень группировки - 0 если без группировки</param>
        /// <param name="ClosePreviousIfOpen">задача закрыть предыдущую строку перед созданием новой</param>
        /// <param name="AddSkipedRows">Добавить пропущенные строки (если пишем 2-ю строку, а первую не записали - будет ошибка)</param>
        public void AddRow(uint RowIndex, uint CollapsedLvl = 0, bool ClosePreviousIfOpen = false, bool AddSkipedRows = false);

        /// <summary> Закрыть строку </summary>
        /// <param name="RowNumber">Номер строки</param>
        public void CloseRow(uint RowNumber);

        #endregion

        #region Cells

        /// <summary> Добавляет значение в ячейку документа </summary>
        /// <param name="text">текст для записи</param>
        /// <param name="CellNum">номер колонки</param>
        /// <param name="RowNum">номер строки</param>
        /// <param name="StyleIndex">индекс стиля</param>
        /// <param name="Type">тип данных</param>
        /// <param name="CanReWrite">разрешить перезапись данных (иначе при повторной записи в ячейку будет генерирование ошибки)</param>
        public void AddCell(string text, uint CellNum, uint RowNum, uint StyleIndex = 0, CellValues Type = CellValues.String, bool CanReWrite = false);

        #endregion

        #region MergeCells

        /// <summary> Формирует объединенную ячейку для документа </summary>
        /// <param name="new_range">новый диапазон для объединения</param>
        public void MergeCells(OpenXmlMergedCellEx new_range);

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то также что и начало)</param>
        /// <returns></returns>
        public void MergeCells(int StartCell, int StartRow, int EndCell, int? EndRow = null);

        /// <summary>
        /// Формирует объединенную ячейку для документа
        /// </summary>
        /// <param name="StartCell">колонка начала диапазона</param>
        /// <param name="StartRow">строка начала диапазона</param>
        /// <param name="EndCell">колонка конца диапазона</param>
        /// <param name="EndRow">строка конца диапазона (если не указано то также что и начало)</param>
        /// <returns></returns>
        public void MergeCells(uint StartCell, uint StartRow, uint EndCell, uint? EndRow = null);

        #endregion

        #region Styles

        /// <summary>
        /// Получить стиль и его номер, похожего на искомый
        /// </summary>
        /// <param name="style">искомый стиль</param>
        /// <returns></returns>

        public KeyValuePair<uint, OpenXmlExStyleCell> FindStyleOrDefault(BaseOpenXmlExStyle style);

        #endregion

        /// <summary>
        /// Вызывается для закрытия записи и освобождения документа
        /// </summary>
        public void Close();

    }
}
