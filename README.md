## OpenXmlEx

### EN
Package wrapper OpenXML to facilitate the recording of heavy documents without allocating a large amount of memory

#### 1 Create styles for the document
```C#
    var styles = new OpenXmlExStyles(
        new List<BaseOpenXmlExStyle>()
        {
            new BaseOpenXmlExStyle() {FontColor = System.Drawing.Color.Crimson, IsBoldFont = true},
            new BaseOpenXmlExStyle() {FontSize = 20, FontName = "Calibri", BorderColor = System.Drawing.Color.Red}
        });
```
but don't forgot that you styles will go after 2 base excel styles;

after you can select style by number from writer
```C#
    var (key, value) = writer.FindStyleOrDefault(
        new BaseOpenXmlExStyle()
        {
            FontColor = System.Drawing.Color.Crimson,
            FontSize = 20,
            IsBoldFont = true,
            LeftBorderStyle = BorderStyleValues.Dashed,
            RightBorderStyle = BorderStyleValues.Dashed
        });
```
Where key is style number;

#### 2 Create new EasyWriter
```C#
    using var writer = new EasyWriter(FileName, styles);
```
#### 3 Add Sheet
```C#
    var sheet_name_1 = "Test_sheet_name";
    writer.AddNewSheet(sheet_name_1);

```
#### 4 Specify how rows are grouped

(set 1 times per sheet)
```C#
    writer.SetGrouping(false, false); // SetGrouping(bool SummaryBelow = false, bool SummaryRight = false)
```
#### 5 Set Width for column

(set 1 times per sheet)
```C#
    var width_setting = new List<WidthOpenXmlEx>
    {
        new (1, 2, 7),
        new (3, 3, 11),
        new (4, 12, 9.5),
        new (13, 13, 17),
        new (14, 14, 40),
        new (15, 16, 15),
        new (18, 20, 15)
    };
    writer.SetWidth(width_setting); //SetWidth(IEnumerable<WidthOpenXmlEx> settings)

```
#### 6 Add Rows 
```C#
    writer.AddRow(3, 0, true, true);
    //AddRow(uint RowIndex, uint CollapsedLvl = 0, bool ClosePreviousIfOpen = false, bool AddSkipedRows = false)
    //CloseRow(uint RowNumber)
```
CollapsedLvl - sets the level of grouping of rows, by default there is no grouping

Use ClosePreviousIfOpen if you want close rows automaticaly when start new, if you start new befor previous will not be closed you get an error.
Use CloseRow if you want total control;

Use AddSkipedRows,or add empty lines manually if you need.
XML does not allow line omissions, for example, a record must go strictly from 1, and after it - the second and so on;

#### 7 AddCell
```C#
    writer.AddCell("Test", 1, 3, 0);
    //AddCell(string text, uint CellNum, uint RowNum, uint StyleIndex = 0, CellValues Type = CellValues.String, bool CanReWrite = false)
```
StyleIndex - style number from styles, but don't forgot that you styles will go after 2 base excel styles; // Use FindStyleOrDefault if you forgot number
CanReWrite - re-writing to a cell will cause an error without this token;
#### 8 SetMergedCell
Set Merged cells range when you need after start new sheet and befor start a new one;
```C#
    writer.MergeCells(6, 3, 10, 5); //MergeCells(int StartCell, int StartRow, int EndCell, int EndRow)
```
#### 9 SetFilter
Set table filter when you need after start new sheet and befor start a new one, but only one filter to sheet!
```C#
    writer.SetFilter(1, 5, 3, 5); //SetFilter(string ListName, uint FirstColumn, uint LastColumn, uint FirstRow, uint LastRow)
```
#### 10 Close or add a new sheet and repeat all steps
the document will add all closing tags when you close the writer, call dispose, or start a new sheet;

### RU
Оболочка для OpenXML, приванная облегчить запись тяжёлых документов без выделения большого количества памяти

#### 1 Создайте стили для документа
```C#
    var styles = new OpenXmlExStyles(
        new List<BaseOpenXmlExStyle>()
        {
            new BaseOpenXmlExStyle() {FontColor = System.Drawing.Color.Crimson, IsBoldFont = true},
            new BaseOpenXmlExStyle() {FontSize = 20, FontName = "Calibri", BorderColor = System.Drawing.Color.Red}
        });
```
Но не забывайте что все ваши стили будут идти с номерами по порядку после двух стандартных стилей

далее вы можете проводить поиск необходимых стилей если забыли номер или сомневаетесь
```C#
    var (key, value) = writer.FindStyleOrDefault(
        new BaseOpenXmlExStyle()
        {
            FontColor = System.Drawing.Color.Crimson,
            FontSize = 20,
            IsBoldFont = true,
            LeftBorderStyle = BorderStyleValues.Dashed,
            RightBorderStyle = BorderStyleValues.Dashed
        });
```
Где key - номер стиля

#### 2 Создайте новый EasyWriter
```C#
    using var writer = new EasyWriter(FileName, styles);
```
#### 3 Добавьте новый лист
```C#
    var sheet_name_1 = "Test_sheet_name";
    writer.AddNewSheet(sheet_name_1);
```
#### 4 Укажите способ группировки строк (если необходимо)


(устанавливается 1 раз на лист)
```C#
    writer.SetGrouping(false, false); // SetGrouping(bool SummaryBelow = false, bool SummaryRight = false)
```
#### 5 Укажите размеры колонок 

(устанавливается 1 раз на лист)
```C#
    var width_setting = new List<WidthOpenXmlEx>
    {
        new (1, 2, 7),
        new (3, 3, 11),
        new (4, 12, 9.5),
        new (13, 13, 17),
        new (14, 14, 40),
        new (15, 16, 15),
        new (18, 20, 15)
    };
    writer.SetWidth(width_setting); //SetWidth(IEnumerable<WidthOpenXmlEx> settings)

```
#### 6 Добавьте строку
```C#
    writer.AddRow(3, 0, true, true);
    //AddRow(uint RowIndex, uint CollapsedLvl = 0, bool ClosePreviousIfOpen = false, bool AddSkipedRows = false)
    //CloseRow(uint RowNumber)
```
CollapsedLvl - устанавливает уровень группировки для строк (вложенность), поумолчанию - 0 (без группировки)

Используйте ClosePreviousIfOpen Если хотите автоматически закрывать строку при создании новой,
если параметр не указан и добавлена новая строка до того как закрыта текущая - будет ошибка записи.

Используйте CloseRow если нужен полный контроль за данными;

используйте AddSkipedRows, или вручную добавляйте пропущенные строки.
XML не поддерживает пропуска строк, данные должны начинаться с 1 строки, после нее должна идти вторая и так далее.

#### 7 Добавьте ячейку
```C#
    writer.AddCell("Test", 1, 3, 0);
    //AddCell(string text, uint CellNum, uint RowNum, uint StyleIndex = 0, CellValues Type = CellValues.String, bool CanReWrite = false)
```
StyleIndex - номер стиля из таблицы которую создали вначале, но если вы не помните номер используйте FindStyleOrDefault,
и не забывайте что ваши стили будут начнаться с номера 2

CanReWrite - перезапись данных в ячейке может вызвать сбой в документе, используйте этот параметр чтобы отключить вызов ошибки;

#### 8 Установите объединенные диапазоны ячеек

Установить диапазон как объединённый можно в любом месте после зодания листа и до закрытия или создания нового
```C#
    writer.MergeCells(6, 3, 10, 5); //MergeCells(int StartCell, int StartRow, int EndCell, int EndRow)
```
#### 9 Установите фильтр

Установите фильтр таблицы если он необходим, вызвать функцию можно в любом месте после содания листа и до закрытия или создания нового, но только 1 раз на лист!

```C#
    writer.SetFilter(1, 5, 3, 5); //SetFilter(string ListName, uint FirstColumn, uint LastColumn, uint FirstRow, uint LastRow)
```
#### 10 Закройте Writer или добавьте новый лист и повторите шаги
Writer добавит все необходимые закрывающие теги когда будет вызван Close, Dispose, или начат новый лист документа.

