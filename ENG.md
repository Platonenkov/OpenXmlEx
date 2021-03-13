## OpenXmlEx
Installation:
Install-Package OpenXmlEx -Version 1.0.0

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
    using var writer = new EasyWriter(FilePath, styles);
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
