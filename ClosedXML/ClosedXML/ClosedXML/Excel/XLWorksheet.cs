using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal delegate void RangeShiftedRowsDelegate(XLRange range, Int32 rowsShifted);

    internal delegate void RangeShiftedColumnsDelegate(XLRange range, Int32 columnsShifted);

    internal class XLWorksheet : XLRangeBase, IXLWorksheet
    {
        #region Constants

        #endregion

        #region Events

        public event RangeShiftedRowsDelegate RangeShiftedRows;
        public event RangeShiftedColumnsDelegate RangeShiftedColumns;

        #endregion

        #region Fields

        private readonly Dictionary<Int32, Int32> _columnOutlineCount = new Dictionary<Int32, Int32>();
        private readonly Dictionary<Int32, Int32> _rowOutlineCount = new Dictionary<Int32, Int32>();
        private String _name;
        internal Int32 _position;

        private Double _rowHeight;
        private IXLSortElements _sortColumns;
        private IXLSortElements _sortRows;
        private Boolean _tabActive;

        #endregion

        #region Constructor

        public XLWorksheet(String sheetName, XLWorkbook workbook)
            : base(
                new XLRangeAddress(
                    new XLAddress(null, ExcelHelper.MinRowNumber, ExcelHelper.MinColumnNumber, false, false),
                    new XLAddress(null, ExcelHelper.MaxRowNumber, ExcelHelper.MaxColumnNumber, false, false)))
        {
            RangeAddress.Worksheet = this;
            RangeAddress.FirstAddress.Worksheet = this;
            RangeAddress.LastAddress.Worksheet = this;

            NamedRanges = new XLNamedRanges(workbook);
            SheetView = new XLSheetView();
            Tables = new XLTables();
            Hyperlinks = new XLHyperlinks();
            DataValidations = new XLDataValidations();
            PivotTables = new XLPivotTables();
            Protection = new XLSheetProtection();
            Workbook = workbook;
            SetStyle(workbook.Style);
            Internals = new XLWorksheetInternals(new XLCellsCollection(), new XLColumnsCollection(),
                                                 new XLRowsCollection(), new XLRanges());
            PageSetup = new XLPageSetup((XLPageSetup)workbook.PageOptions, this);
            Outline = new XLOutline(workbook.Outline);
            _columnWidth = workbook.ColumnWidth;
            _rowHeight = workbook.RowHeight;
            RowHeightChanged = workbook.RowHeight != XLWorkbook.DefaultRowHeight;
            Name = sheetName;
            RangeShiftedRows += XLWorksheetRangeShiftedRows;
            RangeShiftedColumns += XLWorksheetRangeShiftedColumns;
            Charts = new XLCharts();
            ShowFormulas = workbook.ShowFormulas;
            ShowGridLines = workbook.ShowGridLines;
            ShowOutlineSymbols = workbook.ShowOutlineSymbols;
            ShowRowColHeaders = workbook.ShowRowColHeaders;
            ShowRuler = workbook.ShowRuler;
            ShowWhiteSpace = workbook.ShowWhiteSpace;
            ShowZeros = workbook.ShowZeros;
            RightToLeft = workbook.RightToLeft;
            TabColor = new XLColor();
        }

        #endregion

        //private IXLStyle _style;
        public XLWorksheetInternals Internals { get; private set; }

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                yield return GetStyle();
                foreach (XLCell c in Internals.CellsCollection.GetCells())
                    yield return c.Style;
                UpdatingStyle = false;
            }
        }

        public HashSet<Int32> GetStyleIds()
        {
            var retVal = new HashSet<Int32> {GetStyleId()};
            foreach (int id in Internals.CellsCollection.GetCells().Select(c => c.GetStyleId()).Where(id => !retVal.Contains(id)))
            {
                retVal.Add(id);
            }
            return retVal;
        }

        public override Boolean UpdatingStyle { get; set; }

        public override IXLStyle InnerStyle
        {
            get { return GetStyle(); }
            set { SetStyle(value); }
        }

        internal Boolean RowHeightChanged { get; set; }
        internal Boolean ColumnWidthChanged { get; set; }

        public Int32 SheetId { get; set; }
        public String RelId { get; set; }
        public XLDataValidations DataValidations { get; private set; }
        public IXLCharts Charts { get; private set; }

        #region IXLWorksheet Members

        public XLWorkbook Workbook { get; private set; }

        //private Int32 _styleCacheId;
        //public new Int32 GetStyleId()
        //{
        //    if (StyleChanged)
        //        SetStyle(Style);

        //    return _styleCacheId;
        //}
        //private new void SetStyle(IXLStyle styleToUse)
        //{
        //    _styleCacheId = Worksheet.Workbook.GetStyleId(styleToUse);
        //    _style = null;
        //    StyleChanged = false;
        //}
        //private new IXLStyle GetStyle()
        //{
        //    return _style ?? (_style = new XLStyle(this, Worksheet.Workbook.GetStyleById(_styleCacheId)));
        //}

        public override IXLStyle Style
        {
            get { return GetStyle(); }
            set
            {
                SetStyle(value);
                foreach (XLCell cell in Internals.CellsCollection.GetCells())
                    cell.Style = value;
            }
        }

        private Double _columnWidth;
        public string LegacyDrawingId;

        public Double ColumnWidth
        {
            get
            {
                return _columnWidth;
            }
            set
            {
                ColumnWidthChanged = true;
                _columnWidth = value;
            }
        }

        public Double RowHeight
        {
            get { return _rowHeight; }
            set
            {
                RowHeightChanged = true;
                _rowHeight = value;
            }
        }

        private const String InvalidNameChars = @":\/?*[]";
        public String Name
        {
            get { return _name; }
            set
            {
                if (value.IndexOfAny(InvalidNameChars.ToCharArray()) != -1)
                    throw new ArgumentException("Worksheet names cannot contain any of the following characters: " + InvalidNameChars);

                if (StringExtensions.IsNullOrWhiteSpace(value))
                    throw new ArgumentException("Worksheet names cannot be empty");

                if (value.Length > 31)
                    throw new ArgumentException("Worksheet names cannot be more than 31 characters");

                Workbook.WorksheetsInternal.Rename(_name, value);
                _name = value;
            }
        }

        public Int32 Position
        {
            get { return _position; }
            set
            {
                if (value > Workbook.WorksheetsInternal.Count + 1)
                    throw new IndexOutOfRangeException("Index must be equal or less than the number of worksheets + 1.");

                if (value < _position)
                {
                    Workbook.WorksheetsInternal
                        .Where<XLWorksheet>(w => w.Position >= value && w.Position < _position)
                        .ForEach(w => w._position += 1);
                }

                if (value > _position)
                {
                    Workbook.WorksheetsInternal
                        .Where<XLWorksheet>(w => w.Position <= value && w.Position > _position)
                        .ForEach(w => (w)._position -= 1);
                }

                _position = value;
            }
        }

        public IXLPageSetup PageSetup { get; private set; }
        public IXLOutline Outline { get; private set; }

        public XLRow FirstRowUsed()
        {
            return FirstRowUsed(false);
        }

        IXLRow IXLWorksheet.FirstRowUsed()
        {
            return FirstRowUsed();
        }

        public XLRow FirstRowUsed(Boolean includeFormats)
        {
            var rngRow = AsRange().FirstRowUsed(includeFormats);
            return rngRow != null ? Row(rngRow.RangeAddress.FirstAddress.RowNumber) : null;
        }

        IXLRow IXLWorksheet.FirstRowUsed(Boolean includeFormats)
        {
            return FirstRowUsed(includeFormats);
        }

        public XLRow LastRowUsed()
        {
            return LastRowUsed(false);
        }
        IXLRow IXLWorksheet.LastRowUsed()
        {
            return LastRowUsed();
        }
        public XLRow LastRowUsed(Boolean includeFormats)
        {
            var rngRow = AsRange().LastRowUsed(includeFormats);
            return rngRow != null ? Row(rngRow.RangeAddress.LastAddress.RowNumber) : null;
        }
        IXLRow IXLWorksheet.LastRowUsed(Boolean includeFormats)
        {
            return LastRowUsed(includeFormats);
        }

        public XLColumn LastColumn()
        {
            return Column(ExcelHelper.MaxColumnNumber);
        }
        IXLColumn IXLWorksheet.LastColumn()
        {
            return LastColumn();
        }

        public XLColumn FirstColumn()
        {
            return Column(1);
        }
        IXLColumn IXLWorksheet.FirstColumn()
        {
            return FirstColumn();
        }

        public XLRow FirstRow()
        {
            return Row(1);
        }
        IXLRow IXLWorksheet.FirstRow()
        {
            return FirstRow();
        }

        public XLRow LastRow()
        {
            return Row(ExcelHelper.MaxRowNumber);
        }
        IXLRow IXLWorksheet.LastRow()
        {
            return LastRow();
        }

        public XLColumn FirstColumnUsed()
        {
            return FirstColumnUsed(false);
        }
        IXLColumn IXLWorksheet.FirstColumnUsed()
        {
            return FirstColumnUsed();
        }

        public XLColumn FirstColumnUsed(Boolean includeFormats)
        {
            var rngColumn = AsRange().FirstColumnUsed(includeFormats);
            return rngColumn != null ? Column(rngColumn.RangeAddress.FirstAddress.ColumnNumber) : null;
        }
        IXLColumn IXLWorksheet.FirstColumnUsed(Boolean includeFormats)
        {
            return FirstColumnUsed(includeFormats);
        }
        public XLColumn LastColumnUsed()
        {
            return LastColumnUsed(false);
        }
        IXLColumn IXLWorksheet.LastColumnUsed()
        {
            return LastColumnUsed();
        }

        public XLColumn LastColumnUsed(Boolean includeFormats)
        {
            var rngColumn = AsRange().LastColumnUsed(includeFormats);
            return rngColumn != null ? Column(rngColumn.RangeAddress.LastAddress.ColumnNumber) : null;
        }
        IXLColumn IXLWorksheet.LastColumnUsed(Boolean includeFormats)
        {
            return LastColumnUsed(includeFormats);
        }


        public IXLColumns Columns()
        {
            var retVal = new XLColumns(this);
            var columnList = new List<Int32>();

            if (Internals.CellsCollection.Count > 0)
                columnList.AddRange(Internals.CellsCollection.ColumnsUsed.Keys);

            if (Internals.ColumnsCollection.Count > 0)
                columnList.AddRange(Internals.ColumnsCollection.Keys.Where(c => !columnList.Contains(c)));

            foreach (int c in columnList)
                retVal.Add(Column(c));

            return retVal;
        }

        public IXLColumns Columns(String columns)
        {
            var retVal = new XLColumns(null);
            var columnPairs = columns.Split(',');
            foreach (string tPair in columnPairs.Select(pair => pair.Trim()))
            {
                String firstColumn;
                String lastColumn;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    var columnRange = ExcelHelper.SplitRange(tPair);
                    firstColumn = columnRange[0];
                    lastColumn = columnRange[1];
                }
                else
                {
                    firstColumn = tPair;
                    lastColumn = tPair;
                }

                Int32 tmp;
                if (Int32.TryParse(firstColumn, out tmp))
                {
                    foreach (IXLColumn col in Columns(Int32.Parse(firstColumn), Int32.Parse(lastColumn)))
                        retVal.Add((XLColumn)col);
                }
                else
                {
                    foreach (IXLColumn col in Columns(firstColumn, lastColumn))
                        retVal.Add((XLColumn)col);
                }
            }
            return retVal;
        }

        public IXLColumns Columns(String firstColumn, String lastColumn)
        {
            return Columns(ExcelHelper.GetColumnNumberFromLetter(firstColumn),
                           ExcelHelper.GetColumnNumberFromLetter(lastColumn));
        }

        public IXLColumns Columns(Int32 firstColumn, Int32 lastColumn)
        {
            var retVal = new XLColumns(null);

            for (int co = firstColumn; co <= lastColumn; co++)
                retVal.Add(Column(co));
            return retVal;
        }

        public IXLRows Rows()
        {
            var retVal = new XLRows(this);
            var rowList = new List<Int32>();

            if (Internals.CellsCollection.Count > 0)
                rowList.AddRange(Internals.CellsCollection.RowsUsed.Keys);

            if (Internals.RowsCollection.Count > 0)
                rowList.AddRange(Internals.RowsCollection.Keys.Where(r => !rowList.Contains(r)));

            foreach (int r in rowList)
                retVal.Add(Row(r));

            return retVal;
        }

        public IXLRows Rows(String rows)
        {
            var retVal = new XLRows(null);
            var rowPairs = rows.Split(',');
            foreach (string tPair in rowPairs.Select(pair => pair.Trim()))
            {
                String firstRow;
                String lastRow;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    var rowRange = ExcelHelper.SplitRange(tPair);
                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = tPair;
                    lastRow = tPair;
                }
                foreach (IXLRow row in Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
                    retVal.Add((XLRow)row);
            }
            return retVal;
        }

        public IXLRows Rows(Int32 firstRow, Int32 lastRow)
        {
            var retVal = new XLRows(null);

            for (int ro = firstRow; ro <= lastRow; ro++)
                retVal.Add(Row(ro));
            return retVal;
        }

        public XLRow Row(Int32 row)
        {
            return Row(row, true);
        }

        IXLRow IXLWorksheet.Row(Int32 row)
        {
            return Row(row);
        }

        public XLColumn Column(Int32 column)
        {
            if (column <= 0 || column > ExcelHelper.MaxColumnNumber)
                throw new IndexOutOfRangeException(String.Format("Column number must be between 1 and {0}", ExcelHelper.MaxColumnNumber));

            Int32 thisStyleId = GetStyleId();
            if (!Internals.ColumnsCollection.ContainsKey(column))
            {
                // This is a new row so we're going to reference all 
                // cells in this row to preserve their formatting
                Internals.RowsCollection.Keys.ForEach(r => Cell(r, column));
                Internals.ColumnsCollection.Add(column, new XLColumn(column, new XLColumnParameters(this, thisStyleId, false)));
            }

            return new XLColumn(column, new XLColumnParameters(this, thisStyleId, true));
        }

        IXLColumn IXLWorksheet.Column(Int32 column)
        {
            return Column(column);
        }

        public IXLColumn Column(String column)
        {
            return Column(ExcelHelper.GetColumnNumberFromLetter(column));
        }
        IXLColumn IXLWorksheet.Column(String column)
        {
            return Column(column);
        }

        IXLCell IXLWorksheet.Cell(int row, int column)
        {
            return Cell(row, column);
        }

        IXLCell IXLWorksheet.Cell(string cellAddressInRange)
        {
            return Cell(cellAddressInRange);
        }

        IXLCell IXLWorksheet.Cell(int row, string column)
        {
            return Cell(row, column);
        }

        IXLCell IXLWorksheet.Cell(IXLAddress cellAddressInRange)
        {
            return Cell(cellAddressInRange);
        }

        IXLRange IXLWorksheet.Range(IXLRangeAddress rangeAddress)
        {
            return Range(rangeAddress);
        }

        IXLRange IXLWorksheet.Range(string rangeAddress)
        {
            return Range(rangeAddress);
        }

        IXLRange IXLWorksheet.Range(IXLCell firstCell, IXLCell lastCell)
        {
            return Range(firstCell, lastCell);
        }

        IXLRange IXLWorksheet.Range(string firstCellAddress, string lastCellAddress)
        {
            return Range(firstCellAddress, lastCellAddress);
        }

        IXLRange IXLWorksheet.Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            return Range(firstCellAddress, lastCellAddress);
        }

        IXLRange IXLWorksheet.Range(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
        {
            return Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn);
        }

        public override IXLRange AsRange()
        {
            return Range(1, 1, ExcelHelper.MaxRowNumber, ExcelHelper.MaxColumnNumber);
        }

        public IXLWorksheet CollapseRows()
        {
            Enumerable.Range(1, 8).ForEach(i => CollapseRows(i));
            return this;
        }

        public IXLWorksheet CollapseColumns()
        {
            Enumerable.Range(1, 8).ForEach(i => CollapseColumns(i));
            return this;
        }

        public IXLWorksheet ExpandRows()
        {
            Enumerable.Range(1, 8).ForEach(i => ExpandRows(i));
            return this;
        }

        public IXLWorksheet ExpandColumns()
        {
            Enumerable.Range(1, 8).ForEach(i => ExpandRows(i));
            return this;
        }

        public IXLWorksheet CollapseRows(Int32 outlineLevel)
        {
            if (outlineLevel < 1 || outlineLevel > 8)
                throw new ArgumentOutOfRangeException("outlineLevel", "Outline level must be between 1 and 8.");

            Internals.RowsCollection.Values.Where(r => r.OutlineLevel == outlineLevel).ForEach(r => r.Collapse());
            return this;
        }

        public IXLWorksheet CollapseColumns(Int32 outlineLevel)
        {
            if (outlineLevel < 1 || outlineLevel > 8)
                throw new ArgumentOutOfRangeException("outlineLevel", "Outline level must be between 1 and 8.");

            Internals.ColumnsCollection.Values.Where(c => c.OutlineLevel == outlineLevel).ForEach(c => c.Collapse());
            return this;
        }

        public IXLWorksheet ExpandRows(Int32 outlineLevel)
        {
            if (outlineLevel < 1 || outlineLevel > 8)
                throw new ArgumentOutOfRangeException("outlineLevel", "Outline level must be between 1 and 8.");

            Internals.RowsCollection.Values.Where(r => r.OutlineLevel == outlineLevel).ForEach(r => r.Expand());
            return this;
        }

        public IXLWorksheet ExpandColumns(Int32 outlineLevel)
        {
            if (outlineLevel < 1 || outlineLevel > 8)
                throw new ArgumentOutOfRangeException("outlineLevel", "Outline level must be between 1 and 8.");

            Internals.ColumnsCollection.Values.Where(c => c.OutlineLevel == outlineLevel).ForEach(c => c.Expand());
            return this;
        }

        public void Delete()
        {
            Workbook.WorksheetsInternal.Delete(Name);
        }

        public void Clear()
        {
            Internals.CellsCollection.Clear();
            Internals.ColumnsCollection.Clear();
            Internals.MergedRanges.Clear();
            Internals.RowsCollection.Clear();
        }

        public IXLNamedRanges NamedRanges { get; private set; }

        public IXLNamedRange NamedRange(String rangeName)
        {
            return NamedRanges.NamedRange(rangeName);
        }

        public IXLSheetView SheetView { get; private set; }
        public IXLTables Tables { get; private set; }

        public IXLTable Table(Int32 index)
        {
            return Tables.Table(index);
        }
        public IXLTable Table(String name)
        {
            return Tables.Table(name);
        }

        public IXLWorksheet CopyTo(String newSheetName)
        {
            return CopyTo(Workbook, newSheetName, Workbook.WorksheetsInternal.Count + 1);
        }

        public IXLWorksheet CopyTo(String newSheetName, Int32 position)
        {
            return CopyTo(Workbook, newSheetName, position);
        }

        public IXLWorksheet CopyTo(XLWorkbook workbook, String newSheetName)
        {
            return CopyTo(workbook, newSheetName, workbook.WorksheetsInternal.Count + 1);
        }

        public IXLWorksheet CopyTo(XLWorkbook workbook, String newSheetName, Int32 position)
        {
            var targetSheet = (XLWorksheet)workbook.WorksheetsInternal.Add(newSheetName, position);

            Internals.CellsCollection.GetCells().ForEach(c => targetSheet.Cell(c.Address).CopyFrom(c));
            DataValidations.ForEach(dv => targetSheet.DataValidations.Add(new XLDataValidation(dv)));
            Internals.ColumnsCollection.ForEach(
                kp => targetSheet.Internals.ColumnsCollection.Add(kp.Key, new XLColumn(kp.Value)));
            Internals.RowsCollection.ForEach(kp => targetSheet.Internals.RowsCollection.Add(kp.Key, new XLRow(kp.Value)));
            targetSheet.Visibility = Visibility;
            targetSheet.ColumnWidth = ColumnWidth;
            targetSheet.RowHeight = RowHeight;
            targetSheet.SetStyle(Style);
            targetSheet.PageSetup = new XLPageSetup((XLPageSetup)PageSetup, targetSheet);
            targetSheet.Outline = new XLOutline(Outline);
            targetSheet.SheetView = new XLSheetView(SheetView);
            Internals.MergedRanges.ForEach(
                kp => targetSheet.Internals.MergedRanges.Add(targetSheet.Range(kp.RangeAddress.ToString())));

            foreach (IXLNamedRange r in NamedRanges)
            {
                var ranges = new XLRanges();
                r.Ranges.ForEach(ranges.Add);
                targetSheet.NamedRanges.Add(r.Name, ranges);
            }

            foreach (XLTable t in Tables.Cast<XLTable>())
            {
                String tableName = t.Name;
                XLTable table = targetSheet.Tables.Any(tt => tt.Name == tableName) 
                                    ? new XLTable(targetSheet.Range(t.RangeAddress.ToString()), true) 
                                    : new XLTable(targetSheet.Range(t.RangeAddress.ToString()), tableName, true);

                table.RelId = t.RelId;
                table.EmphasizeFirstColumn = t.EmphasizeFirstColumn;
                table.EmphasizeLastColumn = t.EmphasizeLastColumn;
                table.ShowRowStripes = t.ShowRowStripes;
                table.ShowColumnStripes = t.ShowColumnStripes;
                table.ShowAutoFilter = t.ShowAutoFilter;
                table.Theme = t.Theme;
                table._showTotalsRow = t.ShowTotalsRow;
                table._uniqueNames.Clear();

                t._uniqueNames.ForEach(n => table._uniqueNames.Add(n));
                Int32 fieldCount = t.ColumnCount();
                for (Int32 f = 0; f < fieldCount; f++)
                {
                    table.Field(f).Index = t.Field(f).Index;
                    table.Field(f).Name = t.Field(f).Name;
                    (table.Field(f) as XLTableField).totalsRowLabel = (t.Field(f) as XLTableField).totalsRowLabel;
                    (table.Field(f) as XLTableField).totalsRowFunction = (t.Field(f) as XLTableField).totalsRowFunction;
                }
            }

            if (AutoFilterRange != null)
                targetSheet.Range(AutoFilterRange.RangeAddress).SetAutoFilter();

            return targetSheet;
        }

        public new IXLHyperlinks Hyperlinks { get; private set; }

        IXLDataValidations IXLWorksheet.DataValidations
        {
            get { return DataValidations; }
        }

        public XLWorksheetVisibility Visibility { get; set; }

        public IXLWorksheet Hide()
        {
            Visibility = XLWorksheetVisibility.Hidden;
            return this;
        }

        public IXLWorksheet Unhide()
        {
            Visibility = XLWorksheetVisibility.Visible;
            return this;
        }

        public XLSheetProtection Protection { get; private set; }

        IXLSheetProtection IXLWorksheet.Protection { get { return Protection; } }

        public IXLSheetProtection Protect()
        {
            return Protection.Protect();
        }

        public IXLSheetProtection Protect(String password)
        {
            return Protection.Protect(password);
        }

        public IXLSheetProtection Unprotect()
        {
            return Protection.Unprotect();
        }

        public IXLSheetProtection Unprotect(String password)
        {
            return Protection.Unprotect(password);
        }

        public IXLRangeBase AutoFilterRange { get; set; }

        public IXLSortElements SortRows
        {
            get { return _sortRows ?? (_sortRows = new XLSortElements()); }
        }

        public IXLSortElements SortColumns
        {
            get { return _sortColumns ?? (_sortColumns = new XLSortElements()); }
        }

        public IXLRange Sort()
        {
            var range = GetRangeForSort();
            return range.Sort();
        }

        public IXLRange Sort(Boolean matchCase)
        {
            var range = GetRangeForSort();
            return range.Sort(matchCase);
        }

        public IXLRange Sort(XLSortOrder sortOrder)
        {
            var range = GetRangeForSort();
            return range.Sort(sortOrder);
        }

        public IXLRange Sort(XLSortOrder sortOrder, Boolean matchCase)
        {
            var range = GetRangeForSort();
            return range.Sort(sortOrder, matchCase);
        }

        public IXLRange Sort(String columnsToSortBy)
        {
            var range = GetRangeForSort();
            return range.Sort(columnsToSortBy);
        }

        public IXLRange Sort(String columnsToSortBy, Boolean matchCase)
        {
            var range = GetRangeForSort();
            return range.Sort(columnsToSortBy, matchCase);
        }

        public IXLRange Sort(XLSortOrientation sortOrientation)
        {
            var range = GetRangeForSort();
            return range.Sort(sortOrientation);
        }

        public IXLRange Sort(XLSortOrientation sortOrientation, Boolean matchCase)
        {
            var range = GetRangeForSort();
            return range.Sort(sortOrientation, matchCase);
        }

        public IXLRange Sort(XLSortOrientation sortOrientation, XLSortOrder sortOrder)
        {
            var range = GetRangeForSort();
            return range.Sort(sortOrientation, sortOrder);
        }

        public IXLRange Sort(XLSortOrientation sortOrientation, XLSortOrder sortOrder, Boolean matchCase)
        {
            var range = GetRangeForSort();
            return range.Sort(sortOrientation, sortOrder, matchCase);
        }

        public IXLRange Sort(XLSortOrientation sortOrientation, String elementsToSortBy)
        {
            var range = GetRangeForSort();
            return range.Sort(sortOrientation, elementsToSortBy);
        }

        public IXLRange Sort(XLSortOrientation sortOrientation, String elementsToSortBy, Boolean matchCase)
        {
            var range = GetRangeForSort();
            return range.Sort(sortOrientation, elementsToSortBy, matchCase);
        }

        public Boolean ShowFormulas { get; set; }
        public Boolean ShowGridLines { get; set; }
        public Boolean ShowOutlineSymbols { get; set; }
        public Boolean ShowRowColHeaders { get; set; }
        public Boolean ShowRuler { get; set; }
        public Boolean ShowWhiteSpace { get; set; }
        public Boolean ShowZeros { get; set; }

        public IXLWorksheet SetShowFormulas()
        {
            ShowFormulas = true;
            return this;
        }

        public IXLWorksheet SetShowFormulas(Boolean value)
        {
            ShowFormulas = value;
            return this;
        }

        public IXLWorksheet SetShowGridLines()
        {
            ShowGridLines = true;
            return this;
        }

        public IXLWorksheet SetShowGridLines(Boolean value)
        {
            ShowGridLines = value;
            return this;
        }

        public IXLWorksheet SetShowOutlineSymbols()
        {
            ShowOutlineSymbols = true;
            return this;
        }

        public IXLWorksheet SetShowOutlineSymbols(Boolean value)
        {
            ShowOutlineSymbols = value;
            return this;
        }

        public IXLWorksheet SetShowRowColHeaders()
        {
            ShowRowColHeaders = true;
            return this;
        }

        public IXLWorksheet SetShowRowColHeaders(Boolean value)
        {
            ShowRowColHeaders = value;
            return this;
        }

        public IXLWorksheet SetShowRuler()
        {
            ShowRuler = true;
            return this;
        }

        public IXLWorksheet SetShowRuler(Boolean value)
        {
            ShowRuler = value;
            return this;
        }

        public IXLWorksheet SetShowWhiteSpace()
        {
            ShowWhiteSpace = true;
            return this;
        }

        public IXLWorksheet SetShowWhiteSpace(Boolean value)
        {
            ShowWhiteSpace = value;
            return this;
        }

        public IXLWorksheet SetShowZeros()
        {
            ShowZeros = true;
            return this;
        }

        public IXLWorksheet SetShowZeros(Boolean value)
        {
            ShowZeros = value;
            return this;
        }

        public IXLColor TabColor { get; set; }

        public IXLWorksheet SetTabColor(IXLColor color)
        {
            TabColor = color;
            return this;
        }

        public Boolean TabSelected { get; set; }

        public Boolean TabActive
        {
            get { return _tabActive; }
            set
            {
                if (value && !_tabActive)
                {
                    foreach (XLWorksheet ws in Worksheet.Workbook.WorksheetsInternal)
                        ws._tabActive = false;
                }
                _tabActive = value;
            }
        }

        public IXLWorksheet SetTabSelected()
        {
            TabSelected = true;
            return this;
        }

        public IXLWorksheet SetTabSelected(Boolean value)
        {
            TabSelected = value;
            return this;
        }

        public IXLWorksheet SetTabActive()
        {
            TabActive = true;
            return this;
        }

        public IXLWorksheet SetTabActive(Boolean value)
        {
            TabActive = value;
            return this;
        }

        #endregion

        #region Outlines

        public void IncrementColumnOutline(Int32 level)
        {
            if (level <= 0) return;
            if (!_columnOutlineCount.ContainsKey(level))
                _columnOutlineCount.Add(level, 0);

            _columnOutlineCount[level]++;
        }

        public void DecrementColumnOutline(Int32 level)
        {
            if (level <= 0) return;
            if (!_columnOutlineCount.ContainsKey(level))
                _columnOutlineCount.Add(level, 0);

            if (_columnOutlineCount[level] > 0)
                _columnOutlineCount[level]--;
        }

        public Int32 GetMaxColumnOutline()
        {
            return _columnOutlineCount.Count == 0 ? 0 : _columnOutlineCount.Where(kp => kp.Value > 0).Max(kp => kp.Key);
        }

        public void IncrementRowOutline(Int32 level)
        {
            if (level <= 0) return;
            if (!_rowOutlineCount.ContainsKey(level))
                _rowOutlineCount.Add(level, 0);

            _rowOutlineCount[level]++;
        }

        public void DecrementRowOutline(Int32 level)
        {
            if (level <= 0) return;
            if (!_rowOutlineCount.ContainsKey(level))
                _rowOutlineCount.Add(level, 0);

            if (_rowOutlineCount[level] > 0)
                _rowOutlineCount[level]--;
        }

        public Int32 GetMaxRowOutline()
        {
            return _rowOutlineCount.Count == 0 ? 0 : _rowOutlineCount.Where(kp => kp.Value > 0).Max(kp => kp.Key);
        }

        #endregion

        private void XLWorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
        {
            var newMerge = new XLRanges();
            foreach (IXLRange rngMerged in Internals.MergedRanges)
            {
                if (range.RangeAddress.FirstAddress.ColumnNumber <= rngMerged.RangeAddress.FirstAddress.ColumnNumber
                    && rngMerged.RangeAddress.FirstAddress.RowNumber >= range.RangeAddress.FirstAddress.RowNumber
                    && rngMerged.RangeAddress.LastAddress.RowNumber <= range.RangeAddress.LastAddress.RowNumber)
                {
                    var newRng = Range(
                        rngMerged.RangeAddress.FirstAddress.RowNumber,
                        rngMerged.RangeAddress.FirstAddress.ColumnNumber + columnsShifted,
                        rngMerged.RangeAddress.LastAddress.RowNumber,
                        rngMerged.RangeAddress.LastAddress.ColumnNumber + columnsShifted);
                    newMerge.Add(newRng);
                }
                else if (
                    !(range.RangeAddress.FirstAddress.ColumnNumber <= rngMerged.RangeAddress.FirstAddress.ColumnNumber
                      && range.RangeAddress.FirstAddress.RowNumber <= rngMerged.RangeAddress.LastAddress.RowNumber))
                    newMerge.Add(rngMerged);
            }
            Internals.MergedRanges = newMerge;
        }

        private void XLWorksheetRangeShiftedRows(XLRange range, int rowsShifted)
        {
            var newMerge = new XLRanges();
            foreach (IXLRange rngMerged in Internals.MergedRanges)
            {
                if (range.RangeAddress.FirstAddress.RowNumber <= rngMerged.RangeAddress.FirstAddress.RowNumber
                    && rngMerged.RangeAddress.FirstAddress.ColumnNumber >= range.RangeAddress.FirstAddress.ColumnNumber
                    && rngMerged.RangeAddress.LastAddress.ColumnNumber <= range.RangeAddress.LastAddress.ColumnNumber)
                {
                    var newRng = Range(
                        rngMerged.RangeAddress.FirstAddress.RowNumber + rowsShifted,
                        rngMerged.RangeAddress.FirstAddress.ColumnNumber,
                        rngMerged.RangeAddress.LastAddress.RowNumber + rowsShifted,
                        rngMerged.RangeAddress.LastAddress.ColumnNumber);
                    newMerge.Add(newRng);
                }
                else if (!(range.RangeAddress.FirstAddress.RowNumber <= rngMerged.RangeAddress.FirstAddress.RowNumber
                           &&
                           range.RangeAddress.FirstAddress.ColumnNumber <=
                           rngMerged.RangeAddress.LastAddress.ColumnNumber))
                    newMerge.Add(rngMerged);
            }
            Internals.MergedRanges = newMerge;
        }

        public void NotifyRangeShiftedRows(XLRange range, Int32 rowsShifted)
        {
            if (RangeShiftedRows != null)
                RangeShiftedRows(range, rowsShifted);
        }

        public void NotifyRangeShiftedColumns(XLRange range, Int32 columnsShifted)
        {
            if (RangeShiftedColumns != null)
                RangeShiftedColumns(range, columnsShifted);
        }

        public XLRow Row(Int32 row, Boolean pingCells)
        {
            if(row <= 0 || row > ExcelHelper.MaxRowNumber)
                throw  new IndexOutOfRangeException(String.Format("Row number must be between 1 and {0}", ExcelHelper.MaxRowNumber));

            Int32 styleId;
            XLRow rowToUse;
            if (Internals.RowsCollection.TryGetValue(row, out rowToUse))
                styleId = rowToUse.GetStyleId();
            else
            {
                if (pingCells)
                {
                    // This is a new row so we're going to reference all 
                    // cells in columns of this row to preserve their formatting

                    var usedColumns = from c in Internals.ColumnsCollection
                                      join dc in Internals.CellsCollection.ColumnsUsed.Keys
                                          on c.Key equals dc
                                      where !Internals.CellsCollection.Contains(row, dc)
                                      select dc;

                    usedColumns.ForEach(c => Cell(row, c));
                }
                styleId = GetStyleId();
                Internals.RowsCollection.Add(row, new XLRow(row, new XLRowParameters(this, styleId, false)));
            }

            return new XLRow(row, new XLRowParameters(this, styleId));
        }

        private IXLRange GetRangeForSort()
        {
            var range = RangeUsed();
            SortColumns.ForEach(e => range.SortColumns.Add(e.ElementNumber, e.SortOrder, e.IgnoreBlanks, e.MatchCase));
            SortRows.ForEach(e => range.SortRows.Add(e.ElementNumber, e.SortOrder, e.IgnoreBlanks, e.MatchCase));
            return range;
        }

        IXLPivotTable IXLWorksheet.PivotTable(String name)
        {
            return PivotTable(name);
        }
        public XLPivotTable PivotTable(String name)
        {
            return (XLPivotTable)PivotTables.PivotTable(name);
        }
        public IXLPivotTables PivotTables { get; private set; }

        public Boolean RightToLeft { get; set; }

        public IXLWorksheet SetRightToLeft()
        {
            RightToLeft = true;
            return this;
        }

        public IXLWorksheet SetRightToLeft(Boolean value)
        {
            RightToLeft = value;
            return this;
        }

        public new XLCells Cells()
        {
            return CellsUsed(true);
        }

        public new XLCell Cell(String cellAddressInRange)
        {

            if (ExcelHelper.IsValidA1Address(cellAddressInRange))
                return Cell(XLAddress.Create(this, cellAddressInRange));

            if (NamedRanges.Any(n=> String.Compare(n.Name, cellAddressInRange, true) == 0))
                return (XLCell)NamedRange(cellAddressInRange).Ranges.First().FirstCell();

            return (XLCell)Workbook.NamedRanges.First(n =>
                                                    String.Compare(n.Name, cellAddressInRange, true) == 0 
                                                    && n.Ranges.First().Worksheet == this
                                                    && n.Ranges.Count == 1)
                                               .Ranges.First().FirstCell();
        }

        public XLCell CellFast(String cellAddressInRange)
        {
            return Cell(XLAddress.Create(this, cellAddressInRange));
        }

        public override XLRange Range(String rangeAddressStr)
        {
            if (ExcelHelper.IsValidRangeAddress(rangeAddressStr))
                return Range(new XLRangeAddress(Worksheet, rangeAddressStr));

            if (NamedRanges.Any(n => String.Compare(n.Name, rangeAddressStr, true) == 0))
                return (XLRange)NamedRange(rangeAddressStr).Ranges.First();

            return (XLRange)Workbook.NamedRanges.First(n =>
                                                    String.Compare(n.Name, rangeAddressStr, true) == 0
                                                    && n.Ranges.First().Worksheet == this
                                                    && n.Ranges.Count == 1)
                                               .Ranges.First();
        }

        public new IXLRanges Ranges(String ranges)
        {
            var retVal = new XLRanges();
            foreach (var rangeAddressStr in ranges.Split(',').Select(s=>s.Trim()))
            {
                if (ExcelHelper.IsValidRangeAddress(rangeAddressStr))
                    retVal.Add(Range(new XLRangeAddress(Worksheet, rangeAddressStr)));
                else if (NamedRanges.Any(n => String.Compare(n.Name, rangeAddressStr, true) == 0))
                    NamedRange(rangeAddressStr).Ranges.ForEach(retVal.Add);
                else
                    Workbook.NamedRanges.First(n =>
                                                    String.Compare(n.Name, rangeAddressStr, true) == 0
                                                    && n.Ranges.First().Worksheet == this)
                                               .Ranges.ForEach(retVal.Add);
            }
            return retVal;
        }
    }
}