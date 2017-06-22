using ClosedXML.Excel.CalcEngine;
using ClosedXML.Excel.Drawings;
using ClosedXML.Excel.Misc;
using ClosedXML.Extensions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLWorksheet : XLRangeBase, IXLWorksheet
    {
        #region Events

        public XLReentrantEnumerableSet<XLCallbackAction> RangeShiftedRows;
        public XLReentrantEnumerableSet<XLCallbackAction> RangeShiftedColumns;

        #endregion Events

        #region Fields

        private readonly Dictionary<Int32, Int32> _columnOutlineCount = new Dictionary<Int32, Int32>();
        private readonly Dictionary<Int32, Int32> _rowOutlineCount = new Dictionary<Int32, Int32>();
        internal Int32 ZOrder = 1;
        private String _name;
        internal Int32 _position;

        private Double _rowHeight;
        private Boolean _tabActive;
        internal Boolean EventTrackingEnabled;

        #endregion Fields

        #region Constructor

        public XLWorksheet(String sheetName, XLWorkbook workbook)
            : base(
                new XLRangeAddress(
                    new XLAddress(null, XLHelper.MinRowNumber, XLHelper.MinColumnNumber, false, false),
                    new XLAddress(null, XLHelper.MaxRowNumber, XLHelper.MaxColumnNumber, false, false)))
        {
            EventTrackingEnabled = workbook.EventTracking == XLEventTracking.Enabled;

            Workbook = workbook;

            RangeShiftedRows = new XLReentrantEnumerableSet<XLCallbackAction>();
            RangeShiftedColumns = new XLReentrantEnumerableSet<XLCallbackAction>();

            RangeAddress.Worksheet = this;
            RangeAddress.FirstAddress.Worksheet = this;
            RangeAddress.LastAddress.Worksheet = this;

            Pictures = new XLPictures(this);
            NamedRanges = new XLNamedRanges(this);
            SheetView = new XLSheetView();
            Tables = new XLTables();
            Hyperlinks = new XLHyperlinks();
            DataValidations = new XLDataValidations();
            PivotTables = new XLPivotTables();
            Protection = new XLSheetProtection();
            AutoFilter = new XLAutoFilter();
            ConditionalFormats = new XLConditionalFormats();
            SetStyle(workbook.Style);
            Internals = new XLWorksheetInternals(new XLCellsCollection(), new XLColumnsCollection(),
                                                 new XLRowsCollection(), new XLRanges());
            PageSetup = new XLPageSetup((XLPageSetup)workbook.PageOptions, this);
            Outline = new XLOutline(workbook.Outline);
            _columnWidth = workbook.ColumnWidth;
            _rowHeight = workbook.RowHeight;
            RowHeightChanged = Math.Abs(workbook.RowHeight - XLWorkbook.DefaultRowHeight) > XLHelper.Epsilon;
            Name = sheetName;
            SubscribeToShiftedRows((range, rowsShifted) => this.WorksheetRangeShiftedRows(range, rowsShifted));
            SubscribeToShiftedColumns((range, columnsShifted) => this.WorksheetRangeShiftedColumns(range, columnsShifted));
            Charts = new XLCharts();
            ShowFormulas = workbook.ShowFormulas;
            ShowGridLines = workbook.ShowGridLines;
            ShowOutlineSymbols = workbook.ShowOutlineSymbols;
            ShowRowColHeaders = workbook.ShowRowColHeaders;
            ShowRuler = workbook.ShowRuler;
            ShowWhiteSpace = workbook.ShowWhiteSpace;
            ShowZeros = workbook.ShowZeros;
            RightToLeft = workbook.RightToLeft;
            TabColor = XLColor.NoColor;
            SelectedRanges = new XLRanges();

            Author = workbook.Author;
        }

        #endregion Constructor

        //private IXLStyle _style;
        private const String InvalidNameChars = @":\/?*[]";

        public string LegacyDrawingId;
        public Boolean LegacyDrawingIsNew;
        private Double _columnWidth;
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

        public override Boolean UpdatingStyle { get; set; }

        public override IXLStyle InnerStyle
        {
            get { return GetStyle(); }
            set { SetStyle(value); }
        }

        internal Boolean RowHeightChanged { get; set; }
        internal Boolean ColumnWidthChanged { get; set; }

        public Int32 SheetId { get; set; }
        internal String RelId { get; set; }
        public XLDataValidations DataValidations { get; private set; }
        public IXLCharts Charts { get; private set; }
        public XLSheetProtection Protection { get; private set; }
        public XLAutoFilter AutoFilter { get; private set; }

        #region IXLWorksheet Members

        public XLWorkbook Workbook { get; private set; }

        public override IXLStyle Style
        {
            get
            {
                return GetStyle();
            }
            set
            {
                SetStyle(value);
                foreach (XLCell cell in Internals.CellsCollection.GetCells())
                    cell.Style = value;
            }
        }

        public Double ColumnWidth
        {
            get { return _columnWidth; }
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

        public String Name
        {
            get { return _name; }
            set
            {
                if (value.IndexOfAny(InvalidNameChars.ToCharArray()) != -1)
                    throw new ArgumentException("Worksheet names cannot contain any of the following characters: " +
                                                InvalidNameChars);

                if (XLHelper.IsNullOrWhiteSpace(value))
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
                if (value > Workbook.WorksheetsInternal.Count + Workbook.UnsupportedSheets.Count + 1)
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

        IXLRow IXLWorksheet.FirstRowUsed()
        {
            return FirstRowUsed();
        }

        IXLRow IXLWorksheet.FirstRowUsed(Boolean includeFormats)
        {
            return FirstRowUsed(includeFormats);
        }

        IXLRow IXLWorksheet.LastRowUsed()
        {
            return LastRowUsed();
        }

        IXLRow IXLWorksheet.LastRowUsed(Boolean includeFormats)
        {
            return LastRowUsed(includeFormats);
        }

        IXLColumn IXLWorksheet.LastColumn()
        {
            return LastColumn();
        }

        IXLColumn IXLWorksheet.FirstColumn()
        {
            return FirstColumn();
        }

        IXLRow IXLWorksheet.FirstRow()
        {
            return FirstRow();
        }

        IXLRow IXLWorksheet.LastRow()
        {
            return LastRow();
        }

        IXLColumn IXLWorksheet.FirstColumnUsed()
        {
            return FirstColumnUsed();
        }

        IXLColumn IXLWorksheet.FirstColumnUsed(Boolean includeFormats)
        {
            return FirstColumnUsed(includeFormats);
        }

        IXLColumn IXLWorksheet.LastColumnUsed()
        {
            return LastColumnUsed();
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
                    var columnRange = XLHelper.SplitRange(tPair);
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
            return Columns(XLHelper.GetColumnNumberFromLetter(firstColumn),
                           XLHelper.GetColumnNumberFromLetter(lastColumn));
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
                    var rowRange = XLHelper.SplitRange(tPair);
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

        IXLRow IXLWorksheet.Row(Int32 row)
        {
            return Row(row);
        }

        IXLColumn IXLWorksheet.Column(Int32 column)
        {
            return Column(column);
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
            Internals.ColumnsCollection.ForEach(kp => targetSheet.Internals.ColumnsCollection.Add(kp.Key, new XLColumn(kp.Value)));
            Internals.RowsCollection.ForEach(kp => targetSheet.Internals.RowsCollection.Add(kp.Key, new XLRow(kp.Value)));
            Internals.CellsCollection.GetCells().ForEach(c => targetSheet.Cell(c.Address).CopyFrom(c, false));
            DataValidations.ForEach(dv => targetSheet.DataValidations.Add(new XLDataValidation(dv)));
            targetSheet.Visibility = Visibility;
            targetSheet.ColumnWidth = ColumnWidth;
            targetSheet.ColumnWidthChanged = ColumnWidthChanged;
            targetSheet.RowHeight = RowHeight;
            targetSheet.RowHeightChanged = RowHeightChanged;
            targetSheet.SetStyle(Style);
            targetSheet.PageSetup = new XLPageSetup((XLPageSetup)PageSetup, targetSheet);
            (targetSheet.PageSetup.Header as XLHeaderFooter).Changed = true;
            (targetSheet.PageSetup.Footer as XLHeaderFooter).Changed = true;
            targetSheet.Outline = new XLOutline(Outline);
            targetSheet.SheetView = new XLSheetView(SheetView);
            Internals.MergedRanges.ForEach(
                kp => targetSheet.Internals.MergedRanges.Add(targetSheet.Range(kp.RangeAddress.ToString())));

            foreach (var picture in Pictures)
            {
                var newPic = targetSheet.AddPicture(picture.ImageStream, picture.Format, picture.Name)
                    .WithPlacement(picture.Placement)
                    .WithSize(picture.Width, picture.Height);

                switch (picture.Placement)
                {
                    case XLPicturePlacement.FreeFloating:
                        newPic.MoveTo(picture.Left, picture.Top);
                        break;

                    case XLPicturePlacement.Move:
                        var newAddress = new XLAddress(targetSheet, picture.TopLeftCellAddress.RowNumber, picture.TopLeftCellAddress.ColumnNumber, false, false);
                        newPic.MoveTo(newAddress, picture.GetOffset(XLMarkerPosition.TopLeft));
                        break;

                    case XLPicturePlacement.MoveAndSize:
                        var newFromAddress = new XLAddress(targetSheet, picture.TopLeftCellAddress.RowNumber, picture.TopLeftCellAddress.ColumnNumber, false, false);
                        var newToAddress = new XLAddress(targetSheet, picture.BottomRightCellAddress.RowNumber, picture.BottomRightCellAddress.ColumnNumber, false, false);

                        newPic.MoveTo(newFromAddress, picture.GetOffset(XLMarkerPosition.TopLeft), newToAddress, picture.GetOffset(XLMarkerPosition.BottomRight));
                        break;
                }
            }

            foreach (IXLNamedRange r in NamedRanges)
            {
                var ranges = new XLRanges();
                r.Ranges.ForEach(ranges.Add);
                targetSheet.NamedRanges.Add(r.Name, ranges);
            }

            foreach (XLTable t in Tables.Cast<XLTable>())
            {
                String tableName = t.Name;
                var table = targetSheet.Tables.Any(tt => tt.Name == tableName)
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
                    var tableField = table.Field(f) as XLTableField;
                    var tField = t.Field(f) as XLTableField;
                    tableField.Index = tField.Index;
                    tableField.Name = tField.Name;
                    tableField.totalsRowLabel = tField.totalsRowLabel;
                    tableField.totalsRowFunction = tField.totalsRowFunction;
                }
            }

            if (AutoFilter.Enabled)
                targetSheet.Range(AutoFilter.Range.RangeAddress).SetAutoFilter();

            return targetSheet;
        }

        private String ReplaceRelativeSheet(string newSheetName, String value)
        {
            if (XLHelper.IsNullOrWhiteSpace(value)) return value;

            var newValue = new StringBuilder();
            var addresses = value.Split(',');
            foreach (var address in addresses)
            {
                var pair = address.Split('!');
                if (pair.Length == 2)
                {
                    String sheetName = pair[0];
                    if (sheetName.StartsWith("'"))
                        sheetName = sheetName.Substring(1, sheetName.Length - 2);

                    String name = sheetName.ToLower().Equals(Name.ToLower())
                                      ? newSheetName
                                      : sheetName;
                    newValue.Append(String.Format("{0}!{1}", name.WrapSheetNameInQuotesIfRequired(), pair[1]));
                }
                else
                {
                    newValue.Append(address);
                }
            }
            return newValue.ToString();
        }

        public new IXLHyperlinks Hyperlinks { get; private set; }

        IXLDataValidations IXLWorksheet.DataValidations
        {
            get { return DataValidations; }
        }

        private XLWorksheetVisibility _visibility;

        public XLWorksheetVisibility Visibility
        {
            get { return _visibility; }
            set
            {
                if (value != XLWorksheetVisibility.Visible)
                    TabSelected = false;

                _visibility = value;
            }
        }

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

        IXLSheetProtection IXLWorksheet.Protection
        {
            get { return Protection; }
        }

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

        public new IXLRange Sort()
        {
            return GetRangeForSort().Sort();
        }

        public new IXLRange Sort(String columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending,
                                 Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            return GetRangeForSort().Sort(columnsToSortBy, sortOrder, matchCase, ignoreBlanks);
        }

        public new IXLRange Sort(Int32 columnToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending,
                                 Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            return GetRangeForSort().Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);
        }

        public new IXLRange SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false,
                                            Boolean ignoreBlanks = true)
        {
            return GetRangeForSort().SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
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

        public XLColor TabColor { get; set; }

        public IXLWorksheet SetTabColor(XLColor color)
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

        IXLPivotTable IXLWorksheet.PivotTable(String name)
        {
            return PivotTable(name);
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

        public new IXLRanges Ranges(String ranges)
        {
            var retVal = new XLRanges();
            foreach (string rangeAddressStr in ranges.Split(',').Select(s => s.Trim()))
            {
                if (XLHelper.IsValidRangeAddress(rangeAddressStr))
                    retVal.Add(Range(new XLRangeAddress(Worksheet, rangeAddressStr)));
                else if (NamedRanges.Any(n => String.Compare(n.Name, rangeAddressStr, true) == 0))
                    NamedRange(rangeAddressStr).Ranges.ForEach(retVal.Add);
                else
                {
                    Workbook.NamedRanges.First(n =>
                                               String.Compare(n.Name, rangeAddressStr, true) == 0
                                               && n.Ranges.First().Worksheet == this)
                        .Ranges.ForEach(retVal.Add);
                }
            }
            return retVal;
        }

        IXLBaseAutoFilter IXLWorksheet.AutoFilter
        {
            get { return AutoFilter; }
        }

        public IXLRows RowsUsed(Boolean includeFormats = false, Func<IXLRow, Boolean> predicate = null)
        {
            var rows = new XLRows(Worksheet);
            var rowsUsed = new HashSet<Int32>();
            Internals.RowsCollection.Keys.ForEach(r => rowsUsed.Add(r));
            Internals.CellsCollection.RowsUsed.Keys.ForEach(r => rowsUsed.Add(r));
            foreach (var rowNum in rowsUsed)
            {
                var row = Row(rowNum);
                if (!row.IsEmpty(includeFormats) && (predicate == null || predicate(row)))
                    rows.Add(row);
                else
                    row.Dispose();
            }
            return rows;
        }

        public IXLRows RowsUsed(Func<IXLRow, Boolean> predicate = null)
        {
            return RowsUsed(false, predicate);
        }

        public IXLColumns ColumnsUsed(Boolean includeFormats = false, Func<IXLColumn, Boolean> predicate = null)
        {
            var columns = new XLColumns(Worksheet);
            var columnsUsed = new HashSet<Int32>();
            Internals.ColumnsCollection.Keys.ForEach(r => columnsUsed.Add(r));
            Internals.CellsCollection.ColumnsUsed.Keys.ForEach(r => columnsUsed.Add(r));
            foreach (var columnNum in columnsUsed)
            {
                var column = Column(columnNum);
                if (!column.IsEmpty(includeFormats) && (predicate == null || predicate(column)))
                    columns.Add(column);
                else
                    column.Dispose();
            }
            return columns;
        }

        public IXLColumns ColumnsUsed(Func<IXLColumn, Boolean> predicate = null)
        {
            return ColumnsUsed(false, predicate);
        }

        public new void Dispose()
        {
            if (AutoFilter != null)
                AutoFilter.Dispose();

            Internals.Dispose();

            this.Pictures.ForEach(p => p.Dispose());

            base.Dispose();
        }

        #endregion IXLWorksheet Members

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
            var list = _columnOutlineCount.Where(kp => kp.Value > 0).ToList();
            return list.Count == 0 ? 0 : list.Max(kp => kp.Key);
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

        #endregion Outlines

        public HashSet<Int32> GetStyleIds()
        {
            return Internals.CellsCollection.GetStyleIds(GetStyleId());
        }

        public XLRow FirstRowUsed()
        {
            return FirstRowUsed(false);
        }

        public XLRow FirstRowUsed(Boolean includeFormats)
        {
            using (var asRange = AsRange())
            using (var rngRow = asRange.FirstRowUsed(includeFormats))
                return rngRow != null ? Row(rngRow.RangeAddress.FirstAddress.RowNumber) : null;
        }

        public XLRow LastRowUsed()
        {
            return LastRowUsed(false);
        }

        public XLRow LastRowUsed(Boolean includeFormats)
        {
            using (var asRange = AsRange())
            using (var rngRow = asRange.LastRowUsed(includeFormats))
                return rngRow != null ? Row(rngRow.RangeAddress.LastAddress.RowNumber) : null;
        }

        public XLColumn LastColumn()
        {
            return Column(XLHelper.MaxColumnNumber);
        }

        public XLColumn FirstColumn()
        {
            return Column(1);
        }

        public XLRow FirstRow()
        {
            return Row(1);
        }

        public XLRow LastRow()
        {
            return Row(XLHelper.MaxRowNumber);
        }

        public XLColumn FirstColumnUsed()
        {
            return FirstColumnUsed(false);
        }

        public XLColumn FirstColumnUsed(Boolean includeFormats)
        {
            using (var asRange = AsRange())
            using (var rngColumn = asRange.FirstColumnUsed(includeFormats))
                return rngColumn != null ? Column(rngColumn.RangeAddress.FirstAddress.ColumnNumber) : null;
        }

        public XLColumn LastColumnUsed()
        {
            return LastColumnUsed(false);
        }

        public XLColumn LastColumnUsed(Boolean includeFormats)
        {
            using (var asRange = AsRange())
            using (var rngColumn = asRange.LastColumnUsed(includeFormats))
                return rngColumn != null ? Column(rngColumn.RangeAddress.LastAddress.ColumnNumber) : null;
        }

        public XLRow Row(Int32 row)
        {
            return Row(row, true);
        }

        public XLColumn Column(Int32 column)
        {
            if (column <= 0 || column > XLHelper.MaxColumnNumber)
                throw new IndexOutOfRangeException(String.Format("Column number must be between 1 and {0}",
                                                                 XLHelper.MaxColumnNumber));

            Int32 thisStyleId = GetStyleId();
            if (!Internals.ColumnsCollection.ContainsKey(column))
            {
                // This is a new row so we're going to reference all
                // cells in this row to preserve their formatting
                Internals.RowsCollection.Keys.ForEach(r => Cell(r, column));
                Internals.ColumnsCollection.Add(column,
                                                new XLColumn(column, new XLColumnParameters(this, thisStyleId, false)));
            }

            return new XLColumn(column, new XLColumnParameters(this, thisStyleId, true));
        }

        public IXLColumn Column(String column)
        {
            return Column(XLHelper.GetColumnNumberFromLetter(column));
        }

        public override XLRange AsRange()
        {
            return Range(1, 1, XLHelper.MaxRowNumber, XLHelper.MaxColumnNumber);
        }

        public void Clear()
        {
            Internals.CellsCollection.Clear();
            Internals.ColumnsCollection.Clear();
            Internals.MergedRanges.Clear();
            Internals.RowsCollection.Clear();
        }

        private void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
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

            Workbook.Worksheets.ForEach(ws => MoveNamedRangesColumns(range, columnsShifted, ws.NamedRanges));
            MoveNamedRangesColumns(range, columnsShifted, Workbook.NamedRanges);
            ShiftConditionalFormattingColumns(range, columnsShifted);
            ShiftPageBreaksColumns(range, columnsShifted);
        }

        private void ShiftPageBreaksColumns(XLRange range, int columnsShifted)
        {
            for (var i = 0; i < PageSetup.ColumnBreaks.Count; i++)
            {
                int br = PageSetup.ColumnBreaks[i];
                if (range.RangeAddress.FirstAddress.ColumnNumber <= br)
                {
                    PageSetup.ColumnBreaks[i] = br + columnsShifted;
                }
            }
        }

        private void ShiftConditionalFormattingColumns(XLRange range, int columnsShifted)
        {
            Int32 firstColumn = range.RangeAddress.FirstAddress.ColumnNumber;
            if (firstColumn == 1) return;

            Int32 lastColumn = range.RangeAddress.FirstAddress.ColumnNumber + columnsShifted - 1;
            Int32 firstRow = range.RangeAddress.FirstAddress.RowNumber;
            Int32 lastRow = range.RangeAddress.LastAddress.RowNumber;
            var insertedRange = Range(firstRow, firstColumn, lastRow, lastColumn);
            var fc = insertedRange.FirstColumn();
            var model = fc.ColumnLeft();
            Int32 modelFirstRow = model.RangeAddress.FirstAddress.RowNumber;
            if (ConditionalFormats.Any(cf => cf.Range.Intersects(model)))
            {
                for (Int32 ro = firstRow; ro <= lastRow; ro++)
                {
                    var cellModel = model.Cell(ro - modelFirstRow + 1);
                    foreach (var cf in ConditionalFormats.Where(cf => cf.Range.Intersects(cellModel.AsRange())).ToList())
                    {
                        Range(ro, firstColumn, ro, lastColumn).AddConditionalFormat(cf);
                    }
                }
            }
            insertedRange.Dispose();
            model.Dispose();
            fc.Dispose();
        }

        private void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
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
                           && range.RangeAddress.FirstAddress.ColumnNumber <= rngMerged.RangeAddress.LastAddress.ColumnNumber))
                    newMerge.Add(rngMerged);
            }
            Internals.MergedRanges = newMerge;

            Workbook.Worksheets.ForEach(ws => MoveNamedRangesRows(range, rowsShifted, ws.NamedRanges));
            MoveNamedRangesRows(range, rowsShifted, Workbook.NamedRanges);
            ShiftConditionalFormattingRows(range, rowsShifted);
            ShiftPageBreaksRows(range, rowsShifted);
        }

        private void ShiftPageBreaksRows(XLRange range, int rowsShifted)
        {
            for (var i = 0; i < PageSetup.RowBreaks.Count; i++)
            {
                int br = PageSetup.RowBreaks[i];
                if (range.RangeAddress.FirstAddress.RowNumber <= br)
                {
                    PageSetup.RowBreaks[i] = br + rowsShifted;
                }
            }
        }

        private void ShiftConditionalFormattingRows(XLRange range, int rowsShifted)
        {
            Int32 firstRow = range.RangeAddress.FirstAddress.RowNumber;
            if (firstRow == 1) return;

            SuspendEvents();
            var rangeUsed = range.Worksheet.RangeUsed(true);
            IXLRangeAddress usedAddress;
            if (rangeUsed == null)
                usedAddress = range.RangeAddress;
            else
                usedAddress = rangeUsed.RangeAddress;
            ResumeEvents();

            if (firstRow < usedAddress.FirstAddress.RowNumber) firstRow = usedAddress.FirstAddress.RowNumber;

            Int32 lastRow = range.RangeAddress.FirstAddress.RowNumber + rowsShifted - 1;
            if (lastRow > usedAddress.LastAddress.RowNumber) lastRow = usedAddress.LastAddress.RowNumber;

            Int32 firstColumn = range.RangeAddress.FirstAddress.ColumnNumber;
            if (firstColumn < usedAddress.FirstAddress.ColumnNumber) firstColumn = usedAddress.FirstAddress.ColumnNumber;

            Int32 lastColumn = range.RangeAddress.LastAddress.ColumnNumber;
            if (lastColumn > usedAddress.LastAddress.ColumnNumber) lastColumn = usedAddress.LastAddress.ColumnNumber;

            var insertedRange = Range(firstRow, firstColumn, lastRow, lastColumn);
            var fr = insertedRange.FirstRow();
            var model = fr.RowAbove();
            Int32 modelFirstColumn = model.RangeAddress.FirstAddress.ColumnNumber;
            if (ConditionalFormats.Any(cf => cf.Range.Intersects(model)))
            {
                for (Int32 co = firstColumn; co <= lastColumn; co++)
                {
                    var cellModel = model.Cell(co - modelFirstColumn + 1);
                    foreach (var cf in ConditionalFormats.Where(cf => cf.Range.Intersects(cellModel.AsRange())).ToList())
                    {
                        Range(firstRow, co, lastRow, co).AddConditionalFormat(cf);
                    }
                }
            }
            insertedRange.Dispose();
            model.Dispose();
            fr.Dispose();
        }

        internal void BreakConditionalFormatsIntoCells(List<IXLAddress> addresses)
        {
            var newConditionalFormats = new XLConditionalFormats();
            SuspendEvents();
            foreach (var conditionalFormat in ConditionalFormats)
            {
                foreach (XLCell cell in conditionalFormat.Range.Cells(c => !addresses.Contains(c.Address)))
                {
                    var row = cell.Address.RowNumber;
                    var column = cell.Address.ColumnLetter;
                    var newConditionalFormat = new XLConditionalFormat(cell.AsRange(), true);
                    newConditionalFormat.CopyFrom(conditionalFormat);
                    newConditionalFormat.Values.Values.Where(f => f.IsFormula)
                        .ForEach(f => f._value = XLHelper.ReplaceRelative(f.Value, row, column));
                    newConditionalFormats.Add(newConditionalFormat);
                }
                conditionalFormat.Range.Dispose();
            }
            ResumeEvents();
            ConditionalFormats = newConditionalFormats;
        }

        private void MoveNamedRangesRows(XLRange range, int rowsShifted, IXLNamedRanges namedRanges)
        {
            foreach (XLNamedRange nr in namedRanges)
            {
                var newRangeList =
                    nr.RangeList.Select(r => XLCell.ShiftFormulaRows(r, this, range, rowsShifted)).Where(
                        newReference => newReference.Length > 0).ToList();
                nr.RangeList = newRangeList;
            }
        }

        private void MoveNamedRangesColumns(XLRange range, int columnsShifted, IXLNamedRanges namedRanges)
        {
            foreach (XLNamedRange nr in namedRanges)
            {
                var newRangeList =
                    nr.RangeList.Select(r => XLCell.ShiftFormulaColumns(r, this, range, columnsShifted)).Where(
                        newReference => newReference.Length > 0).ToList();
                nr.RangeList = newRangeList;
            }
        }

        public void NotifyRangeShiftedRows(XLRange range, Int32 rowsShifted)
        {
            if (RangeShiftedRows != null)
            {
                foreach (var item in RangeShiftedRows)
                {
                    item.Action(range, rowsShifted);
                }
            }
        }

        public void NotifyRangeShiftedColumns(XLRange range, Int32 columnsShifted)
        {
            if (RangeShiftedColumns != null)
            {
                foreach (var item in RangeShiftedColumns)
                {
                    item.Action(range, columnsShifted);
                }
            }
        }

        public XLRow Row(Int32 row, Boolean pingCells)
        {
            if (row <= 0 || row > XLHelper.MaxRowNumber)
                throw new IndexOutOfRangeException(String.Format("Row number must be between 1 and {0}",
                                                                 XLHelper.MaxRowNumber));

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

        public XLPivotTable PivotTable(String name)
        {
            return (XLPivotTable)PivotTables.PivotTable(name);
        }

        public new IXLCells Cells()
        {
            return Cells(true, true);
        }

        public new IXLCells Cells(Boolean usedCellsOnly)
        {
            if (usedCellsOnly)
                return Cells(true, true);
            else
                return Range(FirstCellUsed(), LastCellUsed()).Cells(false, true);
        }

        public new XLCell Cell(String cellAddressInRange)
        {
            if (XLHelper.IsValidA1Address(cellAddressInRange))
                return Cell(XLAddress.Create(this, cellAddressInRange));

            if (NamedRanges.Any(n => String.Compare(n.Name, cellAddressInRange, true) == 0))
                return (XLCell)NamedRange(cellAddressInRange).Ranges.First().FirstCell();

            var namedRanges = Workbook.NamedRanges.FirstOrDefault(n =>
                                                      String.Compare(n.Name, cellAddressInRange, true) == 0
                                                      && n.Ranges.Count == 1);
            if (namedRanges == null || !namedRanges.Ranges.Any()) return null;

            return (XLCell)namedRanges.Ranges.First().FirstCell();
        }

        internal XLCell CellFast(String cellAddressInRange)
        {
            return Cell(XLAddress.Create(this, cellAddressInRange));
        }

        public override XLRange Range(String rangeAddressStr)
        {
            if (XLHelper.IsValidRangeAddress(rangeAddressStr))
                return Range(new XLRangeAddress(Worksheet, rangeAddressStr));

            if (rangeAddressStr.Contains("["))
                return Table(rangeAddressStr.Substring(0, rangeAddressStr.IndexOf("["))) as XLRange;

            if (NamedRanges.Any(n => String.Compare(n.Name, rangeAddressStr, true) == 0))
                return (XLRange)NamedRange(rangeAddressStr).Ranges.First();

            var namedRanges = Workbook.NamedRanges.FirstOrDefault(n =>
                                                       String.Compare(n.Name, rangeAddressStr, true) == 0
                                                       && n.Ranges.Count == 1
                                                       );
            if (namedRanges == null || !namedRanges.Ranges.Any()) return null;
            return (XLRange)namedRanges.Ranges.First();
        }

        public IXLRanges MergedRanges { get { return Internals.MergedRanges; } }

        public IXLConditionalFormats ConditionalFormats { get; private set; }

        private Boolean _eventTracking;

        public void SuspendEvents()
        {
            _eventTracking = EventTrackingEnabled;
            EventTrackingEnabled = false;
        }

        public void ResumeEvents()
        {
            EventTrackingEnabled = _eventTracking;
        }

        public IXLRanges SelectedRanges { get; internal set; }

        public IXLCell ActiveCell { get; set; }

        private XLCalcEngine _calcEngine;

        private XLCalcEngine CalcEngine
        {
            get { return _calcEngine ?? (_calcEngine = new XLCalcEngine(this)); }
        }

        public Object Evaluate(String expression)
        {
            return CalcEngine.Evaluate(expression);
        }

        public String Author { get; set; }

        public IXLPictures Pictures { get; private set; }

        public IXLPicture Picture(string pictureName)
        {
            return Pictures.Picture(pictureName);
        }

        public IXLPicture AddPicture(Stream stream)
        {
            return Pictures.Add(stream);
        }

        public IXLPicture AddPicture(Stream stream, string name)
        {
            return Pictures.Add(stream, name);
        }

        public Drawings.IXLPicture AddPicture(Stream stream, XLPictureFormat format)
        {
            return Pictures.Add(stream, format);
        }

        public IXLPicture AddPicture(Stream stream, XLPictureFormat format, string name)
        {
            return Pictures.Add(stream, format, name);
        }

        public IXLPicture AddPicture(Bitmap bitmap)
        {
            return Pictures.Add(bitmap);
        }

        public IXLPicture AddPicture(Bitmap bitmap, string name)
        {
            return Pictures.Add(bitmap, name);
        }

        public IXLPicture AddPicture(string imageFile)
        {
            return Pictures.Add(imageFile);
        }

        public IXLPicture AddPicture(string imageFile, string name)
        {
            return Pictures.Add(imageFile, name);
        }
    }
}
