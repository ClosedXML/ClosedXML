#nullable disable
#nullable enable annotations

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ClosedXML.Excel.CalcEngine.Visitors;
using ClosedXML.Extensions;
using ClosedXML.Parser;

namespace ClosedXML.Excel
{
    internal class XLSparklineGroup : IXLSparklineGroup, ISheetListener
    {
        private readonly XLWorksheet _worksheet;
        private readonly Dictionary<XLSheetPoint, SparklineFormula?> _sparklines = new();
        private IXLRange? _dateRange;
        private IXLSparklineStyle _style;

        #region Public Properties

        public IXLRange? DateRange
        {
            get => _dateRange;
            set => SetDateRange(value);
        }

        public XLDisplayBlanksAsValues DisplayEmptyCellsAs { get; set; }

        public Boolean DisplayHidden { get; set; }

        public IXLSparklineHorizontalAxis HorizontalAxis { get; }

        public Double LineWeight { get; set; }

        public XLSparklineMarkers ShowMarkers { get; set; }

        public IXLSparklineStyle Style
        {
            get => _style;
            set => SetStyle(value);
        }

        public XLSparklineType Type { get; set; }

        public IXLSparklineVerticalAxis VerticalAxis { get; }

        public IXLWorksheet Worksheet => _worksheet;

        #endregion Public Properties

        /// <summary>
        /// A collection of sparkline locations and their formulas.
        /// </summary>
        internal IEnumerable<(XLSheetPoint Location, string? SourceDataFormula)> Sparklines
        {
            get => _sparklines.Select(static sl => (sl.Key, sl.Value?.Text));
        }

        #region Public Constructors

        /// <summary>
        /// Add a new sparkline group copied from an existing sparkline group to the specified worksheet
        /// </summary>
        /// <param name="targetWorksheet">The worksheet the sparkline group is being added to</param>
        /// <param name="copyFrom">The sparkline group to copy from</param>
        /// <returns>The new sparkline group added</returns>
        public XLSparklineGroup(IXLWorksheet targetWorksheet, IXLSparklineGroup copyFrom)
            : this(targetWorksheet)
        {
            CopyFrom(copyFrom);
        }

        /// <summary>
        /// Add a new sparkline group copied from an existing sparkline group to the specified worksheet
        /// </summary>
        /// <returns>The new sparkline group added</returns>
        public XLSparklineGroup(IXLWorksheet targetWorksheet, string locationAddress, string sourceDataAddress)
            : this(targetWorksheet)
        {
            Add(locationAddress, sourceDataAddress);
        }

        /// <summary>
        /// Add a new sparkline group copied from an existing sparkline group to the specified worksheet
        /// </summary>
        /// <returns>The new sparkline group added</returns>
        public XLSparklineGroup(IXLCell location, IXLRange sourceData)
            : this(location.Worksheet)
        {
            Add(location, sourceData);
        }

        /// <summary>
        /// Add a new sparkline group copied from an existing sparkline group to the specified worksheet
        /// </summary>
        /// <returns>The new sparkline group added</returns>
        public XLSparklineGroup(IXLRange locationRange, IXLRange sourceDataRange)
            : this(locationRange.Worksheet)
        {
            Add(locationRange, sourceDataRange);
        }

        #endregion Public Constructors

        #region Public Methods

        public IEnumerable<IXLSparkline> Add(IXLRange locationRange, IXLRange sourceDataRange)
        {
            var singleRow = locationRange.RowCount() == 1;
            var singleColumn = locationRange.ColumnCount() == 1;
            var newSparklines = new List<IXLSparkline>();

            if (singleRow && singleColumn)
            {
                newSparklines.Add(Add(locationRange.FirstCell(), sourceDataRange));
            }
            else if (singleRow)
            {
                if (locationRange.ColumnCount() != sourceDataRange.ColumnCount())
                    throw new ArgumentException("locationRange and sourceDataRange must have the same width");
                for (int i = 1; i <= locationRange.ColumnCount(); i++)
                {
                    newSparklines.Add(Add(locationRange.Cell(1, i), sourceDataRange.Column(i).AsRange()));
                }
            }
            else if (singleColumn)
            {
                if (locationRange.RowCount() != sourceDataRange.RowCount())
                    throw new ArgumentException("locationRange and sourceDataRange must have the same height");

                for (int i = 1; i <= locationRange.RowCount(); i++)
                {
                    newSparklines.Add(Add(locationRange.Cell(i, 1), sourceDataRange.Row(i).AsRange()));
                }
            }
            else
                throw new ArgumentException("locationRange must have either a single row or a single column");

            return newSparklines;
        }

        /// <summary>
        /// Add a sparkline to the group.
        /// </summary>
        /// <param name="location">The cell to add sparklines to. If it already contains a sparkline
        /// it will be replaced.</param>
        /// <param name="sourceData">The range the sparkline gets data from</param>
        /// <returns>A newly created sparkline.</returns>
        public IXLSparkline Add(IXLCell location, IXLRange sourceData)
        {
            if (location.Worksheet != _worksheet)
                throw new ArgumentException("The specified sparkline belongs to the different worksheet");

            // Keep invariant that each cell can have at most one sparkline
            _worksheet.SparklineGroupsInternal.Remove(location);
            var point = XLSheetPoint.FromCell(location);
            AddSparkline(point, sourceData);
            return new XLSparkline(this, point);
        }

        public IEnumerable<IXLSparkline> Add(string locationRangeAddress, string sourceDataAddress)
        {
            var sourceDataRange = _worksheet.Workbook.Range(sourceDataAddress) ??
                                  _worksheet.Range(sourceDataAddress);
            return Add(_worksheet.Range(locationRangeAddress), sourceDataRange);
        }

        /// <summary>
        /// Copy the details from a specified sparkline group
        /// </summary>
        /// <param name="sparklineGroup">The sparkline group to copy from</param>
        public void CopyFrom(IXLSparklineGroup sparklineGroup)
        {
            if (sparklineGroup.DateRange != null)
            {
                DateRange = sparklineGroup.DateRange.Worksheet == sparklineGroup.Worksheet
                    ? _worksheet.Range(sparklineGroup.DateRange.RangeAddress.ToString())
                    : sparklineGroup.DateRange;
            }

            DisplayEmptyCellsAs = sparklineGroup.DisplayEmptyCellsAs;
            DisplayHidden = sparklineGroup.DisplayHidden;
            LineWeight = sparklineGroup.LineWeight;
            ShowMarkers = sparklineGroup.ShowMarkers;
            Type = sparklineGroup.Type;

            XLSparklineStyle.Copy(sparklineGroup.Style, Style);
            XLSparklineHorizontalAxis.Copy(sparklineGroup.HorizontalAxis, HorizontalAxis);
            XLSparklineVerticalAxis.Copy(sparklineGroup.VerticalAxis, VerticalAxis);
        }

        /// <inheritdoc cref="IXLSparklineGroup.CopyTo(IXLWorksheet)"/>
        IXLSparklineGroup IXLSparklineGroup.CopyTo(IXLWorksheet targetSheet)
        {
            return CopyTo((XLWorksheet)targetSheet);
        }

        internal XLSparklineGroup CopyTo(XLWorksheet targetSheet)
        {
            if (targetSheet == _worksheet)
                throw new InvalidOperationException("Cannot copy the sparkline group to the same worksheet it belongs to");

            var groupCopy = new XLSparklineGroup(targetSheet, this);
            targetSheet.SparklineGroupsInternal.Add(groupCopy);
            foreach (var (sparklineLocation, sourceData) in _sparklines)
            {
                var copiedSourceData = sourceData?.CopyFromTo(_worksheet, targetSheet);
                groupCopy._sparklines.Add(sparklineLocation, copiedSourceData);
            }

            return groupCopy;
        }

        public IEnumerator<XLSparkline> GetEnumerator()
        {
            foreach (var sparklinePoint in _sparklines.Keys)
                yield return new XLSparkline(this, sparklinePoint);
        }

        IEnumerator<IXLSparkline> IEnumerable<IXLSparkline>.GetEnumerator()
        {
            return GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLSparkline? GetSparkline(IXLCell cell)
        {
            if (cell.Worksheet != _worksheet)
                return null;

            var location = XLSheetPoint.FromCell(cell);
            if (!_sparklines.ContainsKey(location))
                return null;

            return new XLSparkline(this, location);
        }

        public IEnumerable<IXLSparkline> GetSparklines(IXLRangeBase searchRange)
        {
            if (searchRange.Worksheet != _worksheet)
                yield break;

            var searchArea = XLSheetRange.FromRangeAddress(searchRange.RangeAddress);
            foreach (var location in _sparklines.Keys.Where(searchArea.Contains))
            {
                yield return new XLSparkline(this, location);
            }
        }

        /// <summary>
        /// Remove all sparklines in the specified cell from this group
        /// </summary>
        /// <param name="cell">The cell to remove sparklines from</param>
        public void Remove(IXLCell cell)
        {
            if (cell.Worksheet != _worksheet)
                return;

            Remove(XLSheetPoint.FromCell(cell));
        }

        /// <summary>
        /// Remove the sparkline from this group
        /// </summary>
        /// <param name="sparkline"></param>
        public void Remove(IXLSparkline sparkline)
        {
            Remove(sparkline.Location);
        }

        /// <summary>
        /// Remove all sparklines from this group
        /// </summary>
        public void RemoveAll()
        {
            _sparklines.Clear();
        }

        public IXLSparklineGroup SetDateRange(IXLRange value)
        {
            if (value != null)
            {
                if (value.RowCount() != 1 && value.ColumnCount() != 1)
                    throw new ArgumentException("The date range must be either one row high or one column wide");
            }

            _dateRange = value;
            return this;
        }

        public IXLSparklineGroup SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues displayEmptyCellsAs)
        {
            DisplayEmptyCellsAs = displayEmptyCellsAs;
            return this;
        }

        public IXLSparklineGroup SetDisplayHidden(Boolean displayHidden)
        {
            DisplayHidden = displayHidden;
            return this;
        }

        public IXLSparklineGroup SetLineWeight(Double lineWeight)
        {
            LineWeight = lineWeight;
            return this;
        }

        public IXLSparklineGroup SetShowMarkers(XLSparklineMarkers value)
        {
            ShowMarkers = value;
            return this;
        }

        public IXLSparklineGroup SetStyle(IXLSparklineStyle value)
        {
            _style = value ?? throw new ArgumentNullException(nameof(value));
            return this;
        }

        public IXLSparklineGroup SetType(XLSparklineType type)
        {
            Type = type;
            return this;
        }

        #endregion Public Methods

        /// <summary>
        /// Set sparkline at the location to the specified formula.
        /// </summary>
        internal void SetSparkline(XLSheetPoint location, string? sourceDataFormula)
        {
            _sparklines[location] = !string.IsNullOrWhiteSpace(sourceDataFormula) ? new SparklineFormula(sourceDataFormula) : null;
        }

        internal void Remove(XLSheetPoint location)
        {
            _sparklines.Remove(location);
        }

        internal void MoveSparkline(XLSheetPoint originalLocation, XLSheetPoint sparklineDestination)
        {
            if (!_sparklines.TryGetValue(originalLocation, out var sourceData))
                throw new InvalidOperationException($"No sparkline at the source cell {originalLocation}.");

            // Target can contain sparkline from different group, ensure invariant that only one sparkline per cell
            _worksheet.SparklineGroupsInternal.Remove(sparklineDestination);
            _sparklines.Remove(originalLocation);
            _sparklines[sparklineDestination] = sourceData;
        }

        internal bool TryGetSparkline(XLSheetPoint location, [NotNullWhen(true)] out XLSparkline? sparkline)
        {
            if (!_sparklines.ContainsKey(location))
            {
                sparkline = null;
                return false;
            }

            sparkline = new XLSparkline(this, location);
            return true;
        }

        internal IXLRange? GetSparklineSourceData(XLSheetPoint sparklineLocation)
        {
            if (!_sparklines.TryGetValue(sparklineLocation, out var sourceData))
                throw new InvalidOperationException($"No sparkline at the source cell {sparklineLocation}.");

            // Sparkline formula is always specified with a sheet (or a global name), it doesn't need current worksheet.
            return sourceData is not null ? _worksheet.Workbook.Range(sourceData.Value.Text) : null;
        }

        internal void SetSparklineSourceData(XLSheetPoint sparklineLocation, IXLRange? sourceDataRange)
        {
            if (!_sparklines.Remove(sparklineLocation))
                throw new InvalidOperationException($"No sparkline at the source cell {sparklineLocation}.");

            AddSparkline(sparklineLocation, sourceDataRange);
        }

        internal IXLCell GetLocation(XLSheetPoint sparklineLocation)
        {
            if (!_sparklines.ContainsKey(sparklineLocation))
                throw new InvalidOperationException($"No sparkline at the source cell {sparklineLocation}.");

            return _worksheet.Cell(sparklineLocation);
        }

        private void AddSparkline(XLSheetPoint location, IXLRange? sourceData)
        {
            if (sourceData is not null && sourceData.Worksheet.Workbook != _worksheet.Workbook)
                throw new ArgumentException("Range is from different workbook.");

            if (sourceData is not null && sourceData.RowCount() != 1 && sourceData.ColumnCount() != 1)
                throw new ArgumentException("SourceData range must have either a single row or a single column");

            _sparklines.Add(location, SparklineFormula.From(sourceData));
        }

        #region Private Constructors

        /// <summary>
        /// Add a new sparkline group to the specified worksheet
        /// </summary>
        /// <param name="targetWorksheet">The worksheet the sparkline group is being added to</param>
        /// <returns>The new sparkline group added</returns>
        internal XLSparklineGroup(IXLWorksheet targetWorksheet)
        {
            _worksheet = targetWorksheet as XLWorksheet ?? throw new ArgumentNullException(nameof(targetWorksheet));
            HorizontalAxis = new XLSparklineHorizontalAxis(this);
            VerticalAxis = new XLSparklineVerticalAxis(this);
            HorizontalAxis.Color = XLColor.Black;
            Style = XLSparklineTheme.Default;
            LineWeight = 0.75d;
        }

        #endregion Private Constructors

        // TODO: Sparklines locations should use ST_Sqref semantic for shifting, despite constraint "This sqref element MUST contain exactly one ref element". The code assumes it just shifts individual locations points.
        #region ISheetListner

        void ISheetListener.OnInsertAreaAndShiftDown(XLWorksheet sheet, XLSheetRange insertedArea)
        {
            var insertedBookArea = new XLBookArea(sheet.Name, insertedArea);
            ShiftLocation(insertedBookArea, static (location, insertedArea) =>
            {
                if (!location.InRangeOrBelow(insertedArea))
                    return location;

                var shiftedRow = location.Row + insertedArea.Height;
                if (shiftedRow <= XLHelper.MaxRowNumber)
                    return new XLSheetPoint(shiftedRow, location.Column);

                return null;
            });

            var refMod = new ReferenceShiftOnInsertRefModVisitor(insertedBookArea, true);
            AdjustSourceData(refMod);
        }

        void ISheetListener.OnInsertAreaAndShiftRight(XLWorksheet sheet, XLSheetRange insertedArea)
        {
            var insertedBookArea = new XLBookArea(sheet.Name, insertedArea);
            ShiftLocation(insertedBookArea, static (location, insertedArea) =>
            {
                if (!location.InRangeOrToRight(insertedArea))
                    return location;

                var shiftedColumn = location.Column + insertedArea.Width;
                if (shiftedColumn <= XLHelper.MaxColumnNumber)
                    return new XLSheetPoint(location.Row, shiftedColumn);

                return null;
            });

            var refMod = new ReferenceShiftOnInsertRefModVisitor(insertedBookArea, false);
            AdjustSourceData(refMod);
        }

        void ISheetListener.OnDeleteAreaAndShiftLeft(XLWorksheet sheet, XLSheetRange deletedArea)
        {
            var deletedBookArea = new XLBookArea(sheet.Name, deletedArea);
            ShiftLocation(deletedBookArea, static (location, deletedArea) =>
            {
                if (!location.InRangeOrToRight(deletedArea))
                    return location;

                var shiftedColumn = location.Column - deletedArea.Width;
                if (shiftedColumn >= XLHelper.MinColumnNumber)
                    return new XLSheetPoint(location.Row, shiftedColumn);

                return null;
            });

            var refMod = new ReferenceShiftOnDeleteRefModVisitor(deletedBookArea, XLShiftDeletedCells.ShiftCellsLeft);
            AdjustSourceData(refMod);
        }

        void ISheetListener.OnDeleteAreaAndShiftUp(XLWorksheet sheet, XLSheetRange deletedArea)
        {
            var deletedBookArea = new XLBookArea(sheet.Name, deletedArea);
            ShiftLocation(deletedBookArea, static (location, deletedArea) =>
            {
                if (!location.InRangeOrBelow(deletedArea))
                    return location;

                var shiftedRow = location.Row - deletedArea.Height;
                if (shiftedRow is >= XLHelper.MinRowNumber)
                    return new XLSheetPoint(shiftedRow, location.Column);

                return null;
            });

            var refMod = new ReferenceShiftOnDeleteRefModVisitor(deletedBookArea, XLShiftDeletedCells.ShiftCellsUp);
            AdjustSourceData(refMod);
        }

        private void ShiftLocation(XLBookArea shiftedRange, Func<XLSheetPoint, XLSheetRange, XLSheetPoint?> shiftLocation)
        {
            // If shift was on another worksheet, there is no way to affect sparklines for this worksheet of this group
            if (!XLHelper.SheetComparer.Equals(shiftedRange.Name, _worksheet.Name))
                return;

            var sparklinesCopy = new Dictionary<XLSheetPoint, SparklineFormula?>(_sparklines);

            // Clear to avoid problems during shifting (e.g. A1 and A2 have sparklines, A1 is
            // shifted to A2, but A2 hasn't yet been shifted). Just reinsert everything.
            _sparklines.Clear();
            foreach (var (originalLocation, sourceData) in sparklinesCopy)
            {
                var shiftedLocation = shiftLocation(originalLocation, shiftedRange.Area);
                if (shiftedLocation is not null)
                    _sparklines.Add(shiftedLocation.Value, sourceData);
            }
        }

        private void AdjustSourceData(CopyVisitor refMod)
        {
            // Can't modify dictionary while iterating over it, make a copy.
            var locationsCopy = new List<XLSheetPoint>(_sparklines.Keys);
            foreach (var location in locationsCopy)
            {
                var originalSourceData = _sparklines[location];
                if (originalSourceData is not null)
                {
                    var shiftedSourceData = FormulaConverter.ModifyA1(originalSourceData.Value.Text, _worksheet.Name, location.Row, location.Column, refMod);
                    _sparklines[location] = new SparklineFormula(shiftedSourceData);
                }
            }
        }

        #endregion

        /// <summary>
        /// The source data area referenced by a sparkline. The grammar is should rather limited:
        /// <c>sparkline-formula = single-sheet-area / [single-sheet-prefix / book-prefix] name</c>.
        /// Additionally, if a single-sheet - area is specified, that single-sheet-area MUST contain cells from either
        /// a single row or a single column. In reality, it can be more encompassing (e.g. <c>'[1]Contract Tail YLT'!B46:E46</c>).
        /// </summary>
        /// <param name="Text">Text of the formula.</param>
        private readonly record struct SparklineFormula(string Text)
        {
            /// <summary>
            /// Factory method to create a formula from a reference with a sheet from the range.
            /// </summary>
            [return: NotNullIfNotNull(nameof(range))]
            internal static SparklineFormula? From(IXLRange? range)
            {
                if (range is null)
                    return null;

                var formula = range.RangeAddress.ToStringRelative(true);
                return new SparklineFormula(formula);
            }

            /// <summary>
            /// A factory method used for copying worksheets. If formula is a sheet reference/name,
            /// move the formula of <paramref name="sourceSheet"/> to the <paramref name="targetSheet"/>.
            /// Otherwise, return the original formula.
            /// </summary>
            internal SparklineFormula CopyFromTo(XLWorksheet sourceSheet, XLWorksheet targetSheet)
            {
                // If formula is single-sheet-area, i.e. `single-sheet-prefix A1-area`
                if (ReferenceParser.TryParseSheetA1(Text, out var formulaSheetName, out var reference) &&
                    XLHelper.SheetComparer.Equals(formulaSheetName, sourceSheet.Name))
                {
                    var copiedReference = reference.GetDisplayStringA1(targetSheet.Name);
                    return new SparklineFormula(copiedReference);
                }

                // If formula is a `single-sheet-prefix name`
                if (ReferenceParser.TryParseSheetName(Text, out formulaSheetName, out var definedName) &&
                    XLHelper.SheetComparer.Equals(formulaSheetName, sourceSheet.Name))
                {
                    var copiedName = definedName.GetSheetDefinedName(targetSheet.Name);
                    return new SparklineFormula(copiedName);
                }

                // Either just name or from different workbook
                return this;
            }
        }
    }
}
