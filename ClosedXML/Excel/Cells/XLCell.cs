#nullable disable

using ClosedXML.Excel.InsertData;
using ClosedXML.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Graphics;
using ClosedXML.Parser;

namespace ClosedXML.Excel
{
    [DebuggerDisplay("{Address}")]
    internal sealed class XLCell : XLStylizedBase, IXLCell, IXLStylized
    {
        public static readonly Regex A1SimpleRegex = new Regex(
            //  @"(?<=\W)" // Start with non word
            @"(?<Reference>" // Start Group to pick
            + @"(?<Sheet>" // Start Sheet Name, optional
            + @"("
            + @"\'([^\[\]\*/\\\?:\']+|\'\')\'"
            // Sheet name with special characters, surrounding apostrophes are required
            + @"|"
            + @"\'?\w+\'?" // Sheet name with letters and numbers, surrounding apostrophes are optional
            + @")"
            + @"!)?" // End Sheet Name, optional
            + @"(?<Range>" // Start range
            + @"(?<![\w\d])" // Preceded by anything but a letter or a number
            + @"\$?[a-zA-Z]{1,3}\$?\d{1,7}" // A1 Address 1
            + @"(?<RangeEnd>:\$?[a-zA-Z]{1,3}\$?\d{1,7})?" // A1 Address 2, optional
            + @"(?![\w\d])" // followed by anything but a letter or a number
            + @"|"
            + @"(?<ColumnNumbers>\$?\d{1,7}:\$?\d{1,7})" // 1:1
            + @"|"
            + @"(?<ColumnLetters>\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})" // A:A
            + @")" // End Range
            + @")" // End Group to pick
                   //+ @"(?=\W)" // End with non word
            , RegexOptions.Compiled);

        private static readonly Regex A1RowRegex = new Regex(
            @"(\$?\d{1,7}:\$?\d{1,7})" // 1:1
            , RegexOptions.Compiled);

        private static readonly Regex A1ColumnRegex = new Regex(
            @"(\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})" // A:A
            , RegexOptions.Compiled);

        private static readonly Regex utfPattern = new Regex(@"(?<!_x005F)_x(?!005F)([0-9A-F]{4})_", RegexOptions.Compiled);


        private readonly XLCellsCollection _cellsCollection;

        private readonly int _rowNumber;

        private readonly int _columnNumber;

        internal XLCell(XLWorksheet worksheet, int row, int column)
        {
            _cellsCollection = worksheet.Internals.CellsCollection;
            _rowNumber = row;
            _columnNumber = column;
        }

        internal XLCell(XLWorksheet worksheet, XLSheetPoint point) : this(worksheet, point.Row, point.Column)
        {
        }

        public XLWorksheet Worksheet => _cellsCollection.Worksheet;

        public XLAddress Address => new(Worksheet, _rowNumber, _columnNumber, false, false);

        internal XLSheetPoint SheetPoint => new(_rowNumber, _columnNumber);

        #region Slice fields

        /// <summary>
        /// A flag indicating if a string should be stored in the shared table or inline.
        /// </summary>
        public bool ShareString
        {
            get => _cellsCollection.ValueSlice.GetShareString(SheetPoint);
            set => _cellsCollection.ValueSlice.SetShareString(SheetPoint, value);
        }

        /// <summary>
        /// Overriden <see cref="XLStylizedBase.StyleValue"/>, because we can't store the value
        /// in the cell.
        /// </summary>
        internal override XLStyleValue StyleValue
        {
            get => Worksheet.GetStyleValue(SheetPoint);
            private protected set => _cellsCollection.StyleSlice.Set(_rowNumber, _columnNumber, value);
        }

        internal int SharedStringId => _cellsCollection.ValueSlice.GetShareStringId(SheetPoint);

        internal XLImmutableRichText RichText => SliceRichText;

        private XLCellValue SliceCellValue
        {
            get => _cellsCollection.ValueSlice.GetCellValue(SheetPoint);
            set
            {
                _cellsCollection.ValueSlice.SetCellValue(SheetPoint, value);
                Worksheet.Workbook.CalcEngine.MarkDirty(Worksheet, SheetPoint);
            }
        }

        private XLImmutableRichText SliceRichText
        {
            get => _cellsCollection.ValueSlice.GetRichText(SheetPoint);
            set => _cellsCollection.ValueSlice.SetRichText(SheetPoint, value);
        }

        private XLComment SliceComment
        {
            get => _cellsCollection.MiscSlice[_rowNumber, _columnNumber].Comment;
            set
            {
                ref readonly var original = ref _cellsCollection.MiscSlice[_rowNumber, _columnNumber];
                if (original.Comment != value)
                {
                    var modified = original;
                    modified.Comment = value;
                    _cellsCollection.MiscSlice.Set(_rowNumber, _columnNumber, in modified);
                }
            }
        }

        internal UInt32? CellMetaIndex
        {
            get => _cellsCollection.MiscSlice[_rowNumber, _columnNumber].CellMetaIndex;
            set
            {
                ref readonly var original = ref _cellsCollection.MiscSlice[_rowNumber, _columnNumber];
                if (original.CellMetaIndex != value)
                {
                    var modified = original;
                    modified.CellMetaIndex = value;
                    _cellsCollection.MiscSlice.Set(_rowNumber, _columnNumber, in modified);
                }
            }
        }

        internal UInt32? ValueMetaIndex
        {
            get => _cellsCollection.MiscSlice[_rowNumber, _columnNumber].ValueMetaIndex;
            set
            {
                ref readonly var original = ref _cellsCollection.MiscSlice[_rowNumber, _columnNumber];
                if (original.ValueMetaIndex != value)
                {
                    var modified = original;
                    modified.ValueMetaIndex = value;
                    _cellsCollection.MiscSlice.Set(_rowNumber, _columnNumber, in modified);
                }
            }
        }

        /// <summary>
        /// A formula in the cell. Null, if cell doesn't contain formula.
        /// </summary>
        internal XLCellFormula Formula
        {
            get => _cellsCollection.FormulaSlice.Get(SheetPoint);
            set
            {
                _cellsCollection.FormulaSlice.Set(SheetPoint, value);

                // Because text values of evaluated formulas are stored in a worksheet part, mark it as inlined string and store in sst.
                // If we are clearing formula, we should enable shareString back on, because it is a default position.
                // If we are setting formula, we should disable shareString (=inline), because it must be written to the worksheet part
                var clearFormula = value is null;
                ShareString = clearFormula;
                Worksheet.Workbook.CalcEngine.MarkDirty(Worksheet, SheetPoint);
            }
        }

        #endregion Slice fields

        internal XLComment GetComment()
        {
            return SliceComment ?? CreateComment();
        }

        internal XLComment CreateComment(int? shapeId = null)
        {
            return SliceComment = new XLComment(this, shapeId: shapeId);
        }

        public XLRichText GetRichText()
        {
            var sliceRichText = SliceRichText;
            if (sliceRichText is not null)
                return new XLRichText(this, sliceRichText);

            return CreateRichText();
        }

        public XLRichText CreateRichText()
        {
            var font = new XLFont(GetStyleForRead().Font.Key);

            // Don't include rich text string with 0 length to a new rich text
            var richText = DataType == XLDataType.Blank
                ? new XLRichText(this, font)
                : new XLRichText(this, GetFormattedString(), font);
            SliceRichText = XLImmutableRichText.Create(richText);
            return richText;
        }

        #region IXLCell Members

        IXLWorksheet IXLCell.Worksheet
        {
            get { return Worksheet; }
        }

        IXLAddress IXLCell.Address
        {
            get { return Address; }
        }

        IXLRange IXLCell.AsRange()
        {
            return AsRange();
        }

        internal IXLCell SetValue(XLCellValue value, bool setTableHeader, bool checkMergedRanges)
        {
            if (checkMergedRanges && IsInferiorMergedCell())
                return this;

            SetValueAndStyle(value);

            FormulaA1 = null;

            if (setTableHeader)
            {
                var cellRange = new XLSheetRange(SheetPoint, SheetPoint);
                foreach (var table in Worksheet.Tables)
                    table.RefreshFieldsFromCells(cellRange);
            }

            return this;
        }

        /// <summary>
        /// Set value of a cell and its format (if necessary) from the passed value.
        /// It doesn't clear formulas or checks merged cells or tables.
        /// </summary>
        private void SetValueAndStyle(XLCellValue value)
        {
            var modifiedStyleValue = Worksheet.GetStyleForValue(value, SheetPoint);
            if (modifiedStyleValue is not null)
                StyleValue = modifiedStyleValue;

            // Modify value after style, because we might strip the '
            if (value.Type == XLDataType.Text)
            {
                var text = value.GetText();
                if (text.Length > 0 && text[0] == '\'')
                {
                    value = text.Substring(1);
                }
            }

            SetOnlyValue(value);
        }

        public Boolean GetBoolean() => Value.GetBoolean();

        public Double GetDouble() => Value.GetNumber();

        public string GetText() => Value.GetText();

        public XLError GetError() => Value.GetError();

        public DateTime GetDateTime() => Value.GetDateTime();

        public TimeSpan GetTimeSpan() => Value.GetTimeSpan();

        public Boolean TryGetValue<T>(out T value)
        {
            XLCellValue currentValue;
            try
            {
                currentValue = Value;
            }
            catch
            {
                // May fail for formula evaluation
                value = default;
                return false;
            }

            var targetType = typeof(T);
            var isNullable = targetType.IsNullableType();
            if (isNullable && currentValue.TryConvert(out Blank _))
            {
                value = default;
                return true;
            }

            // JIT compiles a separate version for each T value type and one for all reference types
            // Optimization then removes the double casting for value types.
            var underlyingType = targetType.GetUnderlyingType();
            if (underlyingType == typeof(DateTime) && currentValue.TryConvert(out DateTime dateTime))
            {
                value = (T)(object)dateTime;
                return true;
            }

            var culture = CultureInfo.CurrentCulture;
            if (underlyingType == typeof(TimeSpan) && currentValue.TryConvert(out TimeSpan timeSpan, culture))
            {
                value = (T)(object)timeSpan;
                return true;
            }

            if (underlyingType == typeof(Boolean) && currentValue.TryConvert(out Boolean boolean))
            {
                value = (T)(object)boolean;
                return true;
            }

            if (TryGetStringValue(out value, currentValue)) return true;

            if (underlyingType == typeof(XLError))
            {
                if (currentValue.IsError)
                {
                    value = (T)(object)currentValue.GetError();
                    return true;
                }

                return false;
            }

            // Type code of an enum is a type of an integer, so do this check before numbers
            if (underlyingType.IsEnum)
            {
                var strValue = currentValue.ToString(culture);
                if (Enum.IsDefined(underlyingType, strValue))
                {
                    value = (T)Enum.Parse(underlyingType, strValue, ignoreCase: false);
                    return true;
                }
                value = default;
                return false;
            }

            var typeCode = Type.GetTypeCode(underlyingType);

            // T is a floating point numbers
            if (typeCode >= TypeCode.Single && typeCode <= TypeCode.Decimal)
            {
                if (!currentValue.TryConvert(out Double doubleValue, culture))
                    return false;

                if (typeCode == TypeCode.Single && doubleValue is < Single.MinValue or > Single.MaxValue)
                    return false;

                value = typeCode switch
                {
                    TypeCode.Single => (T)(object)(Single)doubleValue,
                    TypeCode.Double => (T)(object)doubleValue,
                    TypeCode.Decimal => (T)(object)(Decimal)doubleValue,
                    _ => throw new NotSupportedException()
                };
                return true;
            }

            // T is an integer
            if (typeCode >= TypeCode.SByte && typeCode <= TypeCode.UInt64)
            {
                if (!currentValue.TryConvert(out Double doubleValue, culture))
                    return false;

                if (!doubleValue.Equals(Math.Truncate(doubleValue)))
                    return false;

                var valueIsWithinBounds = typeCode switch
                {
                    TypeCode.SByte => doubleValue >= SByte.MinValue && doubleValue <= SByte.MaxValue,
                    TypeCode.Byte => doubleValue >= Byte.MinValue && doubleValue <= Byte.MaxValue,
                    TypeCode.Int16 => doubleValue >= Int16.MinValue && doubleValue <= Int16.MaxValue,
                    TypeCode.UInt16 => doubleValue >= UInt16.MinValue && doubleValue <= UInt16.MaxValue,
                    TypeCode.Int32 => doubleValue >= Int32.MinValue && doubleValue <= Int32.MaxValue,
                    TypeCode.UInt32 => doubleValue >= UInt32.MinValue && doubleValue <= UInt32.MaxValue,
                    TypeCode.Int64 => doubleValue >= Int64.MinValue && doubleValue <= Int64.MaxValue,
                    TypeCode.UInt64 => doubleValue >= UInt64.MinValue && doubleValue <= UInt64.MaxValue,
                    _ => throw new NotSupportedException()
                };
                if (!valueIsWithinBounds)
                    return false;

                value = typeCode switch
                {
                    TypeCode.SByte => (T)(object)(SByte)doubleValue,
                    TypeCode.Byte => (T)(object)(Byte)doubleValue,
                    TypeCode.Int16 => (T)(object)(Int16)doubleValue,
                    TypeCode.UInt16 => (T)(object)(UInt16)doubleValue,
                    TypeCode.Int32 => (T)(object)(Int32)doubleValue,
                    TypeCode.UInt32 => (T)(object)(UInt32)doubleValue,
                    TypeCode.Int64 => (T)(object)(Int64)doubleValue,
                    TypeCode.UInt64 => (T)(object)(UInt64)doubleValue,
                    _ => throw new NotSupportedException()
                };
                return true;
            }

            return false;
        }

        private static bool TryGetStringValue<T>(out T value, XLCellValue currentValue)
        {
            if (typeof(T) == typeof(String))
            {
                var s = currentValue.ToString(CultureInfo.CurrentCulture);
                var matches = utfPattern.Matches(s);

                if (matches.Count == 0)
                {
                    value = (T)Convert.ChangeType(s, typeof(T));
                    return true;
                }

                var sb = new StringBuilder();
                var lastIndex = 0;

                foreach (var match in matches.Cast<Match>())
                {
                    var matchString = match.Value;
                    var matchIndex = match.Index;
                    sb.Append(s.Substring(lastIndex, matchIndex - lastIndex));

                    sb.Append((char)int.Parse(match.Groups[1].Value, NumberStyles.AllowHexSpecifier));

                    lastIndex = matchIndex + matchString.Length;
                }

                if (lastIndex < s.Length)
                    sb.Append(s.Substring(lastIndex));

                value = (T)Convert.ChangeType(sb.ToString(), typeof(T));
                return true;
            }
            value = default;
            return false;
        }

        public T GetValue<T>()
        {
            if (TryGetValue(out T retVal))
                return retVal;

            throw new InvalidCastException($"Cannot convert {Address.ToStringRelative(true)}'s value to " + typeof(T));
        }

        public String GetString() => Value.ToString(CultureInfo.CurrentCulture);

        public string GetFormattedString(CultureInfo culture = null)
        {
            XLCellValue value;
            try
            {
                // Need to get actual value because formula might be out of date or value wasn't set at all
                // Unimplemented functions and features throw exceptions
                value = Value;
            }
            catch
            {
                value = CachedValue;
            }

            return GetFormattedString(value, culture);
        }

        internal string GetFormattedString(XLCellValue value, CultureInfo culture = null)
        {
            culture ??= CultureInfo.CurrentCulture;
            var format = GetFormat();
            return value.IsUnifiedNumber
                ? value.GetUnifiedNumber().ToExcelFormat(format, culture)
                : value.ToString(culture);
        }

        public void InvalidateFormula()
        {
            if (Formula is null)
            {
                return;
            }

            Formula.IsDirty = true;
        }

        /// <summary>
        /// Perform an evaluation of cell formula. If cell does not contain formula nothing happens, if cell does not need
        /// recalculation (<see cref="NeedsRecalculation"/> is False) nothing happens either, unless <paramref name="force"/> flag is specified.
        /// Otherwise recalculation is performed, result value is preserved in <see cref="CachedValue"/> and returned.
        /// </summary>
        /// <param name="force">Flag indicating whether a recalculation must be performed even is cell does not need it.</param>
        /// <returns>Null if cell does not contain a formula. Calculated value otherwise.</returns>
        public void Evaluate(Boolean force)
        {
            if (Formula is null)
            {
                return;
            }

            var shouldRecalculate = force || NeedsRecalculation;
            if (!shouldRecalculate)
            {
                return;
            }

            // TODO: Only one cell, somehow
            var wb = Worksheet.Workbook;
            wb.CalcEngine.Recalculate(wb, null);
        }

        /// <summary>
        /// Set only value, don't clear formula, don't set format.
        /// Sets the value even for merged cells.
        /// </summary>
        internal void SetOnlyValue(XLCellValue value)
        {
            SliceCellValue = value;
        }

        public IXLCell SetValue(XLCellValue value)
        {
            return SetValue(value, true, true);
        }

        public override string ToString() => ToString("A");

        public string ToString(string format)
        {
            return (format.ToUpper()) switch
            {
                "A" => this.Address.ToString(),
                "F" => HasFormula ? this.FormulaA1 : string.Empty,
                "NF" => Style.NumberFormat.Format,
                "FG" => Style.Font.FontColor.ToString(),
                "BG" => Style.Fill.BackgroundColor.ToString(),
                "V" => GetFormattedString(),
                _ => throw new FormatException($"Format {format} was not recognised."),
            };
        }

        public XLCellValue Value
        {
            get
            {
                if (Formula is not null)
                {
                    Evaluate(false);
                }

                return SliceCellValue;
            }
            set => SetValue(value);
        }

        public IXLTable InsertTable<T>(IEnumerable<T> data)
        {
            return InsertTable(data, null, true);
        }

        public IXLTable InsertTable<T>(IEnumerable<T> data, bool createTable)
        {
            return InsertTable(data, null, createTable);
        }

        public IXLTable InsertTable<T>(IEnumerable<T> data, string tableName)
        {
            return InsertTable(data, tableName, true);
        }

        public IXLTable InsertTable<T>(IEnumerable<T> data, String tableName, Boolean createTable)
        {
            return InsertTable(data, tableName, createTable, addHeadings: true, transpose: false);
        }

        public IXLTable InsertTable<T>(IEnumerable<T> data, String tableName, Boolean createTable, Boolean addHeadings, Boolean transpose)
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);
            return Worksheet.InsertTable(SheetPoint, reader, tableName, createTable, addHeadings, transpose);
        }

        public IXLTable InsertTable(DataTable data)
        {
            return InsertTable(data, null, true);
        }

        public IXLTable InsertTable(DataTable data, Boolean createTable)
        {
            return InsertTable(data, null, createTable);
        }

        public IXLTable InsertTable(DataTable data, String tableName)
        {
            return InsertTable(data, tableName, true);
        }

        public IXLTable InsertTable(DataTable data, String tableName, Boolean createTable)
        {
            if (data == null || data.Columns.Count == 0)
                return null;

            if (XLHelper.IsValidA1Address(tableName) || XLHelper.IsValidRCAddress(tableName))
                throw new InvalidOperationException($"Table name cannot be a valid Cell Address '{tableName}'.");

            if (createTable && this.Worksheet.Tables.Any<XLTable>(t => t.Contains(this)))
                throw new InvalidOperationException($"This cell '{this.Address}' is already part of a table.");

            var reader = InsertDataReaderFactory.Instance.CreateReader(data);
            return Worksheet.InsertTable(SheetPoint, reader, tableName, createTable, addHeadings: true, transpose: false);
        }

        public XLTableCellType TableCellType()
        {
            var table = this.Worksheet.Tables.FirstOrDefault<XLTable>(t => t.AsRange().Contains(this));
            if (table == null) return XLTableCellType.None;

            if (table.ShowHeaderRow && table.HeadersRow().RowNumber().Equals(this._rowNumber)) return XLTableCellType.Header;
            if (table.ShowTotalsRow && table.TotalsRow().RowNumber().Equals(this._rowNumber)) return XLTableCellType.Total;

            return XLTableCellType.Data;
        }

        public IXLRange InsertData(IEnumerable data)
        {
            if (data == null || data is String)
                return null;

            return InsertData(data, transpose: false);
        }

        public IXLRange InsertData(IEnumerable data, Boolean transpose)
        {
            if (data == null || data is String)
                return null;

            var reader = InsertDataReaderFactory.Instance.CreateReader(data);
            return Worksheet.InsertData(SheetPoint, reader, addHeadings: false, transpose: transpose);
        }

        public IXLRange InsertData(DataTable dataTable)
        {
            if (dataTable == null)
                return null;

            var reader = InsertDataReaderFactory.Instance.CreateReader(dataTable);
            return Worksheet.InsertData(SheetPoint, reader, addHeadings: false, transpose: false);
        }

        public XLDataType DataType => SliceCellValue.Type;

        public IXLCell Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            return Clear(clearOptions, false);
        }

        internal IXLCell Clear(XLClearOptions clearOptions, bool calledFromRange)
        {
            //Note: We have to check if the cell is part of a merged range. If so we have to clear the whole range
            //Checking if called from range to avoid stack overflow
            if (!calledFromRange && IsMerged())
            {
                var firstOrDefault = Worksheet.Internals.MergedRanges.GetIntersectedRanges(Address).FirstOrDefault();
                if (firstOrDefault != null)
                    firstOrDefault.Clear(clearOptions);
            }
            else
            {
                if (clearOptions.HasFlag(XLClearOptions.Contents))
                {
                    SetHyperlink(null);
                    SliceCellValue = Blank.Value;
                    FormulaA1 = String.Empty;
                }

                if (clearOptions.HasFlag(XLClearOptions.NormalFormats))
                    SetStyle(Worksheet.Style);

                if (clearOptions.HasFlag(XLClearOptions.ConditionalFormats))
                {
                    AsRange().RemoveConditionalFormatting();
                }

                if (clearOptions.HasFlag(XLClearOptions.Comments))
                    SliceComment = null;

                if (clearOptions.HasFlag(XLClearOptions.Sparklines))
                {
                    AsRange().RemoveSparklines();
                }

                if (clearOptions.HasFlag(XLClearOptions.DataValidation) && HasDataValidation)
                {
                    var validation = CreateDataValidation();
                    Worksheet.DataValidations.Delete(validation);
                }

                if (clearOptions.HasFlag(XLClearOptions.MergedRanges) && IsMerged())
                {
                    ClearMerged();
                }
            }

            return this;
        }

        public void Delete(XLShiftDeletedCells shiftDeleteCells)
        {
            Worksheet.Range(Address, Address).Delete(shiftDeleteCells);
        }

        public string FormulaA1
        {
            get => Formula?.A1 ?? String.Empty;

            set
            {
                if (IsInferiorMergedCell())
                    return;

                value = value?.TrimFormulaEqual();
                Formula = !String.IsNullOrWhiteSpace(value)
                    ? XLCellFormula.NormalA1(value)
                    : null;
                InvalidateFormula();
            }
        }

        public string FormulaR1C1
        {
            get => Formula?.GetFormulaR1C1(SheetPoint) ?? String.Empty;

            set
            {
                if (IsInferiorMergedCell())
                    return;

                value = value?.TrimFormulaEqual();
                Formula = !String.IsNullOrWhiteSpace(value)
                    ? XLCellFormula.NormalA1(FormulaConverter.ToA1(value, _rowNumber, _columnNumber))
                    : null;
                InvalidateFormula();
            }
        }

        public XLHyperlink GetHyperlink()
        {
            if (Worksheet.Hyperlinks.TryGet(SheetPoint, out var hyperlink))
                return hyperlink;

            return CreateHyperlink();
        }

#nullable enable
        /// <inheritdoc />
        public void SetHyperlink(XLHyperlink? hyperlink)
        {
            if (Worksheet.Hyperlinks.TryGet(SheetPoint, out var existingHyperlink))
                Worksheet.Hyperlinks.Delete(existingHyperlink);

            if (hyperlink is null)
                return;

            Worksheet.Hyperlinks.Add(SheetPoint, hyperlink);

            if (GetStyleForRead().Font.FontColor.Equals(Worksheet.StyleValue.Font.FontColor))
                Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);

            if (GetStyleForRead().Font.Underline == Worksheet.StyleValue.Font.Underline)
                Style.Font.Underline = XLFontUnderlineValues.Single;
        }

        internal void SetCellHyperlink(XLHyperlink hyperlink)
        {
            Worksheet.Hyperlinks.Clear(SheetPoint);
            Worksheet.Hyperlinks.Add(SheetPoint, hyperlink);
        }
#nullable disable

        public XLHyperlink CreateHyperlink()
        {
            SetHyperlink(new XLHyperlink());
            return GetHyperlink();
        }

        public IXLCells InsertCellsAbove(int numberOfRows)
        {
            return AsRange().InsertRowsAbove(numberOfRows).Cells();
        }

        public IXLCells InsertCellsBelow(int numberOfRows)
        {
            return AsRange().InsertRowsBelow(numberOfRows).Cells();
        }

        public IXLCells InsertCellsAfter(int numberOfColumns)
        {
            return AsRange().InsertColumnsAfter(numberOfColumns).Cells();
        }

        public IXLCells InsertCellsBefore(int numberOfColumns)
        {
            return AsRange().InsertColumnsBefore(numberOfColumns).Cells();
        }

        public IXLCell AddToNamed(string rangeName)
        {
            AsRange().AddToNamed(rangeName);
            return this;
        }

        public IXLCell AddToNamed(string rangeName, XLScope scope)
        {
            AsRange().AddToNamed(rangeName, scope);
            return this;
        }

        public IXLCell AddToNamed(string rangeName, XLScope scope, string comment)
        {
            AsRange().AddToNamed(rangeName, scope, comment);
            return this;
        }

        /// <summary>
        /// Flag indicating that previously calculated cell value may be not valid anymore and has to be re-evaluated.
        /// </summary>
        public bool NeedsRecalculation => Formula is not null && Formula.IsDirty;

        public XLCellValue CachedValue => SliceCellValue;

        IXLRichText IXLCell.GetRichText() => GetRichText();

        public bool HasRichText => SliceRichText is not null;

        IXLRichText IXLCell.CreateRichText() => CreateRichText();

        IXLComment IXLCell.GetComment() => GetComment();

        public bool HasComment
        {
            get { return SliceComment != null; }
        }

        IXLComment IXLCell.CreateComment()
        {
            return CreateComment(shapeId: null);
        }

        public Boolean IsMerged()
        {
            return Worksheet.Internals.MergedRanges.Contains(this);
        }

        public IXLRange MergedRange()
        {
            return Worksheet
                .Internals
                .MergedRanges
                .GetIntersectedRanges(this)
                .FirstOrDefault();
        }

        public Boolean IsEmpty()
        {
            return IsEmpty(XLCellsUsedOptions.AllContents);
        }

        public Boolean IsEmpty(XLCellsUsedOptions options)
        {
            if (options.HasFlag (XLCellsUsedOptions.Contents))
            {
                var isValueEmpty = SliceCellValue.Type switch
                {
                    XLDataType.Blank => true,
                    XLDataType.Text => SliceCellValue.GetText().Length == 0,
                    _ => false
                };

                if (!isValueEmpty || HasFormula)
                    return false;
            }

            if (options.HasFlag(XLCellsUsedOptions.NormalFormats))
            {
                if (StyleValue.IncludeQuotePrefix)
                    return false;

                if (!StyleValue.Equals(Worksheet.StyleValue))
                    return false;

                if (StyleValue.Equals(Worksheet.StyleValue))
                {
                    if (Worksheet.Internals.RowsCollection.TryGetValue(_rowNumber, out XLRow row) && !row.StyleValue.Equals(Worksheet.StyleValue))
                        return false;

                    if (Worksheet.Internals.ColumnsCollection.TryGetValue(_columnNumber, out XLColumn column) && !column.StyleValue.Equals(Worksheet.StyleValue))
                        return false;
                }
            }

            if (options.HasFlag(XLCellsUsedOptions.MergedRanges) && IsMerged())
                return false;

            if (options.HasFlag(XLCellsUsedOptions.Comments) && HasComment)
                return false;

            if (options.HasFlag(XLCellsUsedOptions.DataValidation) && HasDataValidation)
                return false;

            if (options.HasFlag(XLCellsUsedOptions.ConditionalFormats)
                && Worksheet.ConditionalFormats.SelectMany(cf => cf.Ranges).Any(range => range.Contains(this)))
                return false;

            if (options.HasFlag(XLCellsUsedOptions.Sparklines) && HasSparkline)
                return false;

            return true;
        }

        public IXLColumn WorksheetColumn()
        {
            return Worksheet.Column(_columnNumber);
        }

        public IXLRow WorksheetRow()
        {
            return Worksheet.Row(_rowNumber);
        }

        public IXLCell CopyTo(IXLCell target)
        {
            (target as XLCell).CopyFrom(this, XLCellCopyOptions.All);
            return target;
        }

        public IXLCell CopyTo(String target)
        {
            return CopyTo(GetTargetCell(target, Worksheet));
        }

        public IXLCell CopyFrom(IXLCell otherCell)
        {
            return CopyFrom(otherCell as XLCell, XLCellCopyOptions.All);
        }

        public IXLCell CopyFrom(String otherCell)
        {
            return CopyFrom(GetTargetCell(otherCell, Worksheet));
        }

        public IXLCell SetFormulaA1(String formula)
        {
            FormulaA1 = formula;
            return this;
        }

        public IXLCell SetFormulaR1C1(String formula)
        {
            FormulaR1C1 = formula;
            return this;
        }

        public Boolean HasSparkline => Sparkline != null;

        /// <summary> The sparkline assigned to the cell </summary>
        public IXLSparkline Sparkline => Worksheet.SparklineGroups.GetSparkline(this);

        public IXLDataValidation GetDataValidation()
        {
            return FindDataValidation() ?? CreateDataValidation();
        }

        public Boolean HasDataValidation
        {
            get { return FindDataValidation() != null; }
        }

        /// <summary>
        /// Get the data validation rule containing current cell.
        /// </summary>
        /// <returns>The data validation rule applying to the current cell or null if there is no such rule.</returns>
        private IXLDataValidation FindDataValidation()
        {
            Worksheet.DataValidations.TryGet(new XLRangeAddress(Address, Address), out var dataValidation);
            return dataValidation;
        }

        public IXLDataValidation CreateDataValidation()
        {
            var validation = new XLDataValidation(AsRange());
            Worksheet.DataValidations.Add(validation);
            return validation;
        }

        [Obsolete("Use GetDataValidation() to access the existing rule, or CreateDataValidation() to create a new one.")]
        public IXLDataValidation SetDataValidation()
        {
            var validation = GetDataValidation();
            if (validation == null)
            {
                validation = new XLDataValidation(AsRange());
                Worksheet.DataValidations.Add(validation);
            }
            return validation;
        }

        public void Select()
        {
            AsRange().Select();
        }

        public IXLConditionalFormat AddConditionalFormat()
        {
            return AsRange().AddConditionalFormat();
        }

        public Boolean Active
        {
            get => Worksheet.ActiveCell == SheetPoint;
            set
            {
                if (value)
                    Worksheet.ActiveCell = SheetPoint;
                else if (Active)
                    Worksheet.ActiveCell = null;
            }
        }

        public IXLCell SetActive(Boolean value = true)
        {
            Active = value;
            return this;
        }

        public Boolean HasHyperlink => Worksheet.Hyperlinks.TryGet(SheetPoint, out _);

        /// <inheritdoc />
        public Boolean ShowPhonetic
        {
            get => _cellsCollection.MiscSlice[_rowNumber, _columnNumber].HasPhonetic;
            set
            {
                ref readonly var original = ref _cellsCollection.MiscSlice[_rowNumber, _columnNumber];
                if (original.HasPhonetic != value)
                {
                    var modified = original;
                    modified.HasPhonetic = value;
                    _cellsCollection.MiscSlice.Set(_rowNumber, _columnNumber, in modified);
                }
            }
        }

        #endregion IXLCell Members

        #region IXLStylized Members

        void IXLStylized.ModifyStyle(Func<XLStyleKey, XLStyleKey> modification)
        {
            //XLCell cannot have children so the base method may be optimized
            var styleKey = modification(StyleValue.Key);
            StyleValue = XLStyleValue.FromKey(ref styleKey);
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get { yield break; }
        }

        public override IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges { AsRange() };
                return retVal;
            }
        }

        #endregion IXLStylized Members

        /// <summary>
        /// Ensure the cell has style set directly on the cell, not inherited from column/row/worksheet styles.
        /// </summary>
        internal void PingStyle()
        {
            StyleValue = StyleValue;
        }

        public XLRange AsRange()
        {
            return Worksheet.Range(Address, Address);
        }

        #region Styles

        private XLStyleValue GetStyleForRead()
        {
            return StyleValue;
        }

        private void SetStyle(IXLStyle styleToUse)
        {
            Style = styleToUse;
        }

        public Boolean IsDefaultWorksheetStyle()
        {
            return StyleValue == Worksheet.StyleValue;
        }

        #endregion Styles

        public void DeleteComment()
        {
            Clear(XLClearOptions.Comments);
        }

        public void DeleteSparkline()
        {
            Clear(XLClearOptions.Sparklines);
        }

        private string GetFormat()
        {
            var style = GetStyleForRead();
            if (String.IsNullOrWhiteSpace(style.NumberFormat.Format))
            {
                var formatCodes = XLPredefinedFormat.FormatCodes;
                if (formatCodes.TryGetValue(style.NumberFormat.NumberFormatId, out string format))
                    return format;
                else
                    return string.Empty;
            }
            else
                return style.NumberFormat.Format;
        }

        public IXLCell CopyFrom(IXLRangeBase rangeObject)
        {
            if (rangeObject is null)
                throw new ArgumentNullException(nameof(rangeObject));

            var asRange = (XLRangeBase)rangeObject;
            var maxRows = asRange.RowCount();
            var maxColumns = asRange.ColumnCount();

            var lastRow = Math.Min(_rowNumber + maxRows - 1, XLHelper.MaxRowNumber);
            var lastColumn = Math.Min(_columnNumber + maxColumns - 1, XLHelper.MaxColumnNumber);

            var targetRange = Worksheet.Range(_rowNumber, _columnNumber, lastRow, lastColumn);

            if (!(asRange is XLRow || asRange is XLColumn))
            {
                targetRange.Clear();
            }

            var minRow = asRange.RangeAddress.FirstAddress.RowNumber;
            var minColumn = asRange.RangeAddress.FirstAddress.ColumnNumber;
            var cellsUsed = asRange.CellsUsed(XLCellsUsedOptions.All
                                              & ~XLCellsUsedOptions.ConditionalFormats
                                              & ~XLCellsUsedOptions.DataValidation
                                              & ~XLCellsUsedOptions.MergedRanges);
            foreach (var sourceCell in cellsUsed)
            {
                Worksheet.Cell(
                    _rowNumber + sourceCell.Address.RowNumber - minRow,
                    _columnNumber + sourceCell.Address.ColumnNumber - minColumn
                    ).CopyFromInternal(sourceCell as XLCell,
                    XLCellCopyOptions.All
                        & ~XLCellCopyOptions.ConditionalFormats
                        & ~XLCellCopyOptions.DataValidations); //Conditional formats and data validation are copied separately
            }

            var rangesToMerge = asRange.Worksheet.Internals.MergedRanges
                .Where(mr => asRange.Contains(mr))
                .Select(mr =>
                {
                    var firstRow = _rowNumber + (mr.RangeAddress.FirstAddress.RowNumber - asRange.RangeAddress.FirstAddress.RowNumber);
                    var firstColumn = _columnNumber + (mr.RangeAddress.FirstAddress.ColumnNumber - asRange.RangeAddress.FirstAddress.ColumnNumber);
                    return (IXLRange)Worksheet.Range
                    (
                        firstRow,
                        firstColumn,
                        firstRow + mr.RowCount() - 1,
                        firstColumn + mr.ColumnCount() - 1
                    );
                })
                .ToList();

            rangesToMerge.ForEach(r => r.Merge(false));

            var dataValidations = asRange.Worksheet.DataValidations
                .GetAllInRange(asRange.RangeAddress)
                .ToList();

            foreach (var dataValidation in dataValidations)
            {
                XLDataValidation newDataValidation = null;
                foreach (var dvRange in dataValidation.Ranges.Where(r => r.Intersects(asRange)))
                {
                    var dvTargetAddress = dvRange.RangeAddress.Relative(asRange.RangeAddress, targetRange.RangeAddress);
                    var dvTargetRange = Worksheet.Range(dvTargetAddress);
                    if (newDataValidation == null)
                    {
                        newDataValidation = dvTargetRange.CreateDataValidation() as XLDataValidation;
                        newDataValidation.CopyFrom(dataValidation);
                    }
                    else
                        newDataValidation.AddRange(dvTargetRange);
                }
            }

            CopyConditionalFormatsFrom(asRange);
            return this;
        }

        private void CopyConditionalFormatsFrom(XLCell otherCell)
        {
            var conditionalFormats = otherCell
                .Worksheet
                .ConditionalFormats
                .Where(c => c.Ranges.GetIntersectedRanges(otherCell).Any())
                .ToList();

            foreach (var cf in conditionalFormats)
            {
                if (otherCell.Worksheet == Worksheet)
                {
                    if (!cf.Ranges.GetIntersectedRanges(this).Any())
                    {
                        cf.Ranges.Add(this);
                    }
                }
                else
                {
                    CopyConditionalFormatsFrom(otherCell.AsRange());
                }
            }
        }

        private void CopyConditionalFormatsFrom(XLRangeBase fromRange)
        {
            var srcSheet = fromRange.Worksheet;
            int minRo = fromRange.RangeAddress.FirstAddress.RowNumber;
            int minCo = fromRange.RangeAddress.FirstAddress.ColumnNumber;
            if (srcSheet.ConditionalFormats.Any(r => r.Ranges.GetIntersectedRanges(fromRange.RangeAddress).Any()))
            {
                var fs = srcSheet.ConditionalFormats.SelectMany(cf => cf.Ranges.GetIntersectedRanges(fromRange.RangeAddress)).ToArray();
                if (fs.Any())
                {
                    minRo = fs.Max(r => r.RangeAddress.LastAddress.RowNumber);
                    minCo = fs.Max(r => r.RangeAddress.LastAddress.ColumnNumber);
                }
            }
            int rCnt = minRo - fromRange.RangeAddress.FirstAddress.RowNumber + 1;
            int cCnt = minCo - fromRange.RangeAddress.FirstAddress.ColumnNumber + 1;
            rCnt = Math.Min(rCnt, fromRange.RowCount());
            cCnt = Math.Min(cCnt, fromRange.ColumnCount());
            var toRange = Worksheet.Range(this, Worksheet.Cell(_rowNumber + rCnt - 1, _columnNumber + cCnt - 1));
            var formats = srcSheet.ConditionalFormats.Where(f => f.Ranges.GetIntersectedRanges(fromRange.RangeAddress).Any());

            foreach (var cf in formats.ToList())
            {
                var fmtRanges = cf.Ranges
                    .GetIntersectedRanges(fromRange.RangeAddress)
                    .Select(r => r.RangeAddress.Intersection(fromRange.RangeAddress).Relative(fromRange.RangeAddress, toRange.RangeAddress).AsRange() as XLRange)
                    .ToList();

                var c = new XLConditionalFormat(fmtRanges, true);
                c.CopyFrom(cf);
                c.AdjustFormulas((XLCell)cf.Ranges.First().FirstCell(), fmtRanges.First().FirstCell());

                Worksheet.ConditionalFormats.Add(c);
            }
        }

        private bool SetDataTable(object o)
        {
            if (o is DataTable dataTable)
                return InsertData(dataTable) != null;
            else
                return false;
        }

        private bool SetEnumerable(object collectionObject)
        {
            // IXLRichText implements IEnumerable, but we don't want to handle this here.
            if (collectionObject is IXLRichText) return false;

            var asEnumerable = collectionObject as IEnumerable;
            return InsertData(asEnumerable) != null;
        }

        private void ClearMerged()
        {
            List<IXLRange> mergeToDelete = Worksheet.Internals.MergedRanges.GetIntersectedRanges(Address).ToList();

            mergeToDelete.ForEach(m => Worksheet.Internals.MergedRanges.Remove(m));
        }

        internal string GetFormulaR1C1(string value)
        {
            return XLCellFormula.GetFormula(value, FormulaConversionType.A1ToR1C1, new XLSheetPoint(_rowNumber, _columnNumber));
        }

        internal string GetFormulaA1(string value)
        {
            return XLCellFormula.GetFormula(value, FormulaConversionType.R1C1ToA1, new XLSheetPoint(_rowNumber, _columnNumber));
        }

        internal void CopyValuesFrom(XLCell source)
        {
            // Rich text is basically a super set of a value. Setting a value would override rich text and vice versa.
            var sourceRichText = source.SliceRichText;
            if (sourceRichText is null)
                SliceCellValue = source.SliceCellValue;
            else
                SliceRichText = sourceRichText;

            FormulaR1C1 = source.FormulaR1C1;
            SliceComment = source.SliceComment == null ? null : new XLComment(this, source.SliceComment, source.Style.Font, source.SliceComment.Style);

            if (Worksheet.Hyperlinks.TryGet(source.SheetPoint, out var sourceHyperlink))
            {
                SetCellHyperlink(new XLHyperlink(sourceHyperlink));
            }
        }

        private IXLCell GetTargetCell(String target, XLWorksheet defaultWorksheet)
        {
            var pair = target.Split('!');
            if (pair.Length == 1)
                return defaultWorksheet.Cell(target);

            var wsName = pair[0];
            if (wsName.StartsWith("'"))
                wsName = wsName.Substring(1, wsName.Length - 2);
            return defaultWorksheet.Workbook.Worksheet(wsName).Cell(pair[1]);
        }

        internal IXLCell CopyFromInternal(XLCell otherCell, XLCellCopyOptions options)
        {
            if (options.HasFlag(XLCellCopyOptions.Values))
                CopyValuesFrom(otherCell);

            if (options.HasFlag(XLCellCopyOptions.Styles))
                InnerStyle = otherCell.InnerStyle;

            if (options.HasFlag(XLCellCopyOptions.Sparklines))
                CopySparklineFrom(otherCell);

            if (options.HasFlag(XLCellCopyOptions.ConditionalFormats))
                CopyConditionalFormatsFrom(otherCell);

            if (options.HasFlag(XLCellCopyOptions.DataValidations))
                CopyDataValidationFrom(otherCell);

            return this;
        }

        private void CopySparklineFrom(XLCell otherCell)
        {
            if (!otherCell.HasSparkline) return;

            var sourceDataAddress = otherCell.Sparkline.SourceData.RangeAddress.ToString();
            var shiftedRangeAddress = GetFormulaA1(otherCell.GetFormulaR1C1(sourceDataAddress));
            var sourceDataWorksheet = otherCell.Worksheet == otherCell.Sparkline.SourceData.Worksheet
                ? Worksheet
                : (XLWorksheet)otherCell.Sparkline.SourceData.Worksheet;
            var sourceData = sourceDataWorksheet.Range(shiftedRangeAddress);

            IXLSparklineGroup group;
            if (otherCell.Worksheet == Worksheet)
            {
                group = otherCell.Sparkline.SparklineGroup;
            }
            else
            {
                group = Worksheet.SparklineGroups.Add(new XLSparklineGroup(Worksheet, otherCell.Sparkline.SparklineGroup));
                if (otherCell.Sparkline.SparklineGroup.DateRange != null)
                {
                    var dateRangeWorksheet =
                        otherCell.Worksheet == otherCell.Sparkline.SparklineGroup.DateRange.Worksheet
                            ? Worksheet
                            : otherCell.Sparkline.SparklineGroup.DateRange.Worksheet;
                    var dateRangeAddress = otherCell.Sparkline.SparklineGroup.DateRange.RangeAddress.ToString();
                    var shiftedDateRangeAddress = GetFormulaA1(otherCell.GetFormulaR1C1(dateRangeAddress));
                    group.SetDateRange(dateRangeWorksheet.Range(shiftedDateRangeAddress));
                }
            }

            group.Add(this, sourceData);
        }

        public IXLCell CopyFrom(IXLCell otherCell, XLCellCopyOptions options)
        {
            var source = otherCell as XLCell; // To expose GetFormulaR1C1, etc

            CopyFromInternal(source, options);
            return this;
        }

        private void CopyDataValidationFrom(XLCell otherCell)
        {
            if (otherCell.HasDataValidation)
                CopyDataValidation(otherCell, otherCell.GetDataValidation());
            else if (HasDataValidation)
            {
                Worksheet.DataValidations.Delete(AsRange());
            }
        }

        internal void CopyDataValidation(XLCell otherCell, IXLDataValidation otherDv)
        {
            var thisDv = GetDataValidation() as XLDataValidation;
            thisDv.CopyFrom(otherDv);
            thisDv.Value = GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.Value));
            thisDv.MinValue = GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.MinValue));
            thisDv.MaxValue = GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.MaxValue));
        }

        internal void ShiftFormulaRows(XLRange shiftedRange, int rowsShifted)
        {
            FormulaA1 = ShiftFormulaRows(FormulaA1, Worksheet, shiftedRange, rowsShifted);
        }

        internal static String ShiftFormulaRows(String formulaA1, XLWorksheet worksheetInAction, XLRange shiftedRange,
                                                int rowsShifted)
        {
            if (String.IsNullOrWhiteSpace(formulaA1)) return String.Empty;

            var value = formulaA1;

            var regex = A1SimpleRegex;

            var sb = new StringBuilder();
            var lastIndex = 0;

            var shiftedWsName = shiftedRange.Worksheet.Name;
            foreach (var match in regex.Matches(value).Cast<Match>())
            {
                var matchString = match.Value;
                var matchIndex = match.Index;
                if (value.Substring(0, matchIndex).CharCount('"') % 2 == 0)
                {
                    // Check that the match is not between quotes
                    sb.Append(value.Substring(lastIndex, matchIndex - lastIndex));
                    string sheetName;
                    var useSheetName = false;
                    if (matchString.Contains('!'))
                    {
                        sheetName = matchString.Substring(0, matchString.IndexOf('!'));
                        if (sheetName[0] == '\'')
                            sheetName = sheetName.Substring(1, sheetName.Length - 2);
                        useSheetName = true;
                    }
                    else
                        sheetName = worksheetInAction.Name;

                    if (String.Compare(sheetName, shiftedWsName, true) == 0)
                    {
                        var rangeAddress = matchString.Substring(matchString.IndexOf('!') + 1);
                        if (!A1ColumnRegex.IsMatch(rangeAddress))
                        {
                            var matchRange = worksheetInAction.Workbook.Worksheet(sheetName).Range(rangeAddress);
                            if (shiftedRange.RangeAddress.FirstAddress.RowNumber <= matchRange.RangeAddress.LastAddress.RowNumber
                                && shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= matchRange.RangeAddress.FirstAddress.ColumnNumber
                                && shiftedRange.RangeAddress.LastAddress.ColumnNumber >= matchRange.RangeAddress.LastAddress.ColumnNumber)
                            {
                                if (useSheetName)
                                {
                                    sb.Append(sheetName.EscapeSheetName());
                                    sb.Append('!');
                                }

                                if (A1RowRegex.IsMatch(rangeAddress))
                                {
                                    var rows = rangeAddress.Split(':');
                                    var row1String = rows[0];
                                    var row2String = rows[1];
                                    string row1;
                                    if (row1String[0] == '$')
                                    {
                                        row1 = "$" +
                                                (XLHelper.TrimRowNumber(Int32.Parse(row1String.Substring(1)) + rowsShifted)).ToInvariantString();
                                    }
                                    else
                                        row1 = (XLHelper.TrimRowNumber(Int32.Parse(row1String) + rowsShifted)).ToInvariantString();

                                    string row2;
                                    if (row2String[0] == '$')
                                    {
                                        row2 = "$" +
                                                (XLHelper.TrimRowNumber(Int32.Parse(row2String.Substring(1)) + rowsShifted)).ToInvariantString();
                                    }
                                    else
                                        row2 = (XLHelper.TrimRowNumber(Int32.Parse(row2String) + rowsShifted)).ToInvariantString();

                                    sb.Append(row1);
                                    sb.Append(':');
                                    sb.Append(row2);
                                }
                                else if (shiftedRange.RangeAddress.FirstAddress.RowNumber <=
                                            matchRange.RangeAddress.FirstAddress.RowNumber)
                                {
                                    if (rangeAddress.Contains(':'))
                                    {
                                        sb.Append(
                                            new XLAddress(
                                                worksheetInAction,
                                                XLHelper.TrimRowNumber(matchRange.RangeAddress.FirstAddress.RowNumber + rowsShifted),
                                                matchRange.RangeAddress.FirstAddress.ColumnLetter,
                                                matchRange.RangeAddress.FirstAddress.FixedRow,
                                                matchRange.RangeAddress.FirstAddress.FixedColumn));
                                        sb.Append(':');
                                        sb.Append(
                                            new XLAddress(
                                                worksheetInAction,
                                                XLHelper.TrimRowNumber(matchRange.RangeAddress.LastAddress.RowNumber + rowsShifted),
                                                matchRange.RangeAddress.LastAddress.ColumnLetter,
                                                matchRange.RangeAddress.LastAddress.FixedRow,
                                                matchRange.RangeAddress.LastAddress.FixedColumn));
                                    }
                                    else
                                    {
                                        sb.Append(
                                            new XLAddress(
                                                worksheetInAction,
                                                XLHelper.TrimRowNumber(matchRange.RangeAddress.FirstAddress.RowNumber + rowsShifted),
                                                matchRange.RangeAddress.FirstAddress.ColumnLetter,
                                                matchRange.RangeAddress.FirstAddress.FixedRow,
                                                matchRange.RangeAddress.FirstAddress.FixedColumn));
                                    }
                                }
                                else
                                {
                                    sb.Append(matchRange.RangeAddress.FirstAddress);
                                    sb.Append(':');
                                    sb.Append(
                                        new XLAddress(
                                            worksheetInAction,
                                            XLHelper.TrimRowNumber(matchRange.RangeAddress.LastAddress.RowNumber + rowsShifted),
                                            matchRange.RangeAddress.LastAddress.ColumnLetter,
                                            matchRange.RangeAddress.LastAddress.FixedRow,
                                            matchRange.RangeAddress.LastAddress.FixedColumn));
                                }
                            }
                            else
                                sb.Append(matchString);
                        }
                        else
                            sb.Append(matchString);
                    }
                    else
                        sb.Append(matchString);
                }
                else
                    sb.Append(value.Substring(lastIndex, matchIndex - lastIndex + matchString.Length));

                lastIndex = matchIndex + matchString.Length;
            }

            if (lastIndex < value.Length)
                sb.Append(value.Substring(lastIndex));

            return sb.ToString();
        }

        internal void ShiftFormulaColumns(XLRange shiftedRange, int columnsShifted)
        {
            FormulaA1 = ShiftFormulaColumns(FormulaA1, Worksheet, shiftedRange, columnsShifted);
        }

        internal static String ShiftFormulaColumns(String formulaA1, XLWorksheet worksheetInAction, XLRange shiftedRange,
                                                   int columnsShifted)
        {
            if (String.IsNullOrWhiteSpace(formulaA1)) return String.Empty;

            var value = formulaA1;

            var regex = A1SimpleRegex;

            var sb = new StringBuilder();
            var lastIndex = 0;

            foreach (var match in regex.Matches(value).Cast<Match>())
            {
                var matchString = match.Value;
                var matchIndex = match.Index;
                if (value.Substring(0, matchIndex).CharCount('"') % 2 == 0)
                {
                    // Check that the match is not between quotes
                    sb.Append(value.Substring(lastIndex, matchIndex - lastIndex));
                    string sheetName;
                    var useSheetName = false;
                    if (matchString.Contains('!'))
                    {
                        sheetName = matchString.Substring(0, matchString.IndexOf('!'));
                        if (sheetName[0] == '\'')
                            sheetName = sheetName.Substring(1, sheetName.Length - 2);
                        useSheetName = true;
                    }
                    else
                        sheetName = worksheetInAction.Name;

                    if (String.Compare(sheetName, shiftedRange.Worksheet.Name, true) == 0)
                    {
                        var rangeAddress = matchString.Substring(matchString.IndexOf('!') + 1);
                        if (!A1RowRegex.IsMatch(rangeAddress))
                        {
                            var matchRange = worksheetInAction.Workbook.Worksheet(sheetName).Range(rangeAddress);

                            if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                                matchRange.RangeAddress.LastAddress.ColumnNumber
                                &&
                                shiftedRange.RangeAddress.FirstAddress.RowNumber <=
                                matchRange.RangeAddress.FirstAddress.RowNumber
                                &&
                                shiftedRange.RangeAddress.LastAddress.RowNumber >=
                                matchRange.RangeAddress.LastAddress.RowNumber)
                            {
                                if (useSheetName)
                                {
                                    sb.Append(sheetName.EscapeSheetName());
                                    sb.Append('!');
                                }

                                if (A1ColumnRegex.IsMatch(rangeAddress))
                                {
                                    var columns = rangeAddress.Split(':');
                                    var column1String = columns[0];
                                    var column2String = columns[1];
                                    string column1;
                                    if (column1String[0] == '$')
                                    {
                                        column1 = "$" +
                                                    XLHelper.GetColumnLetterFromNumber(
                                                        XLHelper.GetColumnNumberFromLetter(
                                                            column1String.Substring(1)) + columnsShifted, true);
                                    }
                                    else
                                    {
                                        column1 =
                                            XLHelper.GetColumnLetterFromNumber(
                                                XLHelper.GetColumnNumberFromLetter(column1String) +
                                                columnsShifted, true);
                                    }

                                    string column2;
                                    if (column2String[0] == '$')
                                    {
                                        column2 = "$" +
                                                    XLHelper.GetColumnLetterFromNumber(
                                                        XLHelper.GetColumnNumberFromLetter(
                                                            column2String.Substring(1)) + columnsShifted, true);
                                    }
                                    else
                                    {
                                        column2 =
                                            XLHelper.GetColumnLetterFromNumber(
                                                XLHelper.GetColumnNumberFromLetter(column2String) +
                                                columnsShifted, true);
                                    }

                                    sb.Append(column1);
                                    sb.Append(':');
                                    sb.Append(column2);
                                }
                                else if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                                            matchRange.RangeAddress.FirstAddress.ColumnNumber)
                                {
                                    if (rangeAddress.Contains(':'))
                                    {
                                        sb.Append(
                                            new XLAddress(
                                                worksheetInAction,
                                                matchRange.RangeAddress.FirstAddress.RowNumber,
                                                XLHelper.TrimColumnNumber(matchRange.RangeAddress.FirstAddress.ColumnNumber + columnsShifted),
                                                matchRange.RangeAddress.FirstAddress.FixedRow,
                                                matchRange.RangeAddress.FirstAddress.FixedColumn));
                                        sb.Append(':');
                                        sb.Append(
                                            new XLAddress(
                                                worksheetInAction,
                                                matchRange.RangeAddress.LastAddress.RowNumber,
                                                XLHelper.TrimColumnNumber(matchRange.RangeAddress.LastAddress.ColumnNumber + columnsShifted),
                                                matchRange.RangeAddress.LastAddress.FixedRow,
                                                matchRange.RangeAddress.LastAddress.FixedColumn));
                                    }
                                    else
                                    {
                                        sb.Append(
                                            new XLAddress(
                                                worksheetInAction,
                                                matchRange.RangeAddress.FirstAddress.RowNumber,
                                                XLHelper.TrimColumnNumber(matchRange.RangeAddress.FirstAddress.ColumnNumber + columnsShifted),
                                                matchRange.RangeAddress.FirstAddress.FixedRow,
                                                matchRange.RangeAddress.FirstAddress.FixedColumn));
                                    }
                                }
                                else
                                {
                                    sb.Append(matchRange.RangeAddress.FirstAddress);
                                    sb.Append(':');
                                    sb.Append(
                                        new XLAddress(
                                            worksheetInAction,
                                            matchRange.RangeAddress.LastAddress.RowNumber,
                                            XLHelper.TrimColumnNumber(matchRange.RangeAddress.LastAddress.ColumnNumber + columnsShifted),
                                            matchRange.RangeAddress.LastAddress.FixedRow,
                                            matchRange.RangeAddress.LastAddress.FixedColumn));
                                }
                            }
                            else
                                sb.Append(matchString);
                        }
                        else
                            sb.Append(matchString);
                    }
                    else
                        sb.Append(matchString);
                }
                else
                    sb.Append(value.Substring(lastIndex, matchIndex - lastIndex + matchString.Length));
                lastIndex = matchIndex + matchString.Length;
            }

            if (lastIndex < value.Length)
                sb.Append(value.Substring(lastIndex));

            return sb.ToString();
        }

        private XLCell CellShift(Int32 rowsToShift, Int32 columnsToShift)
        {
            return Worksheet.Cell(_rowNumber + rowsToShift, _columnNumber + columnsToShift);
        }

        #region XLCell Above

        IXLCell IXLCell.CellAbove()
        {
            return CellAbove();
        }

        IXLCell IXLCell.CellAbove(Int32 step)
        {
            return CellAbove(step);
        }

        public XLCell CellAbove()
        {
            return CellAbove(1);
        }

        public XLCell CellAbove(Int32 step)
        {
            return CellShift(step * -1, 0);
        }

        #endregion XLCell Above

        #region XLCell Below

        IXLCell IXLCell.CellBelow()
        {
            return CellBelow();
        }

        IXLCell IXLCell.CellBelow(Int32 step)
        {
            return CellBelow(step);
        }

        public XLCell CellBelow()
        {
            return CellBelow(1);
        }

        public XLCell CellBelow(Int32 step)
        {
            return CellShift(step, 0);
        }

        #endregion XLCell Below

        #region XLCell Left

        IXLCell IXLCell.CellLeft()
        {
            return CellLeft();
        }

        IXLCell IXLCell.CellLeft(Int32 step)
        {
            return CellLeft(step);
        }

        public XLCell CellLeft()
        {
            return CellLeft(1);
        }

        public XLCell CellLeft(Int32 step)
        {
            return CellShift(0, step * -1);
        }

        #endregion XLCell Left

        #region XLCell Right

        IXLCell IXLCell.CellRight()
        {
            return CellRight();
        }

        IXLCell IXLCell.CellRight(Int32 step)
        {
            return CellRight(step);
        }

        public XLCell CellRight()
        {
            return CellRight(1);
        }

        public XLCell CellRight(Int32 step)
        {
            return CellShift(0, step);
        }

        #endregion XLCell Right

        public Boolean HasFormula => Formula is not null;

        public Boolean HasArrayFormula => Formula?.Type == FormulaType.Array;

        public IXLRangeAddress FormulaReference
        {
            get
            {
                if (Formula is null)
                    return null;

                var range = Formula.Range;
                if (range == default)
                    return null;

                return XLRangeAddress.FromSheetRange(Worksheet, range);
            }
            set
            {
                if (Formula is null)
                    throw new ArgumentException("Cell doesn't contain a formula.");

                if (value is null)
                {
                    Formula.Range = default;
                    return;
                }

                if (value.Worksheet is not null && Worksheet != value.Worksheet)
                    throw new ArgumentException("The reference worksheet must be same as worksheet of the cell or null.");

                Formula.Range = XLSheetRange.FromRangeAddress(value);
            }
        }

        public IXLRange CurrentRegion => Worksheet.Range(FindCurrentRegion());

        private IXLRangeAddress FindCurrentRegion()
        {
            var sheet = Worksheet;

            var minRow = _rowNumber;
            var minCol = _columnNumber;
            var maxRow = _rowNumber;
            var maxCol = _columnNumber;

            bool hasRegionExpanded;

            do
            {
                hasRegionExpanded = false;

                var borderMinRow = Math.Max(minRow - 1, XLHelper.MinRowNumber);
                var borderMaxRow = Math.Min(maxRow + 1, XLHelper.MaxRowNumber);
                var borderMinColumn = Math.Max(minCol - 1, XLHelper.MinColumnNumber);
                var borderMaxColumn = Math.Min(maxCol + 1, XLHelper.MaxColumnNumber);

                if (minCol > XLHelper.MinColumnNumber &&
                    !IsVerticalBorderBlank(sheet, borderMinColumn, borderMinRow, borderMaxRow))
                {
                    hasRegionExpanded = true;
                    minCol = borderMinColumn;
                }

                if (maxCol < XLHelper.MaxColumnNumber &&
                    !IsVerticalBorderBlank(sheet, borderMaxColumn, borderMinRow, borderMaxRow))
                {
                    hasRegionExpanded = true;
                    maxCol = borderMaxColumn;
                }

                if (minRow > XLHelper.MinRowNumber &&
                    !IsHorizontalBorderBlank(sheet, borderMinRow, borderMinColumn, borderMaxColumn))
                {
                    hasRegionExpanded = true;
                    minRow = borderMinRow;
                }

                if (maxRow < XLHelper.MaxRowNumber &&
                    !IsHorizontalBorderBlank(sheet, borderMaxRow, borderMinColumn, borderMaxColumn))
                {
                    hasRegionExpanded = true;
                    maxRow = borderMaxRow;
                }
            } while (hasRegionExpanded);

            return new XLRangeAddress(
                new XLAddress(sheet, minRow, minCol, false, false),
                new XLAddress(sheet, maxRow, maxCol, false, false));

            static bool IsVerticalBorderBlank(XLWorksheet sheet, int borderColumn, int borderMinRow, int borderMaxRow)
            {
                for (var row = borderMinRow; row <= borderMaxRow; row++)
                {
                    var verticalBorderCell = sheet.Cell(row, borderColumn);
                    if (!verticalBorderCell.IsEmpty(XLCellsUsedOptions.AllContents))
                    {
                        return false;
                    }
                }

                return true;
            }

            static bool IsHorizontalBorderBlank(XLWorksheet sheet, int borderRow, int borderMinColumn, int borderMaxColumn)
            {
                for (var col = borderMinColumn; col <= borderMaxColumn; col++)
                {
                    var horizontalBorderCell = sheet.Cell(borderRow, col);
                    if (!horizontalBorderCell.IsEmpty(XLCellsUsedOptions.AllContents))
                    {
                        return false;
                    }
                }

                return true;
            }
        }

        internal bool IsInferiorMergedCell()
        {
            return this.IsMerged() && !this.Address.Equals(this.MergedRange().RangeAddress.FirstAddress);
        }

        internal bool IsSuperiorMergedCell()
        {
            return this.IsMerged() && this.Address.Equals(this.MergedRange().RangeAddress.FirstAddress);
        }

        /// <summary>
        /// Get glyph bounding boxes for each grapheme in the text. Box size is determined according to
        /// the font of a grapheme. New lines are represented as default (all dimensions zero) box.
        /// A line without any text (i.e. contains only new line) should be represented by a box
        /// with zero advance width, but with a line height of corresponding font.
        /// </summary>
        /// <param name="engine">Engine used to determine box size.</param>
        /// <param name="dpi">DPI used to determine size of glyphs.</param>
        /// <param name="output">List where items are added.</param>
        internal void GetGlyphBoxes(IXLGraphicEngine engine, Dpi dpi, List<GlyphBox> output)
        {
            var richText = SliceRichText;
            if (richText is not null)
            {
                foreach (var richTextRun in richText.Runs)
                {
                    var text = richText.GetRunText(richTextRun);
                    var font = new XLFont(richTextRun.Font.Key);
                    AddGlyphs(text, font, engine, dpi, output);
                }
            }
            else
            {
                var text = GetFormattedString();
                AddGlyphs(text, Style.Font, engine, dpi, output);
            }

            static void AddGlyphs(string text, IXLFontBase font, IXLGraphicEngine engine, Dpi dpi, List<GlyphBox> output)
            {
                Span<int> zeroWidthJoiner = stackalloc int[1] { 0x200D };
                var prevWasNewLine = false;
                var graphemeStarts = StringInfo.ParseCombiningCharacters(text);
                var textSpan = text.AsSpan();

                // If we have more than 1 code unit per grapheme, the code units can
                // be distributed through multiple grapheme. In the worst case, all extra
                // code units are in exactly one grapheme -> allocate buffer of that size.
                Span<int> codePointsBuffer = stackalloc int[1 + text.Length - graphemeStarts.Length];
                for (var i = 0; i < graphemeStarts.Length; ++i)
                {
                    var startIdx = graphemeStarts[i];
                    var slice = textSpan.Slice(startIdx);
                    if (slice.TrySliceNewLine(out var eolLen))
                    {
                        i += eolLen - 1;
                        if (prevWasNewLine)
                        {
                            // If there are consecutive new lines, we need height of new the lines between them
                            var box = engine.GetGlyphBox(zeroWidthJoiner, font, dpi);
                            output.Add(box);
                        }

                        output.Add(GlyphBox.LineBreak);
                        prevWasNewLine = true;
                    }
                    else
                    {
                        var codeUnits = i + 1 < graphemeStarts.Length
                            ? textSpan.Slice(startIdx, graphemeStarts[i + 1] - startIdx)
                            : textSpan.Slice(startIdx);
                        var count = codeUnits.ToCodePoints(codePointsBuffer);
                        ReadOnlySpan<int> grapheme = codePointsBuffer.Slice(0, count);
                        var box = engine.GetGlyphBox(grapheme, font, dpi);
                        output.Add(box);
                        prevWasNewLine = false;
                    }
                }
            }
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (SheetPoint.GetHashCode() * 397) ^ Worksheet.GetHashCode();
            }
        }

        public override bool Equals(object obj)
        {
            return obj is XLCell cell && cell.Worksheet == Worksheet && cell.SheetPoint == SheetPoint;
        }
    }
}
