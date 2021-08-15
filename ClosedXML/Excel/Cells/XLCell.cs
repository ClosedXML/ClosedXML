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

namespace ClosedXML.Excel
{
    [DebuggerDisplay("{Address}")]
    internal class XLCell : XLStylizedBase, IXLCell, IXLStylized
    {
        public static readonly DateTime BaseDate = new DateTime(1899, 12, 30);

        private static readonly Regex A1Regex = new Regex(
            @"(?<=\W)(\$?[a-zA-Z]{1,3}\$?\d{1,7})(?=\W)" // A1
            + @"|(?<=\W)(\$?\d{1,7}:\$?\d{1,7})(?=\W)" // 1:1
            + @"|(?<=\W)(\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})(?=\W)", RegexOptions.Compiled); // A:A

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

        private static readonly Regex R1C1Regex = new Regex(
            @"(?<=\W)([Rr](?:\[-?\d{0,7}\]|\d{0,7})?[Cc](?:\[-?\d{0,7}\]|\d{0,7})?)(?=\W)" // R1C1
            + @"|(?<=\W)([Rr]\[?-?\d{0,7}\]?:[Rr]\[?-?\d{0,7}\]?)(?=\W)" // R:R
            + @"|(?<=\W)([Cc]\[?-?\d{0,5}\]?:[Cc]\[?-?\d{0,5}\]?)(?=\W)", RegexOptions.Compiled); // C:C

        private static readonly Regex utfPattern = new Regex(@"(?<!_x005F)_x(?!005F)([0-9A-F]{4})_", RegexOptions.Compiled);

        #region Fields

        private string _cellValue = String.Empty;

        private XLComment _comment;
        private XLDataType _dataType;
        private XLHyperlink _hyperlink;
        private XLRichText _richText;

        public bool SettingHyperlink;
        internal int SharedStringId { get; set; }
        private string _formulaA1;
        private string _formulaR1C1;

        #endregion Fields

        #region Constructor

        internal XLCell(XLWorksheet worksheet, XLAddress address, XLStyleValue styleValue)
            : base(styleValue)
        {
            Address = address;
            ShareString = true;
            Worksheet = worksheet;
        }

        public XLCell(XLWorksheet worksheet, XLAddress address, IXLStyle style)
            : this(worksheet, address, (style as XLStyle).Value)
        {
        }

        public XLCell(XLWorksheet worksheet, XLAddress address)
            : this(worksheet, address, XLStyle.Default.Value)
        {
        }

        #endregion Constructor

        public XLWorksheet Worksheet { get; private set; }

        private int _rowNumber;
        private int _columnNumber;
        private bool _fixedRow;
        private bool _fixedCol;

        public XLAddress Address
        {
            get
            {
                return new XLAddress(Worksheet, _rowNumber, _columnNumber, _fixedRow, _fixedCol);
            }
            internal set
            {
                if (value == null)
                    return;
                _rowNumber = value.RowNumber;
                _columnNumber = value.ColumnNumber;
                _fixedRow = value.FixedRow;
                _fixedCol = value.FixedColumn;
            }
        }

        public string InnerText
        {
            get
            {
                if (HasRichText)
                    return _richText.ToString();

                return string.Empty == _cellValue ? FormulaA1 : _cellValue;
            }
        }

        internal XLComment GetComment()
        {
            return _comment ?? CreateComment();
        }

        internal XLComment CreateComment(int? shapeId = null)
        {
            _comment = new XLComment(this, shapeId: shapeId);
            return _comment;
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

        public IXLCell SetValue<T>(T value)
        {
            if (value is IEnumerable ie && !(value is string))
            {
                this.InsertData(ie);
                return this;
            }
            else
                return SetValue(value, setTableHeader: true, checkMergedRanges: true);
        }

        internal IXLCell SetValue<T>(T value, bool setTableHeader, bool checkMergedRanges)
        {
            if (checkMergedRanges && IsInferiorMergedCell())
                return this;

            if (value == null)
                return this.Clear(XLClearOptions.Contents);

            _richText = null;
            _formulaA1 = String.Empty;
            _formulaR1C1 = null;
            cachedValue = null;

            if (setTableHeader)
            {
                if (SetTableHeaderValue(value)) return this;
                if (SetTableTotalsRowLabel(value)) return this;
            }

            var style = GetStyleForRead();

            // For SetValue<T> we set the cell value directly to the parameter
            // as opposed to the other SetValue(object value) where we parse the string and try to deduce the value

            // If parsing was unsuccessful, we throw an ArgumentException
            // because we are using SetValue<T> (typed).
            // Only in SetValue(object value) to we try to fall back to a value of a different type
            if (!TrySetInternallySupportedType(ConvertOtherSupportedTypes(value), style, out var parsedValue))
                throw new ArgumentException($"Unable to set cell value to {value.ObjectToInvariantString()}");

            SetInternalCellValueString(parsedValue, validate: true, parseToCachedValue: false);

            return this;
        }

        /// <summary>
        /// This method converts a wider set up supported types to internal types
        /// </summary>
        /// <param name="value"></param>
        internal static object ConvertOtherSupportedTypes(object value)
        {
            return value switch
            {
                DBNull dbnull when dbnull.Equals(DBNull.Value) => string.Empty,
                Guid _ or char _ or Enum _ => value.ObjectToInvariantString(),
                DateTimeOffset dto => dto.DateTime,
                _ => value,
            };
        }

        /// <summary>
        /// This method accepts only values of the supported internal types: String, DateTime, TimeSpan, Boolean and Numbers.
        /// Any other types should be converted with <see cref="ConvertOtherSupportedTypes(object)"/> first.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="value"></param>
        /// <param name="style"></param>
        /// <param name="internalValue"></param>
        private bool TrySetInternallySupportedType<T>(T value, XLStyleValue style, out string internalValue)
        {
            internalValue = null;

            switch (value)
            {
                case string s:
                    internalValue = s;
                    _dataType = XLDataType.Text;
                    if (s.Contains(Environment.NewLine) && !style.Alignment.WrapText)
                        Style.Alignment.WrapText = true;

                    return true;

                case DateTime d when d >= BaseDate:
                    SetDateTimeFormat(style, d.Date == d);
                    internalValue = d.ToOADate().ToInvariantString();
                    return true;

                case TimeSpan ts:
                    SetTimeSpanFormat(style);
                    internalValue = ts.TotalDays.ToInvariantString();
                    return true;

                case bool b:
                    _dataType = XLDataType.Boolean;
                    internalValue = b ? "1" : "0";
                    return true;

                default:
                    if (value.IsNumber())
                    {
                        if (
                            (value is double d1 && (double.IsNaN(d1) || double.IsInfinity(d1)))
                            || (value is float f && (float.IsNaN(f) || float.IsInfinity(f)))
                           )
                        {
                            _dataType = XLDataType.Text;
                            internalValue = value.ToString();
                            return true;
                        }
                        else
                        {
                            _dataType = XLDataType.Number;
                            internalValue = value.ObjectToInvariantString();
                            return true;
                        }
                    }
                    return false;
            }
        }

        private string DeduceCellValueByParsing(string value, XLStyleValue style)
        {
            if (String.IsNullOrEmpty(value))
            {
                _dataType = XLDataType.Text;
            }
            else if (value[0] == '\'')
            {
                // If a user sets a cell value to a value starting with a single quote
                // ensure the data type is text
                // and that it will be prefixed with a quote in Excel too

                value = value.Substring(1, value.Length - 1);

                _dataType = XLDataType.Text;
                if (value.Contains(Environment.NewLine) && !style.Alignment.WrapText)
                    Style.Alignment.WrapText = true;

                this.Style.SetIncludeQuotePrefix();
            }
            else if (!string.Equals(value.Trim(), "NaN", StringComparison.OrdinalIgnoreCase) &&
                     Double.TryParse(value, XLHelper.NumberStyle, XLHelper.ParseCulture, out Double _))
                _dataType = XLDataType.Number;
            else if (TimeSpan.TryParse(value, out TimeSpan ts))
            {
                value = ts.ToInvariantString();
                SetTimeSpanFormat(style);
            }
            else if (DateTime.TryParse(value, out DateTime dt) && dt >= BaseDate)
            {
                value = dt.ToOADate().ToInvariantString();
                SetDateTimeFormat(style, dt.Date == dt);
            }
            else if (Boolean.TryParse(value, out Boolean b))
            {
                value = b ? "1" : "0";
                _dataType = XLDataType.Boolean;
            }
            else
            {
                _dataType = XLDataType.Text;
                if (value.Contains(Environment.NewLine) && !style.Alignment.WrapText)
                    Style.Alignment.WrapText = true;
            }

            return value;
        }

        public T GetValue<T>()
        {
            if (TryGetValue(out T retVal))
                return retVal;

            throw new FormatException($"Cannot convert {this.Address.ToStringRelative(true)}'s value to " + typeof(T));
        }

        public string GetString()
        {
            return GetValue<string>();
        }

        public double GetDouble()
        {
            return GetValue<double>();
        }

        public bool GetBoolean()
        {
            return GetValue<bool>();
        }

        public DateTime GetDateTime()
        {
            return GetValue<DateTime>();
        }

        public TimeSpan GetTimeSpan()
        {
            return GetValue<TimeSpan>();
        }

        public string GetFormattedString()
        {
            var format = GetFormat();
            try
            {
                return Value.ToExcelFormat(format);
            }
            catch { }

            try
            {
                return CachedValue.ToExcelFormat(format);
            }
            catch { }

            return _cellValue;
        }

        /// <summary>
        /// Flag showing that the cell is in formula evaluation state.
        /// </summary>
        internal bool IsEvaluating { get; private set; }

        /// <summary>
        /// Calculate a value of the specified formula.
        /// </summary>
        /// <param name="fA1">Cell formula to evaluate.</param>
        /// <returns>Null if formula is empty or null, calculated value otherwise.</returns>
        private object RecalculateFormula(string fA1)
        {
            if (string.IsNullOrEmpty(fA1))
                return null;

            if (IsEvaluating)
                throw new InvalidOperationException("Circular Reference");

            if (fA1[0] == '{')
                fA1 = fA1.Substring(1, fA1.Length - 2);

            string sName;
            string cAddress;
            if (fA1.Contains('!'))
            {
                sName = fA1.Substring(0, fA1.IndexOf('!'));
                if (sName[0] == '\'')
                    sName = sName.Substring(1, sName.Length - 2);

                cAddress = fA1.Substring(fA1.IndexOf('!') + 1);
            }
            else
            {
                sName = Worksheet.Name;
                cAddress = fA1;
            }

            if (Worksheet.Workbook.Worksheets.Contains(sName)
                && XLHelper.IsValidA1Address(cAddress))
            {
                try
                {
                    IsEvaluating = true;
                    var referenceCell = Worksheet.Workbook.Worksheet(sName).Cell(cAddress);
                    if (referenceCell.IsEmpty(XLCellsUsedOptions.AllContents))
                        return 0;
                    else
                        return referenceCell.Value;
                }
                finally
                {
                    IsEvaluating = false;
                }
            }

            object retVal;
            try
            {
                IsEvaluating = true;

                if (Worksheet.Workbook.Worksheets.Contains(sName)
                    && XLHelper.IsValidA1Address(cAddress))
                {
                    var referenceCell = Worksheet.Workbook.Worksheet(sName).Cell(cAddress);
                    if (referenceCell.IsEmpty(XLCellsUsedOptions.AllContents))
                        return 0;
                    else
                        return referenceCell.Value;
                }

                retVal = Worksheet.Evaluate(fA1);
            }
            finally
            {
                IsEvaluating = false;
            }

            if (retVal is IEnumerable retValEnumerable && !(retVal is String))
                return retValEnumerable.Cast<object>().First();

            return retVal;
        }

        public void InvalidateFormula()
        {
            NeedsRecalculation = true;
            Worksheet.Workbook.InvalidateFormulas();
            ModifiedAtVersion = Worksheet.Workbook.RecalculationCounter;
        }

        /// <summary>
        /// Perform an evaluation of cell formula. If cell does not contain formula nothing happens, if cell does not need
        /// recalculation (<see cref="NeedsRecalculation"/> is False) nothing happens either, unless <paramref name="force"/> flag is specified.
        /// Otherwise recalculation is performed, result value is preserved in <see cref="CachedValue"/> and returned.
        /// </summary>
        /// <param name="force">Flag indicating whether a recalculation must be performed even is cell does not need it.</param>
        /// <returns>Null if cell does not contain a formula. Calculated value otherwise.</returns>
        public Object Evaluate(Boolean force = false)
        {
            if (force || NeedsRecalculation)
            {
                if (HasFormula)
                {
                    CachedValue = RecalculateFormula(FormulaA1);
                    UpdateCachedValueFromDataType();
                }
                else
                    CachedValue = null;

                EvaluatedAtVersion = Worksheet.Workbook.RecalculationCounter;
                NeedsRecalculation = false;
            }
            return CachedValue;
        }

        internal void SetInternalCellValueString(String cellValue)
        {
            SetInternalCellValueString(cellValue, validate: false, parseToCachedValue: this.HasFormula);
        }

        private void SetInternalCellValueString(String cellValue, Boolean validate, Boolean parseToCachedValue)
        {
            if (validate)
            {
                if (cellValue.Length > 32767) throw new ArgumentOutOfRangeException(nameof(cellValue), "Cells can hold a maximum of 32,767 characters.");
            }

            this._cellValue = cellValue;

            if (parseToCachedValue)
                CachedValue = ParseCellValueFromString();
        }

        private void UpdateCachedValueFromDataType()
        {
            if (CachedValue is double d)
            {
                if (this.DataType == XLDataType.DateTime && d.IsValidOADateNumber())
                    CachedValue = DateTime.FromOADate(d);
                else if (this.DataType == XLDataType.TimeSpan)
                    CachedValue = XLHelper.GetTimeSpan(d);
            }
            else if (CachedValue is DateTime dt)
            {
                if (this.DataType == XLDataType.Number)
                    CachedValue = dt.ToOADate();
                else if (this.DataType == XLDataType.TimeSpan)
                    CachedValue = XLHelper.GetTimeSpan(dt.ToOADate());
            }
            else if (CachedValue is TimeSpan ts)
            {
                if (this.DataType == XLDataType.DateTime)
                    CachedValue = DateTime.FromOADate(ts.TotalDays);
                else if (this.DataType == XLDataType.Number)
                    CachedValue = ts.TotalDays;
            }
        }

        internal void SetDataTypeFast(XLDataType dataType)
        {
            this._dataType = dataType;
        }

        private Object ParseCellValueFromString()
        {
            return ParseCellValueFromString(_cellValue, _dataType, out String error);
        }

        private Object ParseCellValueFromString(String cellValue, XLDataType dataType, out String error)
        {
            error = "";
            if ("" == cellValue)
                return "";

            if (dataType == XLDataType.Boolean)
            {
                if (bool.TryParse(cellValue, out Boolean b))
                    return b;
                else if (cellValue == "0")
                    return false;
                else if (cellValue == "1")
                    return true;
                else
                    return !string.IsNullOrEmpty(cellValue);
            }

            if (dataType == XLDataType.DateTime)
            {
                if (Double.TryParse(cellValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out Double d))
                {
                    if (d.IsValidOADateNumber())
                        return DateTime.FromOADate(d);
                    else
                        return d;
                }
                else if (DateTime.TryParse(cellValue, out DateTime dt))
                    return dt;
                else
                {
                    error = string.Format("Cannot set data type to DateTime because '{0}' is not recognized as a date.", cellValue);
                    return null;
                }
            }
            if (dataType == XLDataType.Number)
            {
                var v = cellValue;
                Double factor = 1.0;
                if (v.EndsWith("%"))
                {
                    v = v.Substring(0, v.Length - 1);
                    factor = 1 / 100.0;
                }

                if (Double.TryParse(v, XLHelper.NumberStyle, CultureInfo.InvariantCulture, out Double d))
                    return d * factor;
                else
                {
                    error = string.Format("Cannot set data type to Number because '{0}' is not recognized as a number.", cellValue);
                    return null;
                }
            }

            if (dataType == XLDataType.TimeSpan)
            {
                if (TimeSpan.TryParse(cellValue, out TimeSpan ts))
                    return ts;
                else if (Double.TryParse(cellValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out Double d))
                    return XLHelper.GetTimeSpan(d);
                else
                {
                    error = string.Format("Cannot set data type to TimeSpan because '{0}' is not recognized as a TimeSpan.", cellValue);
                    return null;
                }
            }

            return cellValue;
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

        public object Value
        {
            get
            {
                if (!String.IsNullOrWhiteSpace(_formulaA1) ||
                    !String.IsNullOrEmpty(_formulaR1C1))
                {
                    return Evaluate();
                }

                var cellValue = HasRichText ? _richText.ToString() : _cellValue;
                return ParseCellValueFromString(cellValue, _dataType, out _);
            }
            set
            {
                if (IsInferiorMergedCell())
                    return;

                FormulaA1 = String.Empty;

                if (value is XLCells) throw new ArgumentException("Cannot assign IXLCells object to the cell value.");

                if (SetTableHeaderValue(value)) return;

                if (SetRangeRows(value)) return;

                if (SetRangeColumns(value)) return;

                if (SetDataTable(value)) return;

                if (SetEnumerable(value)) return;

                if (SetRange(value)) return;

                if (!SetRichText(value))
                    SetValue(value);

                CachedValue = null;

                if (_cellValue.Length > 32767) throw new ArgumentOutOfRangeException(nameof(value), "Cells can hold only 32,767 characters.");
            }
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
            return InsertTableInternal(reader, tableName, createTable, addHeadings, transpose);
        }

        private IXLTable InsertTableInternal(IInsertDataReader reader, String tableName, Boolean createTable, Boolean addHeadings,
            Boolean transpose)
        {
            if (createTable && this.Worksheet.Tables.Any(t => t.Contains(this)))
                throw new InvalidOperationException(String.Format("This cell '{0}' is already part of a table.", this.Address.ToString()));

            var range = InsertDataInternal(reader, addHeadings, transpose);

            if (createTable)
                // Create a table and save it in the file
                return tableName == null ? range.CreateTable() : range.CreateTable(tableName);
            else
                // Create a table, but keep it in memory. Saved file will contain only "raw" data and column headers
                return tableName == null ? range.AsTable() : range.AsTable(tableName);
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

            if (createTable && this.Worksheet.Tables.Any(t => t.Contains(this)))
                throw new InvalidOperationException($"This cell '{this.Address}' is already part of a table.");

            var reader = InsertDataReaderFactory.Instance.CreateReader(data);
            return InsertTableInternal(reader, tableName, createTable, addHeadings: true, transpose: false);
        }

        internal XLRange InsertDataInternal(IInsertDataReader reader, Boolean addHeadings, Boolean transpose)
        {
            if (reader == null)
                return null;

            var currentRowNumber = _rowNumber;
            var currentColumnNumber = _columnNumber;
            var maximumColumnNumber = currentColumnNumber;
            var maximumRowNumber = currentRowNumber;

            if (transpose)
            {
                maximumColumnNumber += reader.GetRecordsCount() - 1;
                maximumRowNumber += reader.GetPropertiesCount() - 1;
            }
            else
            {
                maximumColumnNumber += reader.GetPropertiesCount() - 1;
                maximumRowNumber += reader.GetRecordsCount() - 1;
            }

            // Inline functions to handle looping with transposing
            //////////////////////////////////////////////////////
            void incrementFieldPosition()
            {
                if (transpose)
                {
                    maximumRowNumber = Math.Max(maximumRowNumber, currentRowNumber);
                    currentRowNumber++;
                }
                else
                {
                    maximumColumnNumber = Math.Max(maximumColumnNumber, currentColumnNumber);
                    currentColumnNumber++;
                }
            }

            void incrementRecordPosition()
            {
                if (transpose)
                {
                    maximumColumnNumber = Math.Max(maximumColumnNumber, currentColumnNumber);
                    currentColumnNumber++;
                }
                else
                {
                    maximumRowNumber = Math.Max(maximumRowNumber, currentRowNumber);
                    currentRowNumber++;
                }
            }

            void resetRecordPosition()
            {
                if (transpose)
                    currentRowNumber = _rowNumber;
                else
                    currentColumnNumber = _columnNumber;
            }
            //////////////////////////////////////////////////////

            var empty = maximumRowNumber <= _rowNumber ||
                        maximumColumnNumber <= _columnNumber;

            if (!empty)
            {
                Worksheet.Range(
                        _rowNumber,
                        _columnNumber,
                        maximumRowNumber,
                        maximumColumnNumber)
                    .Clear();
            }

            if (addHeadings)
            {
                for (int i = 0; i < reader.GetPropertiesCount(); i++)
                {
                    var propertyName = reader.GetPropertyName(i);
                    Worksheet.SetValue(propertyName, currentRowNumber, currentColumnNumber);
                    incrementFieldPosition();
                }

                incrementRecordPosition();
            }

            var data = reader.GetData();

            foreach (var item in data)
            {
                resetRecordPosition();
                foreach (var value in item)
                {
                    Worksheet.SetValue(value, currentRowNumber, currentColumnNumber);
                    incrementFieldPosition();
                }
                incrementRecordPosition();
            }

            var range = Worksheet.Range(
                _rowNumber,
                _columnNumber,
                maximumRowNumber,
                maximumColumnNumber);

            return range;
        }

        public XLTableCellType TableCellType()
        {
            var table = this.Worksheet.Tables.FirstOrDefault(t => t.AsRange().Contains(this));
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
            return InsertDataInternal(reader, addHeadings: false, transpose: transpose);
        }

        public IXLRange InsertData(DataTable dataTable)
        {
            if (dataTable == null)
                return null;

            var reader = InsertDataReaderFactory.Instance.CreateReader(dataTable);
            return InsertDataInternal(reader, addHeadings: false, transpose: false);
        }

        public IXLCell SetDataType(XLDataType dataType)
        {
            DataType = dataType;
            return this;
        }

        public XLDataType DataType
        {
            get { return _dataType; }
            set
            {
                if (_dataType == value) return;

                if (HasRichText)
                {
                    _cellValue = _richText.ToString();
                    _richText = null;
                }

                if (!string.IsNullOrEmpty(_cellValue))
                {
                    // If we're converting the DataType to Text, there are some quirky rules currently
                    if (value == XLDataType.Text)
                    {
                        var v = Value;
                        switch (v)
                        {
                            case DateTime d:
                                _cellValue = d.ToOADate().ToInvariantString();
                                break;

                            case TimeSpan ts:
                                _cellValue = ts.TotalDays.ToInvariantString();
                                break;

                            case Boolean b:
                                _cellValue = b ? "True" : "False";
                                break;

                            default:
                                _cellValue = v.ObjectToInvariantString();
                                break;
                        }
                    }
                    else
                    {
                        var v = ParseCellValueFromString(_cellValue, value, out String error);

                        if (!String.IsNullOrWhiteSpace(error))
                            throw new ArgumentException(error, nameof(value));

                        _cellValue = v?.ObjectToInvariantString() ?? "";

                        var style = GetStyleForRead();
                        switch (v)
                        {
                            case DateTime d:
                                _cellValue = d.ToOADate().ToInvariantString();

                                if (style.NumberFormat.Format.Length == 0 && style.NumberFormat.NumberFormatId == 0)
                                    Style.NumberFormat.NumberFormatId = _cellValue.Contains('.') ? 22 : 14;

                                break;

                            case TimeSpan ts:
                                if (style.NumberFormat.Format.Length == 0 && style.NumberFormat.NumberFormatId == 0)
                                    Style.NumberFormat.NumberFormatId = 46;

                                break;

                            case Boolean b:
                                _cellValue = b ? "1" : "0";
                                break;
                        }
                    }
                }

                _dataType = value;

                if (HasFormula && !NeedsRecalculation)
                    UpdateCachedValueFromDataType();
                else
                    CachedValue = null;
            }
        }

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
                    _richText = null;
                    _cellValue = String.Empty;
                    FormulaA1 = String.Empty;
                }

                if (clearOptions.HasFlag(XLClearOptions.DataType))
                    _dataType = XLDataType.Text;

                if (clearOptions.HasFlag(XLClearOptions.NormalFormats))
                    SetStyle(Worksheet.Style);

                if (clearOptions.HasFlag(XLClearOptions.ConditionalFormats))
                {
                    AsRange().RemoveConditionalFormatting();
                }

                if (clearOptions.HasFlag(XLClearOptions.Comments))
                    _comment = null;

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
            get
            {
                if (String.IsNullOrWhiteSpace(_formulaA1))
                {
                    if (!String.IsNullOrWhiteSpace(_formulaR1C1))
                    {
                        _formulaA1 = GetFormulaA1(_formulaR1C1);
                        return FormulaA1;
                    }

                    return String.Empty;
                }

                if (_formulaA1.Trim()[0] == '=')
                    return _formulaA1.Substring(1);

                if (_formulaA1.Trim().StartsWith("{="))
                    return "{" + _formulaA1.Substring(2);

                return _formulaA1;
            }

            set
            {
                if (IsInferiorMergedCell())
                    return;

                InvalidateFormula();

                _formulaA1 = String.IsNullOrWhiteSpace(value) ? null : value;

                _formulaR1C1 = null;
            }
        }

        public string FormulaR1C1
        {
            get
            {
                if (String.IsNullOrWhiteSpace(_formulaR1C1))
                    _formulaR1C1 = GetFormulaR1C1(FormulaA1);

                return _formulaR1C1;
            }

            set
            {
                if (IsInferiorMergedCell())
                    return;

                InvalidateFormula();

                _formulaR1C1 = String.IsNullOrWhiteSpace(value) ? null : value;

                _formulaA1 = null;
            }
        }

        public bool ShareString { get; set; }

        public XLHyperlink GetHyperlink()
        {
            return _hyperlink ?? CreateHyperlink();
        }

        public void SetHyperlink(XLHyperlink hyperlink)
        {
            Worksheet.Hyperlinks.TryDelete(Address);

            _hyperlink = hyperlink;

            if (_hyperlink == null) return;

            _hyperlink.Worksheet = Worksheet;
            _hyperlink.Cell = this;

            Worksheet.Hyperlinks.Add(_hyperlink);

            if (SettingHyperlink) return;

            if (GetStyleForRead().Font.FontColor.Equals(Worksheet.StyleValue.Font.FontColor))
                Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);

            if (GetStyleForRead().Font.Underline == Worksheet.StyleValue.Font.Underline)
                Style.Font.Underline = XLFontUnderlineValues.Single;
        }

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

        private bool _recalculationNeededLastValue;

        /// <summary>
        /// Flag indicating that previously calculated cell value may be not valid anymore and has to be re-evaluated.
        /// </summary>
        public bool NeedsRecalculation
        {
            get
            {
                if (String.IsNullOrWhiteSpace(_formulaA1) && String.IsNullOrEmpty(_formulaR1C1))
                    return false;

                if (NeedsRecalculationEvaluatedAtVersion == Worksheet.Workbook.RecalculationCounter)
                    return _recalculationNeededLastValue;

                bool res = EvaluatedAtVersion < ModifiedAtVersion ||                                       // the cell itself was modified
                           GetAffectingCells().Any(cell => cell.ModifiedAtVersion > EvaluatedAtVersion ||  // the affecting cell was modified after this one was evaluated
                                                           cell.EvaluatedAtVersion > EvaluatedAtVersion || // the affecting cell was evaluated after this one (normally this should not happen)
                                                           cell.NeedsRecalculation);                       // the affecting cell needs recalculation (recursion to walk through dependencies)

                NeedsRecalculation = res;
                return res;
            }
            internal set
            {
                _recalculationNeededLastValue = value;
                NeedsRecalculationEvaluatedAtVersion = Worksheet.Workbook.RecalculationCounter;
            }
        }

        private IEnumerable<XLCell> GetAffectingCells()
        {
            return Worksheet.CalcEngine.GetPrecedentCells(_formulaA1).Cast<XLCell>();
        }

        /// <summary>
        /// The value of <see cref="XLWorkbook.RecalculationCounter"/> that workbook had at the moment of cell last modification.
        /// If this value is greater than <see cref="EvaluatedAtVersion"/> then cell needs re-evaluation, as well as all dependent cells do.
        /// </summary>
        private long ModifiedAtVersion { get; set; }

        /// <summary>
        /// The value of <see cref="XLWorkbook.RecalculationCounter"/> that workbook had at the moment of cell formula evaluation.
        /// If this value equals to <see cref="XLWorkbook.RecalculationCounter"/> it indicates that <see cref="CachedValue"/> stores
        /// correct value and no re-evaluation has to be performed.
        /// </summary>
        private long EvaluatedAtVersion { get; set; }

        /// <summary>
        /// The value of <see cref="XLWorkbook.RecalculationCounter"/> that workbook had at the moment of determining whether the cell
        /// needs re-evaluation (due to it has been edited or some of the affecting cells has). If this value equals to <see cref="XLWorkbook.RecalculationCounter"/>
        /// it indicates that <see cref="_recalculationNeededLastValue"/> stores correct value and no check has to be performed.
        /// </summary>
        private long NeedsRecalculationEvaluatedAtVersion { get; set; }

        private Object cachedValue;

        public Object CachedValue
        {
            get
            {
                if (!HasFormula && cachedValue == null)
                    cachedValue = Value;

                return cachedValue;
            }
            private set
            {
                if (value != null && !HasFormula)
                    throw new InvalidOperationException("Cached values can be set only for cells with formulas");

                cachedValue = value;
            }
        }

        public IXLRichText GetRichText()
        {
            return _richText ?? CreateRichText();
        }

        public bool HasRichText
        {
            get { return _richText != null; }
        }

        public IXLRichText CreateRichText()
        {
            var style = GetStyleForRead();
            _richText = _cellValue.Length == 0
                            ? new XLRichText(new XLFont(Style as XLStyle, style.Font))
                            : new XLRichText(GetFormattedString(), new XLFont(Style as XLStyle, style.Font));

            return _richText;
        }

        IXLComment IXLCell.GetComment()
        {
            return GetComment();
        }

        public bool HasComment
        {
            get { return _comment != null; }
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

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public Boolean IsEmpty(Boolean includeFormats)
        {
            return IsEmpty(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
        }

        public Boolean IsEmpty(XLCellsUsedOptions options)
        {
            if (InnerText.Length > 0)
                return false;

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
            Worksheet.DataValidations.TryGet(AsRange().RangeAddress, out var dataValidation);
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
            get { return Worksheet.ActiveCell == this; }
            set
            {
                if (value)
                    Worksheet.ActiveCell = this;
                else if (Active)
                    Worksheet.ActiveCell = null;
            }
        }

        public IXLCell SetActive(Boolean value = true)
        {
            Active = value;
            return this;
        }

        public Boolean HasHyperlink
        {
            get { return _hyperlink != null; }
        }

        public Boolean TryGetValue<T>(out T value)
        {
            var targetType = typeof(T);
            var underlyingType = targetType.GetUnderlyingType();
            var isNullable = targetType.IsNullableType();

            Object currentValue;
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

            if (isNullable && (currentValue == null || currentValue is string s && String.IsNullOrEmpty(s)))
            {
                value = default;
                return true;
            }

            if (targetType != typeof(String) // Strings are handled later and have some specifics to UTF handling
                && currentValue is T t)
            {
                value = t;
                return true;
            }

            if (TryGetDateTimeValue(out value, currentValue)) return true;

            if (TryGetTimeSpanValue(out value, currentValue)) return true;

            if (TryGetBooleanValue(out value, currentValue)) return true;

            if (TryGetRichStringValue(out value)) return true;

            if (TryGetStringValue(out value, currentValue)) return true;

            if (TryGetHyperlink(out value)) return true;

            if (currentValue.IsNumber())
            {
                try
                {
                    value = (T)Convert.ChangeType(currentValue, underlyingType);
                    return true;
                }
                catch (Exception)
                {
                    value = default;
                    return false;
                }
            }

            var strValue = currentValue.ToString();

            if (underlyingType == typeof(sbyte)) return TryGetBasicValue<T, sbyte>(strValue, sbyte.TryParse, out value);
            if (underlyingType == typeof(byte)) return TryGetBasicValue<T, byte>(strValue, byte.TryParse, out value);
            if (underlyingType == typeof(short)) return TryGetBasicValue<T, short>(strValue, short.TryParse, out value);
            if (underlyingType == typeof(ushort)) return TryGetBasicValue<T, ushort>(strValue, ushort.TryParse, out value);
            if (underlyingType == typeof(int)) return TryGetBasicValue<T, int>(strValue, int.TryParse, out value);
            if (underlyingType == typeof(uint)) return TryGetBasicValue<T, uint>(strValue, uint.TryParse, out value);
            if (underlyingType == typeof(long)) return TryGetBasicValue<T, long>(strValue, long.TryParse, out value);
            if (underlyingType == typeof(ulong)) return TryGetBasicValue<T, ulong>(strValue, ulong.TryParse, out value);
            if (underlyingType == typeof(float)) return TryGetBasicValue<T, float>(strValue, float.TryParse, out value);
            if (underlyingType == typeof(double)) return TryGetBasicValue<T, double>(strValue, double.TryParse, out value);
            if (underlyingType == typeof(decimal)) return TryGetBasicValue<T, decimal>(strValue, decimal.TryParse, out value);

            if (underlyingType.IsEnum)
            {
                if (Enum.IsDefined(underlyingType, strValue))
                {
                    value = (T)Enum.Parse(underlyingType, strValue, ignoreCase: false);
                    return true;
                }
                value = default;
                return false;
            }

            try
            {
                value = (T)Convert.ChangeType(currentValue, targetType);
                return true;
            }
            catch
            {
                value = default;
                return false;
            }
        }

        private static bool TryGetDateTimeValue<T>(out T value, object currentValue)
        {
            if (typeof(T) != typeof(DateTime) && typeof(T) != typeof(DateTime?))
            {
                value = default;
                return false;
            }

            if (currentValue is T v) { value = v; return true; }

            if (currentValue.IsNumber())
            {
                var dbl1 = Convert.ToDouble(currentValue);
                if (dbl1.IsValidOADateNumber())
                {
                    value = (T)Convert.ChangeType(DateTime.FromOADate(dbl1), typeof(T));
                    return true;
                }
            }

            if (DateTime.TryParse(currentValue.ToString(), out DateTime ts))
            {
                value = (T)Convert.ChangeType(ts, typeof(T));
                return true;
            }

            // If the cell value is a string, e.g. "42020", we could theoretically coerce it to a DateTime, but this seems to go against what is expected.
            // Leaving this code block here, though. Maybe we revert our decision later.

            //if (Double.TryParse(currentValue.ObjectToInvariantString(), out Double dbl2) && dbl2.IsValidOADateNumber())
            //{
            //    value = (T)Convert.ChangeType(DateTime.FromOADate(dbl2), typeof(T));
            //    return true;
            //}

            value = default;
            return false;
        }

        private static bool TryGetTimeSpanValue<T>(out T value, object currentValue)
        {
            if (typeof(T) != typeof(TimeSpan) && typeof(T) != typeof(TimeSpan?))
            {
                value = default;
                return false;
            }

            if (currentValue is T v) { value = v; return true; }

            if (!TimeSpan.TryParse(currentValue.ToString(), out TimeSpan ts))
            {
                value = default;
                return false;
            }

            value = (T)Convert.ChangeType(ts, typeof(T));
            return true;
        }

        private bool TryGetRichStringValue<T>(out T value)
        {
            if (typeof(T) == typeof(IXLRichText))
            {
                value = (T)GetRichText();
                return true;
            }
            value = default;
            return false;
        }

        private static bool TryGetStringValue<T>(out T value, object currentValue)
        {
            if (typeof(T) == typeof(String))
            {
                var s = currentValue.ToString();
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

        private static Boolean TryGetBooleanValue<T>(out T value, object currentValue)
        {
            if (typeof(T) != typeof(Boolean) && typeof(T) != typeof(Boolean?))
            {
                value = default;
                return false;
            }

            if (currentValue is T v) { value = v; return true; }

            if (!Boolean.TryParse(currentValue.ToString(), out Boolean b))
            {
                value = default;
                return false;
            }

            value = (T)Convert.ChangeType(b, typeof(T));
            return true;
        }

        private Boolean TryGetHyperlink<T>(out T value)
        {
            if (typeof(T) == typeof(XLHyperlink))
            {
                var hyperlink = _hyperlink ?? Value as XLHyperlink;
                if (hyperlink != null)
                {
                    value = (T)Convert.ChangeType(hyperlink, typeof(T));
                    return true;
                }
            }

            value = default;
            return false;
        }

        private delegate Boolean ParseFunction<T>(String s, NumberStyles style, IFormatProvider provider, out T result);

        private static Boolean TryGetBasicValue<T, U>(String currentValue, ParseFunction<U> parseFunction, out T value)
        {
            if (parseFunction.Invoke(currentValue, NumberStyles.Any, null, out U result))
            {
                value = (T)Convert.ChangeType(result, typeof(T).GetUnderlyingType());
                return true;
            }

            value = default;
            return false;
        }

        #endregion IXLCell Members

        #region IXLStylized Members

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                yield return Style;
            }
        }

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

        private Boolean SetTableHeaderValue(object value)
        {
            foreach (var table in Worksheet.Tables.Where(t => t.ShowHeaderRow))
            {
                var cell = table.HeadersRow().CellsUsed(c => c.Address.Equals(this.Address)).FirstOrDefault();
                if (cell != null)
                {
                    var oldName = cell.GetString();
                    var field = table.Field(oldName);
                    field.Name = value.ToString();
                    return true;
                }
            }

            return false;
        }

        private Boolean SetTableTotalsRowLabel(object value)
        {
            foreach (var table in Worksheet.Tables.Where(t => t.ShowTotalsRow))
            {
                var cell = table.TotalsRow().Cells(c => c.Address.Equals(this.Address)).FirstOrDefault();
                if (cell != null)
                {
                    var field = table.Fields.First(f => f.Column.ColumnNumber() == cell.WorksheetColumn().ColumnNumber());
                    field.TotalsRowFunction = XLTotalsRowFunction.None;

                    SetInternalCellValueString(value.ObjectToInvariantString(), validate: true, parseToCachedValue: false);

                    field.TotalsRowLabel = _cellValue;
                    this.DataType = XLDataType.Text;
                    return true;
                }
            }

            return false;
        }

        private bool SetRangeColumns(object value)
        {
            if (value is XLRangeColumns columns)
            {
                var cell = this;
                foreach (var column in columns)
                {
                    cell.SetRange(column);
                    cell = cell.CellRight();
                }
                return true;
            }
            else

                return SetColumns(value);
        }

        private bool SetColumns(object value)
        {
            if (value is XLColumns columns)
            {
                var cell = this;
                foreach (var column in columns)
                {
                    cell.SetRange(column);
                    cell = cell.CellRight();
                }
                return true;
            }
            else
                return false;
        }

        private bool SetRangeRows(object value)
        {
            if (value is XLRangeRows rows)
            {
                var cell = this;
                foreach (var row in rows)
                {
                    cell.SetRange(row);
                    cell = cell.CellBelow();
                }
                return true;
            }
            else
                return SetRows(value);
        }

        private bool SetRows(object value)
        {
            if (value is XLRows rows)
            {
                var cell = this;
                foreach (var row in rows)
                {
                    cell.SetRange(row);
                    cell = cell.CellBelow();
                }
                return true;
            }
            else
                return false;
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

        private bool IsDateFormat()
        {
            var style = GetStyleForRead();
            return _dataType == XLDataType.Number
                   && String.IsNullOrWhiteSpace(style.NumberFormat.Format)
                   && ((style.NumberFormat.NumberFormatId >= 14
                        && style.NumberFormat.NumberFormatId <= 22)
                       || (style.NumberFormat.NumberFormatId >= 45
                           && style.NumberFormat.NumberFormatId <= 47));
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

        private bool SetRichText(object value)
        {
            if (value is XLRichText asRichString)
            {
                _richText = asRichString;
                _dataType = XLDataType.Text;
                return true;
            }
            else
                return false;
        }

        private Boolean SetRange(Object rangeObject)
        {
            var asRange = (rangeObject as XLRangeBase)
                       ?? (rangeObject as XLCell)?.AsRange();

            if (asRange != null)
            {
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
                        XLCellCopyOptions.All & ~XLCellCopyOptions.ConditionalFormats); //Conditional formats are copied separately
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

                return true;
            }

            return false;
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

        private void SetValue(object value)
        {
            if (value == null)
            {
                this.Clear(XLClearOptions.Contents);
                return;
            }

            FormulaA1 = String.Empty;
            _richText = null;

            var style = GetStyleForRead();
            Boolean parsed = false;
            string parsedValue = string.Empty;

            ////
            // Try easy parsing first. If that doesn't work, we'll have to ToString it and parse it slowly

            // When number format starts with @, we treat any value as text - no parsing required
            // This doesn't happen in the SetValue<T>() version
            if (style.NumberFormat.Format == "@")
            {
                parsedValue = value.ObjectToInvariantString();

                _dataType = XLDataType.Text;
                if (parsedValue.Contains(Environment.NewLine) && !style.Alignment.WrapText)
                    Style.Alignment.WrapText = true;

                parsed = true;
            }
            // Don't accept strings, because we're going to try to parse them later
            else if (value is not string _)
            {
                parsed = TrySetInternallySupportedType(ConvertOtherSupportedTypes(value), style, out parsedValue);
            }

            ////
            if (!parsed)
            {
                // We'll have to parse it slowly :-(
                parsedValue = DeduceCellValueByParsing(value.ToString(), style);
            }

            if (SetTableHeaderValue(parsedValue)) return;
            if (SetTableTotalsRowLabel(parsedValue)) return;

            SetInternalCellValueString(parsedValue, validate: true, parseToCachedValue: false);
            CachedValue = null;
        }

        private void SetDateTimeFormat(XLStyleValue style, Boolean onlyDatePart)
        {
            _dataType = XLDataType.DateTime;

            if (style.NumberFormat.Format.Length == 0 && style.NumberFormat.NumberFormatId == 0)
                Style.NumberFormat.NumberFormatId = onlyDatePart ? 14 : 22;
        }

        private void SetTimeSpanFormat(XLStyleValue style)
        {
            _dataType = XLDataType.TimeSpan;

            if (style.NumberFormat.Format.Length == 0 && style.NumberFormat.NumberFormatId == 0)
                Style.NumberFormat.NumberFormatId = 46;
        }

        internal string GetFormulaR1C1(string value)
        {
            return GetFormula(value, FormulaConversionType.A1ToR1C1, 0, 0);
        }

        internal string GetFormulaA1(string value)
        {
            return GetFormula(value, FormulaConversionType.R1C1ToA1, 0, 0);
        }

        private string GetFormula(string strValue, FormulaConversionType conversionType, int rowsToShift,
                                  int columnsToShift)
        {
            if (String.IsNullOrWhiteSpace(strValue))
                return String.Empty;

            var value = ">" + strValue + "<";

            var regex = conversionType == FormulaConversionType.A1ToR1C1 ? A1Regex : R1C1Regex;

            var sb = new StringBuilder();
            var lastIndex = 0;

            foreach (var match in regex.Matches(value).Cast<Match>())
            {
                var matchString = match.Value;
                var matchIndex = match.Index;
                if (value.Substring(0, matchIndex).CharCount('"') % 2 == 0
                    && value.Substring(0, matchIndex).CharCount('\'') % 2 == 0)
                {
                    // Check if the match is in between quotes
                    sb.Append(value.Substring(lastIndex, matchIndex - lastIndex));
                    sb.Append(conversionType == FormulaConversionType.A1ToR1C1
                        ? GetR1C1Address(matchString, rowsToShift, columnsToShift)
                        : GetA1Address(matchString, rowsToShift, columnsToShift));
                }
                else
                    sb.Append(value.Substring(lastIndex, matchIndex - lastIndex + matchString.Length));
                lastIndex = matchIndex + matchString.Length;
            }

            if (lastIndex < value.Length)
                sb.Append(value.Substring(lastIndex));

            var retVal = sb.ToString();
            return retVal.Substring(1, retVal.Length - 2);
        }

        private string GetA1Address(string r1C1Address, int rowsToShift, int columnsToShift)
        {
            var addressToUse = r1C1Address.ToUpper();

            if (addressToUse.Contains(':'))
            {
                var parts = addressToUse.Split(':');
                var p1 = parts[0];
                var p2 = parts[1];
                string leftPart;
                string rightPart;
                if (p1.StartsWith("R"))
                {
                    leftPart = GetA1Row(p1, rowsToShift);
                    rightPart = GetA1Row(p2, rowsToShift);
                }
                else
                {
                    leftPart = GetA1Column(p1, columnsToShift);
                    rightPart = GetA1Column(p2, columnsToShift);
                }

                return leftPart + ":" + rightPart;
            }

            try
            {
                var rowPart = addressToUse.Substring(0, addressToUse.IndexOf("C"));
                var rowToReturn = GetA1Row(rowPart, rowsToShift);

                var columnPart = addressToUse.Substring(addressToUse.IndexOf("C"));
                var columnToReturn = GetA1Column(columnPart, columnsToShift);

                var retAddress = columnToReturn + rowToReturn;
                return retAddress;
            }
            catch (ArgumentOutOfRangeException)
            {
                return "#REF!";
            }
        }

        private string GetA1Column(string columnPart, int columnsToShift)
        {
            string columnToReturn;
            if (columnPart == "C")
                columnToReturn = XLHelper.GetColumnLetterFromNumber(_columnNumber + columnsToShift);
            else
            {
                var bIndex = columnPart.IndexOf("[");
                var mIndex = columnPart.IndexOf("-");
                if (bIndex >= 0)
                {
                    columnToReturn = XLHelper.GetColumnLetterFromNumber(
                        _columnNumber +
                        Int32.Parse(columnPart.Substring(bIndex + 1, columnPart.Length - bIndex - 2)) + columnsToShift
                        );
                }
                else if (mIndex >= 0)
                {
                    columnToReturn = XLHelper.GetColumnLetterFromNumber(
                        _columnNumber + Int32.Parse(columnPart.Substring(mIndex)) + columnsToShift
                        );
                }
                else
                {
                    columnToReturn = "$" +
                                     XLHelper.GetColumnLetterFromNumber(Int32.Parse(columnPart.Substring(1)) +
                                                                        columnsToShift);
                }
            }

            return columnToReturn;
        }

        private string GetA1Row(string rowPart, int rowsToShift)
        {
            string rowToReturn;
            if (rowPart == "R")
                rowToReturn = (_rowNumber + rowsToShift).ToString();
            else
            {
                var bIndex = rowPart.IndexOf("[");
                if (bIndex >= 0)
                {
                    rowToReturn =
                        (_rowNumber + Int32.Parse(rowPart.Substring(bIndex + 1, rowPart.Length - bIndex - 2)) +
                         rowsToShift).ToString();
                }
                else
                    rowToReturn = "$" + (Int32.Parse(rowPart.Substring(1)) + rowsToShift);
            }

            return rowToReturn;
        }

        private string GetR1C1Address(string a1Address, int rowsToShift, int columnsToShift)
        {
            if (a1Address.Contains(':'))
            {
                var parts = a1Address.Split(':');
                var p1 = parts[0];
                var p2 = parts[1];
                if (Int32.TryParse(p1.Replace("$", string.Empty), out Int32 row1))
                {
                    var row2 = Int32.Parse(p2.Replace("$", string.Empty));
                    var leftPart = GetR1C1Row(row1, p1.Contains('$'), rowsToShift);
                    var rightPart = GetR1C1Row(row2, p2.Contains('$'), rowsToShift);
                    return leftPart + ":" + rightPart;
                }
                else
                {
                    var column1 = XLHelper.GetColumnNumberFromLetter(p1.Replace("$", string.Empty));
                    var column2 = XLHelper.GetColumnNumberFromLetter(p2.Replace("$", string.Empty));
                    var leftPart = GetR1C1Column(column1, p1.Contains('$'), columnsToShift);
                    var rightPart = GetR1C1Column(column2, p2.Contains('$'), columnsToShift);
                    return leftPart + ":" + rightPart;
                }
            }

            var address = XLAddress.Create(Worksheet, a1Address);

            var rowPart = GetR1C1Row(address.RowNumber, address.FixedRow, rowsToShift);
            var columnPart = GetR1C1Column(address.ColumnNumber, address.FixedColumn, columnsToShift);

            return rowPart + columnPart;
        }

        private string GetR1C1Row(int rowNumber, bool fixedRow, int rowsToShift)
        {
            string rowPart;
            rowNumber += rowsToShift;
            var rowDiff = rowNumber - _rowNumber;
            if (rowDiff != 0 || fixedRow)
                rowPart = fixedRow ? "R" + rowNumber : "R[" + rowDiff + "]";
            else
                rowPart = "R";

            return rowPart;
        }

        private string GetR1C1Column(int columnNumber, bool fixedColumn, int columnsToShift)
        {
            string columnPart;
            columnNumber += columnsToShift;
            var columnDiff = columnNumber - _columnNumber;
            if (columnDiff != 0 || fixedColumn)
                columnPart = fixedColumn ? "C" + columnNumber : "C[" + columnDiff + "]";
            else
                columnPart = "C";

            return columnPart;
        }

        internal void CopyValuesFrom(XLCell source)
        {
            _cellValue = source._cellValue;
            _dataType = source._dataType;
            FormulaR1C1 = source.FormulaR1C1;
            _richText = source._richText == null ? null : new XLRichText(source._richText, source.Style.Font);
            _comment = source._comment == null ? null : new XLComment(this, source._comment, source.Style.Font, source._comment.Style);
            if (source._hyperlink != null)
            {
                SettingHyperlink = true;
                SetHyperlink(new XLHyperlink(source.GetHyperlink()));
                SettingHyperlink = false;
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
                : otherCell.Sparkline.SourceData.Worksheet;
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
            var eventTracking = Worksheet.EventTrackingEnabled;
            Worksheet.EventTrackingEnabled = false;
            if (otherCell.HasDataValidation)
                CopyDataValidation(otherCell, otherCell.GetDataValidation());
            else if (HasDataValidation)
            {
                Worksheet.DataValidations.Delete(AsRange());
            }
            Worksheet.EventTrackingEnabled = eventTracking;
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

        #region Nested type: FormulaConversionType

        private enum FormulaConversionType
        {
            A1ToR1C1,
            R1C1ToA1
        };

        #endregion Nested type: FormulaConversionType

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

        public Boolean HasFormula
        { get { return !String.IsNullOrWhiteSpace(FormulaA1); } }

        public Boolean HasArrayFormula
        { get { return FormulaA1.StartsWith("{"); } }

        public IXLRangeAddress FormulaReference { get; set; }

        public IXLRange CurrentRegion
        {
            get
            {
                return this.Worksheet.Range(FindCurrentRegion(this.AsRange()));
            }
        }

        internal IXLRangeAddress FindCurrentRegion(IXLRangeBase range)
        {
            var rangeAddress = range.RangeAddress;

            var filledCells = range
                .SurroundingCells(c => !(c as XLCell).IsEmpty(XLCellsUsedOptions.AllContents))
                .Concat(this.Worksheet.Range(rangeAddress).Cells());

            var grownRangeAddress = new XLRangeAddress(
                new XLAddress(this.Worksheet, filledCells.Min(c => c.Address.RowNumber), filledCells.Min(c => c.Address.ColumnNumber), false, false),
                new XLAddress(this.Worksheet, filledCells.Max(c => c.Address.RowNumber), filledCells.Max(c => c.Address.ColumnNumber), false, false)
            );

            if (rangeAddress.Equals(grownRangeAddress))
                return this.Worksheet.Range(grownRangeAddress).RangeAddress;
            else
                return FindCurrentRegion(this.Worksheet.Range(grownRangeAddress));
        }

        internal bool IsInferiorMergedCell()
        {
            return this.IsMerged() && !this.Address.Equals(this.MergedRange().RangeAddress.FirstAddress);
        }

        internal bool IsSuperiorMergedCell()
        {
            return this.IsMerged() && this.Address.Equals(this.MergedRange().RangeAddress.FirstAddress);
        }
    }
}
