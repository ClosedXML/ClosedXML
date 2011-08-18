namespace ClosedXML.Excel
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Data;
    using System.Globalization;
    using System.Linq;
    using System.Reflection;
    using System.Text;
    using System.Text.RegularExpressions;

    internal partial class XLCell : IXLCell, IXLStylized
    {
        public Boolean StyleChanged { get; set; }
        public static readonly DateTime BaseDate = new DateTime(1899, 12, 30);
        private static Dictionary<int, string> _formatCodes;

        private static readonly Regex A1Regex = new Regex(
            @"(?<=\W)(\$?[a-zA-Z]{1,3}\$?\d{1,7})(?=\W)" // A1
            + @"|(?<=\W)(\d{1,7}:\d{1,7})(?=\W)" // 1:1
            + @"|(?<=\W)([a-zA-Z]{1,3}:[a-zA-Z]{1,3})(?=\W)"); // A:A

        private static readonly Regex A1SimpleRegex = new Regex(
            @"(?<=\W)" // Start with non word
            + @"(" // Start Group to pick
            + @"(" // Start Sheet Name, optional
            + @"("
            + @"\'[^\[\]\*/\\\?:]+\'" // Sheet name with special characters, surrounding apostrophes are required
            + @"|"
            + @"\'?\w+\'?" // Sheet name with letters and numbers, surrounding apostrophes are optional
            + @")"
            + @"!)?" // End Sheet Name, optional
            + @"(" // Start range
            + @"\$?[a-zA-Z]{1,3}\$?\d{1,7}" // A1 Address 1
            + @"(:\$?[a-zA-Z]{1,3}\$?\d{1,7})?" // A1 Address 2, optional
            + @"|"
            + @"(\d{1,7}:\d{1,7})" // 1:1
            + @"|"
            + @"([a-zA-Z]{1,3}:[a-zA-Z]{1,3})" // A:A
            + @")" // End Range
            + @")" // End Group to pick
            + @"(?=\W)" // End with non word
            );

        private static readonly Regex A1RowRegex = new Regex(
            @"(\d{1,7}:\d{1,7})" // 1:1
            );

        private static readonly Regex A1ColumnRegex = new Regex(
            @"([a-zA-Z]{1,3}:[a-zA-Z]{1,3})" // A:A
            );

        private static readonly Regex R1C1Regex = new Regex(
            @"(?<=\W)([Rr]\[?-?\d{0,7}\]?[Cc]\[?-?\d{0,7}\]?)(?=\W)" // R1C1
            + @"|(?<=\W)([Rr]\[?-?\d{0,7}\]?:[Rr]\[?-?\d{0,7}\]?)(?=\W)" // R:R
            + @"|(?<=\W)([Cc]\[?-?\d{0,5}\]?:[Cc]\[?-?\d{0,5}\]?)(?=\W)"); // C:C

        #region Fields

        private readonly XLWorksheet _worksheet;

        internal string _cellValue = String.Empty;
        internal XLCellValues _dataType;
        private XLHyperlink _hyperlink;
        private XLRichText _richText;

        #endregion

        #region Constructor

        private Int32 _styleCacheId;
        public XLCell(XLWorksheet worksheet, XLAddress address, IXLStyle defaultStyle)
        {
            Address = address;
            ShareString = true;
            
            _worksheet = worksheet;

            SetStyle(defaultStyle ?? worksheet.Style);

        }

        private IXLStyle GetStyleForRead()
        {
            return Worksheet.Workbook.GetStyleById(GetStyleId());
        }
        public Int32 GetStyleId()
        {
            if (StyleChanged)
                SetStyle(Style);

            return _styleCacheId;
        }
        private void SetStyle(IXLStyle styleToUse)
        {
            _styleCacheId = Worksheet.Workbook.GetStyleId(styleToUse);
            _style = null;
            StyleChanged = false;
        }

        #endregion

        public bool SettingHyperlink;
        public int SharedStringId;
        private string _formulaA1;
        private string _formulaR1C1;
        private IXLStyle _style;

        public XLWorksheet Worksheet
        {
            get { return _worksheet; }
        }

        public XLAddress Address { get; internal set; }

        public string InnerText
        {
            get
            {
                if (HasRichText)
                    return _richText.ToString();

                return StringExtensions.IsNullOrWhiteSpace(_cellValue) ? FormulaA1 : _cellValue;
            }
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

        public XLRange AsRange()
        {
            return _worksheet.Range(Address, Address);
        }

        IXLRange IXLCell.AsRange()
        {
            return AsRange();
        }

        public IXLCell SetValue<T>(T value)
        {
            FormulaA1 = String.Empty;
            _richText = null;
            if (value is String)
            {
                _cellValue = value.ToString();
                _dataType = XLCellValues.Text;
                if (_cellValue.Contains(Environment.NewLine) && !GetStyleForRead().Alignment.WrapText)
                    Style.Alignment.WrapText = true;
            }
            else if (value is TimeSpan)
            {
                _cellValue = value.ToString();
                _dataType = XLCellValues.TimeSpan;
                Style.NumberFormat.NumberFormatId = 46;
            }
            else if (value is DateTime)
            {
                _dataType = XLCellValues.DateTime;
                var dtTest = (DateTime)Convert.ChangeType(value, typeof(DateTime));
                Style.NumberFormat.NumberFormatId = dtTest.Date == dtTest ? 14 : 22;

                _cellValue = dtTest.ToOADate().ToString();
            }
            else if (
                value is sbyte
                || value is byte
                || value is char
                || value is short
                || value is ushort
                || value is int
                || value is uint
                || value is long
                || value is ulong
                || value is float
                || value is double
                || value is decimal
                )
            {
                _dataType = XLCellValues.Number;
                _cellValue = value.ToString();
            }
            else if (value is Boolean)
            {
                _dataType = XLCellValues.Boolean;
                _cellValue = (Boolean)Convert.ChangeType(value, typeof(Boolean)) ? "1" : "0";
            }
            else
            {
                _cellValue = value.ToString();
                _dataType = XLCellValues.Text;
            }

            return this;
        }

        public T GetValue<T>()
        {
            if (!StringExtensions.IsNullOrWhiteSpace(FormulaA1))
                return (T)Convert.ChangeType(String.Empty, typeof(T));
            if (Value is TimeSpan)
            {
                if (typeof(T) == typeof(String))
                    return (T)Convert.ChangeType(Value.ToString(), typeof(T));

                return (T)Value;
            }

            if (Value is IXLRichText)
                return (T)RichText;
            return (T)Convert.ChangeType(Value, typeof(T));
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
            string cValue = FormulaA1.Length > 0 ? String.Empty : _cellValue;

            if (_dataType == XLCellValues.Boolean)
                return (cValue != "0").ToString();
            if (_dataType == XLCellValues.TimeSpan)
                return cValue;
            if (_dataType == XLCellValues.DateTime || IsDateFormat())
            {
                double dTest;
                if (Double.TryParse(cValue, out dTest))
                {
                    string format = GetFormat();
                    return DateTime.FromOADate(dTest).ToString(format);
                }

                return cValue;
            }

            if (_dataType == XLCellValues.Number)
            {
                double dTest;
                if (Double.TryParse(cValue, out dTest))
                {
                    string format = GetFormat();
                    return dTest.ToString(format);
                }

                return cValue;
            }

            return cValue;
        }


        public object Value
        {
            get
            {
                string fA1 = FormulaA1;
                if (!StringExtensions.IsNullOrWhiteSpace(fA1))
                {
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

                    if (_worksheet.Workbook.WorksheetsInternal.Any<XLWorksheet>(
                        w => String.Compare(w.Name, sName, true)==0)
                        && ExcelHelper.IsValidA1Address(cAddress)
                        )
                        return _worksheet.Workbook.Worksheet(sName).Cell(cAddress).Value;
                    return fA1;
                }

                if (_dataType == XLCellValues.Boolean)
                    return _cellValue != "0";

                if (_dataType == XLCellValues.DateTime)
                    return DateTime.FromOADate(Double.Parse(_cellValue));

                if (_dataType == XLCellValues.Number)
                    return Double.Parse(_cellValue);

                if (_dataType == XLCellValues.TimeSpan)
                {
                    // return (DateTime.FromOADate(Double.Parse(cellValue)) - baseDate);
                    return TimeSpan.Parse(_cellValue);
                }

                return _richText == null ? _cellValue : _richText.ToString();
            }

            set
            {
                FormulaA1 = String.Empty;
                if (SetEnumerable(value)) return;

                if (SetRange(value)) return;

                if (!SetRichText(value))
                    SetValue(value);
            }
        }

        public IXLTable InsertTable(IEnumerable data)
        {
            return InsertTable(data, null, true);
        }

        public IXLTable InsertTable(IEnumerable data, bool createTable)
        {
            return InsertTable(data, null, createTable);
        }

        public IXLTable InsertTable(IEnumerable data, string tableName)
        {
            return InsertTable(data, tableName, true);
        }

        public IXLTable InsertTable(IEnumerable data, string tableName, bool createTable)
        {
            if (data != null && data.GetType() != typeof(String))
            {
                int ro = Address.RowNumber + 1;
                int fRo = Address.RowNumber;
                bool hasTitles = false;
                int maxCo = 0;
                bool isDataTable = false;
                foreach (object m in data)
                {
                    int co = Address.ColumnNumber;

                    if (m.GetType().IsPrimitive || m is string || m is DateTime)
                    {
                        if (!hasTitles)
                        {
                            string fieldName = GetFieldName(m.GetType().GetCustomAttributes(true));
                            if (StringExtensions.IsNullOrWhiteSpace(fieldName))
                                fieldName = m.GetType().Name;

                            SetValue(fieldName, fRo, co);
                            hasTitles = true;
                            co = Address.ColumnNumber;
                        }

                        SetValue(m, ro, co);
                        co++;
                    }
                    else if (m.GetType().IsArray)
                    {
                        foreach (object item in (Array)m)
                        {
                            SetValue(item, ro, co);
                            co++;
                        }
                    }
                    else if (isDataTable || (m as DataRow) != null)
                    {
                        if (!isDataTable)
                            isDataTable = true;
                        if (!hasTitles)
                        {
                            foreach (string fieldName in from DataColumn column in ((DataRow)m).Table.Columns select StringExtensions.IsNullOrWhiteSpace(column.Caption)
                                                                                                                         ? column.ColumnName
                                                                                                                         : column.Caption)
                            {
                                SetValue(fieldName, fRo, co);
                                co++;
                            }

                            co = Address.ColumnNumber;
                            hasTitles = true;
                        }

                        foreach (object item in ((DataRow)m).ItemArray)
                        {
                            SetValue(item, ro, co);
                            co++;
                        }
                    }
                    else
                    {
                        var fieldInfo = m.GetType().GetFields();
                        var propertyInfo = m.GetType().GetProperties();
                        if (!hasTitles)
                        {
                            foreach (FieldInfo info in fieldInfo)
                            {
                                if ((info as IEnumerable) == null)
                                {
                                    string fieldName = GetFieldName(info.GetCustomAttributes(true));
                                    if (StringExtensions.IsNullOrWhiteSpace(fieldName))
                                        fieldName = info.Name;

                                    SetValue(fieldName, fRo, co);
                                }

                                co++;
                            }

                            foreach (PropertyInfo info in propertyInfo)
                            {
                                if ((info as IEnumerable) == null)
                                {
                                    string fieldName = GetFieldName(info.GetCustomAttributes(true));
                                    if (StringExtensions.IsNullOrWhiteSpace(fieldName))
                                        fieldName = info.Name;

                                    SetValue(fieldName, fRo, co);
                                }

                                co++;
                            }

                            co = Address.ColumnNumber;
                            hasTitles = true;
                        }

                        foreach (FieldInfo info in fieldInfo)
                        {
                            SetValue(info.GetValue(m), ro, co);
                            co++;
                        }

                        foreach (PropertyInfo info in propertyInfo)
                        {
                            if ((info as IEnumerable) == null)
                                SetValue(info.GetValue(m, null), ro, co);
                            co++;
                        }
                    }

                    if (co > maxCo)
                        maxCo = co;

                    ro++;
                }

                ClearMerged();
                var range = _worksheet.Range(
                    Address.RowNumber,
                    Address.ColumnNumber,
                    ro - 1,
                    maxCo - 1);

                if (createTable)
                    return tableName == null ? range.CreateTable() : range.CreateTable(tableName);
                return tableName == null ? range.AsTable() : range.AsTable(tableName);
            }

            return null;
        }

        public IXLRange InsertData(IEnumerable data)
        {
            if (data != null && data.GetType() != typeof(String))
            {
                int ro = Address.RowNumber;
                int maxCo = 0;
                foreach (object m in data)
                {
                    int co = Address.ColumnNumber;

                    if (m.GetType().IsPrimitive || m is string || m is DateTime)
                    {
                        SetValue(m, ro, co);
                        co++;
                    }
                    else if (m.GetType().IsArray)
                    {
                        // dynamic arr = m;
                        foreach (object item in (Array)m)
                        {
                            SetValue(item, ro, co);
                            co++;
                        }
                    }
                    else if ((m as DataRow) != null)
                    {
                        foreach (object item in (m as DataRow).ItemArray)
                        {
                            SetValue(item, ro, co);
                            co++;
                        }
                    }
                    else
                    {
                        var fieldInfo = m.GetType().GetFields();
                        foreach (FieldInfo info in fieldInfo)
                        {
                            SetValue(info.GetValue(m), ro, co);
                            co++;
                        }

                        var propertyInfo = m.GetType().GetProperties();
                        foreach (PropertyInfo info in propertyInfo)
                        {
                            if ((info as IEnumerable) == null)
                                SetValue(info.GetValue(m, null), ro, co);
                            co++;
                        }
                    }

                    if (co > maxCo)
                        maxCo = co;

                    ro++;
                }

                ClearMerged();
                return _worksheet.Range(
                    Address.RowNumber,
                    Address.ColumnNumber,
                    ro - 1,
                    maxCo - 1);
            }

            return null;
        }

        private IXLStyle GetStyle()
        {
            //return _style ?? (_style = new XLStyle(this, Worksheet.Workbook.GetStyleById(_styleCacheId)));
            if (_style != null)
                return _style;

            return _style = new XLStyle(this, Worksheet.Workbook.GetStyleById(_styleCacheId));
        }
        public IXLStyle Style
        {
            get { return GetStyle(); }

            set
            {
                SetStyle(value);
            }
        }

        public IXLCell SetDataType(XLCellValues dataType)
        {
            DataType = dataType;
            return this;
        }


        public XLCellValues DataType
        {
            get { return _dataType; }
            set
            {
                if (_dataType == value) return;

                if (_richText != null)
                {
                    _cellValue = _richText.ToString();
                    _richText = null;
                }

                if (_cellValue.Length > 0)
                {
                    if (value == XLCellValues.Boolean)
                    {
                        bool bTest;
                        if (Boolean.TryParse(_cellValue, out bTest))
                            _cellValue = bTest ? "1" : "0";
                        else
                            _cellValue = _cellValue == "0" || String.IsNullOrEmpty(_cellValue) ? "0" : "1";
                    }
                    else if (value == XLCellValues.DateTime)
                    {
                        DateTime dtTest;
                        double dblTest;
                        if (DateTime.TryParse(_cellValue, out dtTest))
                            _cellValue = dtTest.ToOADate().ToString();
                        else if (Double.TryParse(_cellValue, out dblTest))
                            _cellValue = dblTest.ToString();
                        else
                        {
                            throw new ArgumentException(
                                string.Format(
                                    "Cannot set data type to DateTime because '{0}' is not recognized as a date.",
                                    _cellValue));
                        }
                        var style = GetStyleForRead();
                        if (style.NumberFormat.Format == String.Empty && style.NumberFormat.NumberFormatId == 0)
                        {
                            Style.NumberFormat.NumberFormatId = _cellValue.Contains('.') ? 22 : 14;
                        }
                    }
                    else if (value == XLCellValues.TimeSpan)
                    {
                        TimeSpan tsTest;
                        if (TimeSpan.TryParse(_cellValue, out tsTest))
                        {
                            _cellValue = tsTest.ToString();
                            var style = GetStyleForRead();
                            if (style.NumberFormat.Format == String.Empty && style.NumberFormat.NumberFormatId == 0)
                                Style.NumberFormat.NumberFormatId = 46;
                        }
                        else
                        {
                            try
                            {
                                _cellValue = (DateTime.FromOADate(Double.Parse(_cellValue)) - BaseDate).ToString();
                            }
                            catch
                            {
                                throw new ArgumentException(
                                    string.Format(
                                        "Cannot set data type to TimeSpan because '{0}' is not recognized as a TimeSpan.",
                                        _cellValue));
                            }
                        }
                    }
                    else if (value == XLCellValues.Number)
                    {
                        double dTest;
                        if (Double.TryParse(_cellValue, out dTest))
                            _cellValue = Double.Parse(_cellValue).ToString();
                        else
                        {
                            throw new ArgumentException(
                                string.Format(
                                    "Cannot set data type to Number because '{0}' is not recognized as a number.",
                                    _cellValue));
                        }
                    }
                    else
                    {
                        if (_dataType == XLCellValues.Boolean)
                            _cellValue = (_cellValue != "0").ToString();
                        else if (_dataType == XLCellValues.TimeSpan)
                            _cellValue = BaseDate.Add(GetTimeSpan()).ToOADate().ToString(CultureInfo.InvariantCulture);
                    }
                }

                _dataType = value;
            }
        }

        public IXLCell Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
        {
            if (clearOptions == XLClearOptions.Contents || clearOptions == XLClearOptions.ContentsAndFormats)
            {
                Hyperlink = null;
                _richText = null;
                _cellValue = String.Empty;
                FormulaA1 = String.Empty;
            }

            if (clearOptions == XLClearOptions.Formats || clearOptions == XLClearOptions.ContentsAndFormats)
            {
                DataValidation.Clear();
                SetStyle(Worksheet.Style);
            }
            
            return this;
        }


        public void Delete(XLShiftDeletedCells shiftDeleteCells)
        {
            _worksheet.Range(Address, Address).Delete(shiftDeleteCells);
        }

        public string FormulaA1
        {
            get
            {
                if (StringExtensions.IsNullOrWhiteSpace(_formulaA1))
                {
                    if (!StringExtensions.IsNullOrWhiteSpace(_formulaR1C1))
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
                _formulaA1 = value;
                _formulaR1C1 = String.Empty;
            }
        }

        public string FormulaR1C1
        {
            get
            {
                if (StringExtensions.IsNullOrWhiteSpace(_formulaR1C1))
                    _formulaR1C1 = GetFormulaR1C1(FormulaA1);

                return _formulaR1C1;
            }

            set
            {
                _formulaR1C1 = value;

// FormulaA1 = GetFormulaA1(value);
            }
        }

        public bool ShareString { get; set; }

        public XLHyperlink Hyperlink
        {
            get
            {
                if (_hyperlink == null)
                    Hyperlink = new XLHyperlink();

                return _hyperlink;
            }

            set
            {
                if (_worksheet.Hyperlinks.Any(hl => Address.Equals(hl.Cell.Address)))
                    _worksheet.Hyperlinks.Delete(Address);

                _hyperlink = value;

                if (_hyperlink == null) return;

                _hyperlink.Worksheet = _worksheet;
                _hyperlink.Cell = this;

                _worksheet.Hyperlinks.Add(_hyperlink);

                if (SettingHyperlink) return;

                if (GetStyleForRead().Font.FontColor.Equals(_worksheet.Style.Font.FontColor))
                    Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);

                if (GetStyleForRead().Font.Underline == _worksheet.Style.Font.Underline)
                    Style.Font.Underline = XLFontUnderlineValues.Single;
            }
        }

        public XLDataValidation DataValidation
        {
            get { return AsRange().DataValidation; }
        }
        IXLDataValidation IXLCell.DataValidation
        {
            get { return DataValidation; }
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

        public IXLCell CopyTo(IXLCell target)
        {
            target.Value = this;
            return target;
        }

        public string ValueCached { get; internal set; }

        public IXLRichText RichText
        {
            get
            {
                if (_richText == null)
                {
                    var style = GetStyleForRead();
                    _richText = _cellValue.Length == 0 ? new XLRichText(style.Font) : new XLRichText(GetFormattedString(), style.Font);

                    _dataType = XLCellValues.Text;
                }

                return _richText;
            }
        }

        public bool HasRichText
        {
            get { return _richText != null; }
        }

        public Boolean IsMerged()
        {
            return Worksheet.Internals.MergedRanges.Any(AsRange().Intersects);
        }

        public Boolean IsEmpty()
        {
            return IsEmpty(false);
        }

        public Boolean IsEmpty(Boolean includeFormats)
        {
            if (InnerText.Length > 0)
                return false;

            if (includeFormats)
                return Style.Equals(Worksheet.Style) && !IsMerged();

            return true;
        }

        #endregion

        #region IXLStylized Members

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                yield return Style;
                UpdatingStyle = false;
            }
        }

        public bool UpdatingStyle { get; set; }

        public IXLStyle InnerStyle
        {
            get { return Style; }
            set { Style = value; }
        }

        public IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges {AsRange()};
                return retVal;
            }
        }

        #endregion

        private bool IsDateFormat()
        {
            var style = GetStyleForRead();
            return _dataType == XLCellValues.Number
                   && StringExtensions.IsNullOrWhiteSpace(style.NumberFormat.Format)
                   && ((style.NumberFormat.NumberFormatId >= 14
                        && style.NumberFormat.NumberFormatId <= 22)
                       || (style.NumberFormat.NumberFormatId >= 45
                           && style.NumberFormat.NumberFormatId <= 47));
        }

        private string GetFormat()
        {
            string format;
            var style = GetStyleForRead();
            if (StringExtensions.IsNullOrWhiteSpace(style.NumberFormat.Format))
            {
                var formatCodes = GetFormatCodes();
                format = formatCodes[style.NumberFormat.NumberFormatId];
            }
            else
                format = style.NumberFormat.Format;
            return format;
        }

        private bool SetRichText(object value)
        {
            var asRichString = value as XLRichText;

            if (asRichString == null)
                return false;

            _richText = asRichString;
            _dataType = XLCellValues.Text;
            return true;
        }

        private Boolean SetRange(Object rangeObject)
        {
            var asRange = rangeObject as XLRangeBase;
            if (asRange == null)
            {
                var tmp = rangeObject as XLCell;
                if (tmp != null)
                    asRange = tmp.AsRange();
            }

            if (asRange != null)
            {
                
                if (!(asRange is XLRow || asRange is XLColumn))
                {
                    Int32 maxRows = asRange.RowCount();
                    Int32 maxColumns = asRange.ColumnCount();
                    Worksheet.Range(Address.RowNumber, Address.ColumnNumber, maxRows, maxColumns).Clear();
                }

                Int32 minRow = asRange.RangeAddress.FirstAddress.RowNumber;
                Int32 minColumn = asRange.RangeAddress.FirstAddress.ColumnNumber;
                foreach (XLCell sourceCell in asRange.CellsUsed(true))
                {
                    Worksheet.Cell(
                        Address.RowNumber + sourceCell.Address.RowNumber - minRow,
                        Address.ColumnNumber + sourceCell.Address.ColumnNumber - minColumn
                        ).CopyFrom(sourceCell);
                }

                var rangesToMerge = (from mergedRange in (asRange.Worksheet).Internals.MergedRanges
                                     where asRange.Contains(mergedRange)
                                     let initialRo = Address.RowNumber + (mergedRange.RangeAddress.FirstAddress.RowNumber - asRange.RangeAddress.FirstAddress.RowNumber)
                                     let initialCo = Address.ColumnNumber + (mergedRange.RangeAddress.FirstAddress.ColumnNumber - asRange.RangeAddress.FirstAddress.ColumnNumber)
                                     select Worksheet.Range(initialRo, initialCo, initialRo + mergedRange.RowCount() - 1, initialCo + mergedRange.ColumnCount() - 1)).Cast<IXLRange>().ToList();
                rangesToMerge.ForEach(r => r.Merge());

                return true;
            }
            
                return false;
        }

        private bool SetEnumerable(object collectionObject)
        {
            var asEnumerable = collectionObject as IEnumerable;
            return InsertData(asEnumerable) != null;
        }

        private void ClearMerged()
        {
            var mergeToDelete = Worksheet.Internals.MergedRanges.Where(merge => merge.Intersects(AsRange())).ToList();
            mergeToDelete.ForEach(m => Worksheet.Internals.MergedRanges.Remove(m));
        }

        private void SetValue<T>(T value, int ro, int co) where T: class
        {
            if (value == null)
                _worksheet.Cell(ro, co).SetValue(String.Empty);
            else
                _worksheet.Cell(ro, co).SetValue((T)Convert.ChangeType(value, typeof(T)));
        }

        private void SetValue(object value)
        {
            FormulaA1 = String.Empty;
            string val = value == null ? String.Empty : value.ToString();
            _richText = null;
            if (val.Length == 0)
            {
                _dataType = XLCellValues.Text;
            }
            else
            {
                double dTest;
                DateTime dtTest;
                bool bTest;
                TimeSpan tsTest;
                var style = GetStyleForRead();
                if (style.NumberFormat.Format == "@")
                {
                    _dataType = XLCellValues.Text;
                    if (val.Contains(Environment.NewLine) && !style.Alignment.WrapText)
                        Style.Alignment.WrapText = true;
                }
                else if (val[0] == '\'')
                {
                    val = val.Substring(1, val.Length - 1);
                    _dataType = XLCellValues.Text;
                    if (val.Contains(Environment.NewLine) && !style.Alignment.WrapText)
                        Style.Alignment.WrapText = true;
                }
                else if (value is TimeSpan || (TimeSpan.TryParse(val, out tsTest) && !Double.TryParse(val, out dTest)))
                {
                    _dataType = XLCellValues.TimeSpan;
                    if (style.NumberFormat.Format == String.Empty && style.NumberFormat.NumberFormatId == 0)
                        Style.NumberFormat.NumberFormatId = 46;
                }
                else if (Double.TryParse(val, out dTest))
                    _dataType = XLCellValues.Number;
                else if (DateTime.TryParse(val, out dtTest) && dtTest >= BaseDate)
                {
                    _dataType = XLCellValues.DateTime;

                    if (style.NumberFormat.Format == String.Empty && style.NumberFormat.NumberFormatId == 0)
                    {
                        Style.NumberFormat.NumberFormatId = dtTest.Date == dtTest ? 14 : 22;
                    }

                    val = dtTest.ToOADate().ToString();
                }
                else if (Boolean.TryParse(val, out bTest))
                {
                    _dataType = XLCellValues.Boolean;
                    val = bTest ? "1" : "0";
                }
                else
                {
                    _dataType = XLCellValues.Text;
                    if (val.Contains(Environment.NewLine) && !style.Alignment.WrapText)
                        Style.Alignment.WrapText = true;
                }
            }

            _cellValue = val;
        }

        private static Dictionary<int, string> GetFormatCodes()
        {
            if (_formatCodes == null)
            {
                var fCodes = new Dictionary<int, string>
                                 {
                                     {0, string.Empty},
                                     {1, "0"},
                                     {2, "0.00"},
                                     {3, "#,##0"},
                                     {4, "#,##0.00"},
                                     {9, "0%"},
                                     {10, "0.00%"},
                                     {11, "0.00E+00"},
                                     {12, "# ?/?"},
                                     {13, "# ??/??"},
                                     {14, "MM-dd-yy"},
                                     {15, "d-MMM-yy"},
                                     {16, "d-MMM"},
                                     {17, "MMM-yy"},
                                     {18, "h:mm AM/PM"},
                                     {19, "h:mm:ss AM/PM"},
                                     {20, "h:mm"},
                                     {21, "h:mm:ss"},
                                     {22, "M/d/yy h:mm"},
                                     {37, "#,##0 ;(#,##0)"},
                                     {38, "#,##0 ;[Red](#,##0)"},
                                     {39, "#,##0.00;(#,##0.00)"},
                                     {40, "#,##0.00;[Red](#,##0.00)"},
                                     {45, "mm:ss"},
                                     {46, "[h]:mm:ss"},
                                     {47, "mmss.0"},
                                     {48, "##0.0E+0"},
                                     {49, "@"}
                                 };
                _formatCodes = fCodes;
            }

            return _formatCodes;
        }

        private string GetFormulaR1C1(string value)
        {
            return GetFormula(value, FormulaConversionType.A1ToR1C1, 0, 0);
        }

        private string GetFormulaA1(string value)
        {
            return GetFormula(value, FormulaConversionType.R1C1ToA1, 0, 0);
        }

        private string GetFormula(string strValue, FormulaConversionType conversionType, int rowsToShift,
                                  int columnsToShift)
        {
            if (StringExtensions.IsNullOrWhiteSpace(strValue))
                return String.Empty;

            string value = ">" + strValue + "<";

            var regex = conversionType == FormulaConversionType.A1ToR1C1 ? A1Regex : R1C1Regex;

            var sb = new StringBuilder();
            int lastIndex = 0;

            foreach (Match match in regex.Matches(value).Cast<Match>())
            {
                string matchString = match.Value;
                int matchIndex = match.Index;
                if (value.Substring(0, matchIndex).CharCount('"') % 2 == 0)
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

            string retVal = sb.ToString();
            return retVal.Substring(1, retVal.Length - 2);
        }

        private string GetA1Address(string r1C1Address, int rowsToShift, int columnsToShift)
        {
            string addressToUse = r1C1Address.ToUpper();

            if (addressToUse.Contains(':'))
            {
                var parts = addressToUse.Split(':');
                string p1 = parts[0];
                string p2 = parts[1];
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

            string rowPart = addressToUse.Substring(0, addressToUse.IndexOf("C"));
            string rowToReturn = GetA1Row(rowPart, rowsToShift);

            string columnPart = addressToUse.Substring(addressToUse.IndexOf("C"));
            string columnToReturn = GetA1Column(columnPart, columnsToShift);

            string retAddress = columnToReturn + rowToReturn;
            return retAddress;
        }

        private string GetA1Column(string columnPart, int columnsToShift)
        {
            string columnToReturn;
            if (columnPart == "C")
                columnToReturn = ExcelHelper.GetColumnLetterFromNumber(Address.ColumnNumber + columnsToShift);
            else
            {
                int bIndex = columnPart.IndexOf("[");
                int mIndex = columnPart.IndexOf("-");
                if (bIndex >= 0)
                {
                    columnToReturn = ExcelHelper.GetColumnLetterFromNumber(
                        Address.ColumnNumber +
                        Int32.Parse(columnPart.Substring(bIndex + 1, columnPart.Length - bIndex - 2)) + columnsToShift
                        );
                }
                else if (mIndex >= 0)
                {
                    columnToReturn = ExcelHelper.GetColumnLetterFromNumber(
                        Address.ColumnNumber + Int32.Parse(columnPart.Substring(mIndex)) + columnsToShift
                        );
                }
                else
                {
                    columnToReturn = "$" +
                                     ExcelHelper.GetColumnLetterFromNumber(Int32.Parse(columnPart.Substring(1)) +
                                                                           columnsToShift);
                }
            }

            return columnToReturn;
        }

        private string GetA1Row(string rowPart, int rowsToShift)
        {
            string rowToReturn;
            if (rowPart == "R")
                rowToReturn = (Address.RowNumber + rowsToShift).ToString();
            else
            {
                int bIndex = rowPart.IndexOf("[");
                if (bIndex >= 0)
                {
                    rowToReturn =
                        (Address.RowNumber + Int32.Parse(rowPart.Substring(bIndex + 1, rowPart.Length - bIndex - 2)) +
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
                string p1 = parts[0];
                string p2 = parts[1];
                int row1;
                if (Int32.TryParse(p1.Replace("$", string.Empty), out row1))
                {
                    int row2 = Int32.Parse(p2.Replace("$", string.Empty));
                    string leftPart = GetR1C1Row(row1, p1.Contains('$'), rowsToShift);
                    string rightPart = GetR1C1Row(row2, p2.Contains('$'), rowsToShift);
                    return leftPart + ":" + rightPart;
                }
                else
                {
                    int column1 = ExcelHelper.GetColumnNumberFromLetter(p1.Replace("$", string.Empty));
                    int column2 = ExcelHelper.GetColumnNumberFromLetter(p2.Replace("$", string.Empty));
                    string leftPart = GetR1C1Column(column1, p1.Contains('$'), columnsToShift);
                    string rightPart = GetR1C1Column(column2, p2.Contains('$'), columnsToShift);
                    return leftPart + ":" + rightPart;
                }
            }

            var address = XLAddress.Create(_worksheet, a1Address);

            string rowPart = GetR1C1Row(address.RowNumber, address.FixedRow, rowsToShift);
            string columnPart = GetR1C1Column(address.ColumnNumber, address.FixedColumn, columnsToShift);

            return rowPart + columnPart;
        }

        private string GetR1C1Row(int rowNumber, bool fixedRow, int rowsToShift)
        {
            string rowPart;
            rowNumber += rowsToShift;
            int rowDiff = rowNumber - Address.RowNumber;
            if (rowDiff != 0 || fixedRow)
            {
                rowPart = fixedRow ? String.Format("R{0}", rowNumber) : String.Format("R[{0}]", rowDiff);
            }
            else
                rowPart = "R";

            return rowPart;
        }

        private string GetR1C1Column(int columnNumber, bool fixedColumn, int columnsToShift)
        {
            string columnPart;
            columnNumber += columnsToShift;
            int columnDiff = columnNumber - Address.ColumnNumber;
            if (columnDiff != 0 || fixedColumn)
            {
                columnPart = fixedColumn ? String.Format("C{0}", columnNumber) : String.Format("C[{0}]", columnDiff);
            }
            else
                columnPart = "C";

            return columnPart;
        }

        internal void CopyValues(XLCell source)
        {
            _cellValue = source._cellValue;
            _dataType = source._dataType;
            FormulaR1C1 = source.FormulaR1C1;
            _richText = source._richText == null ? null : new XLRichText(source._richText, source.Style.Font);
        }

        public IXLCell CopyFrom(XLCell otherCell)
        {
            var source = otherCell;
            _cellValue = source._cellValue;
            _richText = source._richText == null ? null : new XLRichText(source._richText, source.Style.Font);

            _dataType = source._dataType;
            FormulaR1C1 = source.FormulaR1C1;
            SetStyle(source._style ?? source.Worksheet.Workbook.GetStyleById(source._styleCacheId));

            if (source._hyperlink != null)
            {
                SettingHyperlink = true;
                Hyperlink = new XLHyperlink(source.Hyperlink);
                SettingHyperlink = false;
            }

            var asRange = source.AsRange();
            if (source.Worksheet.DataValidations.Any(dv => dv.Ranges.Contains(asRange)))
                DataValidation.CopyFrom(source.DataValidation);

            return this;
        }


        internal void ShiftFormulaRows(XLRange shiftedRange, int rowsShifted)
        {
            if (StringExtensions.IsNullOrWhiteSpace(FormulaA1)) return;

            string value = ">" + _formulaA1 + "<";

            var regex = A1SimpleRegex;

            var sb = new StringBuilder();
            int lastIndex = 0;

            String shiftedWsName = shiftedRange.Worksheet.Name;
            foreach (Match match in regex.Matches(value).Cast<Match>())
            {
                string matchString = match.Value;
                int matchIndex = match.Index;
                if (value.Substring(0, matchIndex).CharCount('"') % 2 == 0)
                {
// Check that the match is not between quotes
                    sb.Append(value.Substring(lastIndex, matchIndex - lastIndex));
                    string sheetName;
                    bool useSheetName = false;
                    if (matchString.Contains('!'))
                    {
                        sheetName = matchString.Substring(0, matchString.IndexOf('!'));
                        if (sheetName[0] == '\'')
                            sheetName = sheetName.Substring(1, sheetName.Length - 2);
                        useSheetName = true;
                    }
                    else
                        sheetName = _worksheet.Name;

                    if (String.Compare(sheetName, shiftedWsName, true)==0)
                    {
                        string rangeAddress = matchString.Substring(matchString.IndexOf('!') + 1);
                        if (!A1ColumnRegex.IsMatch(rangeAddress))
                        {
                            var matchRange = _worksheet.Workbook.Worksheet(sheetName).Range(rangeAddress);
                            if (shiftedRange.RangeAddress.FirstAddress.RowNumber <=
                                matchRange.RangeAddress.LastAddress.RowNumber
                                &&
                                shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                                matchRange.RangeAddress.FirstAddress.ColumnNumber
                                &&
                                shiftedRange.RangeAddress.LastAddress.ColumnNumber >=
                                matchRange.RangeAddress.LastAddress.ColumnNumber)
                            {
                                if (A1RowRegex.IsMatch(rangeAddress))
                                {
                                    var rows = rangeAddress.Split(':');
                                    string row1String = rows[0];
                                    string row2String = rows[1];
                                    string row1;
                                    if (row1String[0] == '$')
                                    {
                                        row1 = "$" +
                                               (Int32.Parse(row1String.Substring(1)) + rowsShifted).ToStringLookup();
                                    }
                                    else
                                        row1 = (Int32.Parse(row1String) + rowsShifted).ToStringLookup();

                                    string row2;
                                    if (row2String[0] == '$')
                                    {
                                        row2 = "$" +
                                               (Int32.Parse(row2String.Substring(1)) + rowsShifted).ToStringLookup();
                                    }
                                    else
                                        row2 = (Int32.Parse(row2String) + rowsShifted).ToStringLookup();

                                    sb.Append(useSheetName
                                                  ? String.Format("'{0}'!{1}:{2}", sheetName, row1, row2)
                                                  : String.Format("{0}:{1}", row1, row2));
                                }
                                else if (shiftedRange.RangeAddress.FirstAddress.RowNumber <=
                                         matchRange.RangeAddress.FirstAddress.RowNumber)
                                {
                                    if (rangeAddress.Contains(':'))
                                    {
                                        if (useSheetName)
                                        {
                                            sb.Append(String.Format("'{0}'!{1}:{2}",
                                                                    sheetName,
                                                                    new XLAddress(_worksheet,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.RowNumber +
                                                                                  rowsShifted,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.ColumnLetter,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedColumn),
                                                                    new XLAddress(_worksheet,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.RowNumber +
                                                                                  rowsShifted,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.ColumnLetter,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.FixedColumn)));
                                        }
                                        else
                                        {
                                            sb.Append(String.Format("{0}:{1}",
                                                                    new XLAddress(_worksheet,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.RowNumber +
                                                                                  rowsShifted,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.ColumnLetter,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedColumn),
                                                                    new XLAddress(_worksheet,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.RowNumber +
                                                                                  rowsShifted,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.ColumnLetter,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.FixedColumn)));
                                        }
                                    }
                                    else
                                    {
                                        if (useSheetName)
                                        {
                                            sb.Append(String.Format("'{0}'!{1}",
                                                                    sheetName,
                                                                    new XLAddress(_worksheet,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.RowNumber +
                                                                                  rowsShifted,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.ColumnLetter,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedColumn)));
                                        }
                                        else
                                        {
                                            sb.Append(String.Format("{0}",
                                                                    new XLAddress(_worksheet,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.RowNumber +
                                                                                  rowsShifted,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.ColumnLetter,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedColumn)));
                                        }
                                    }
                                }
                                else
                                {
                                    if (useSheetName)
                                    {
                                        sb.Append(String.Format("'{0}'!{1}:{2}",
                                                                sheetName,
                                                                matchRange.RangeAddress.FirstAddress,
                                                                new XLAddress(_worksheet,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.RowNumber +
                                                                              rowsShifted,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.ColumnLetter,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.FixedRow,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.FixedColumn)));
                                    }
                                    else
                                    {
                                        sb.Append(String.Format("{0}:{1}",
                                                                matchRange.RangeAddress.FirstAddress,
                                                                new XLAddress(_worksheet,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.RowNumber +
                                                                              rowsShifted,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.ColumnLetter,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.FixedRow,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.FixedColumn)));
                                    }
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

            string retVal = sb.ToString();
            _formulaA1 = retVal.Substring(1, retVal.Length - 2);
        }

        internal void ShiftFormulaColumns(XLRange shiftedRange, int columnsShifted)
        {
            if (StringExtensions.IsNullOrWhiteSpace(FormulaA1)) return;

            string value = ">" + _formulaA1 + "<";

            var regex = A1SimpleRegex;

            var sb = new StringBuilder();
            int lastIndex = 0;

            foreach (Match match in regex.Matches(value).Cast<Match>())
            {
                string matchString = match.Value;
                int matchIndex = match.Index;
                if (value.Substring(0, matchIndex).CharCount('"') % 2 == 0)
                {
// Check that the match is not between quotes
                    sb.Append(value.Substring(lastIndex, matchIndex - lastIndex));
                    string sheetName;
                    bool useSheetName = false;
                    if (matchString.Contains('!'))
                    {
                        sheetName = matchString.Substring(0, matchString.IndexOf('!'));
                        if (sheetName[0] == '\'')
                            sheetName = sheetName.Substring(1, sheetName.Length - 2);
                        useSheetName = true;
                    }
                    else
                        sheetName = _worksheet.Name;

                    if (String.Compare(sheetName, shiftedRange.Worksheet.Name, true) == 0)
                    {
                        string rangeAddress = matchString.Substring(matchString.IndexOf('!') + 1);
                        if (!A1RowRegex.IsMatch(rangeAddress))
                        {
                            var matchRange = _worksheet.Workbook.Worksheet(sheetName).Range(rangeAddress);
                            if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                                matchRange.RangeAddress.LastAddress.ColumnNumber
                                &&
                                shiftedRange.RangeAddress.FirstAddress.RowNumber <=
                                matchRange.RangeAddress.FirstAddress.RowNumber
                                &&
                                shiftedRange.RangeAddress.LastAddress.RowNumber >=
                                matchRange.RangeAddress.LastAddress.RowNumber)
                            {
                                if (A1ColumnRegex.IsMatch(rangeAddress))
                                {
                                    var columns = rangeAddress.Split(':');
                                    string column1String = columns[0];
                                    string column2String = columns[1];
                                    string column1;
                                    if (column1String[0] == '$')
                                    {
                                        column1 = "$" +
                                                  ExcelHelper.GetColumnLetterFromNumber(
                                                      ExcelHelper.GetColumnNumberFromLetter(
                                                          column1String.Substring(1)) + columnsShifted);
                                    }
                                    else
                                    {
                                        column1 =
                                            ExcelHelper.GetColumnLetterFromNumber(
                                                ExcelHelper.GetColumnNumberFromLetter(column1String) +
                                                columnsShifted);
                                    }

                                    string column2;
                                    if (column2String[0] == '$')
                                    {
                                        column2 = "$" +
                                                  ExcelHelper.GetColumnLetterFromNumber(
                                                      ExcelHelper.GetColumnNumberFromLetter(
                                                          column2String.Substring(1)) + columnsShifted);
                                    }
                                    else
                                    {
                                        column2 =
                                            ExcelHelper.GetColumnLetterFromNumber(
                                                ExcelHelper.GetColumnNumberFromLetter(column2String) +
                                                columnsShifted);
                                    }

                                    sb.Append(useSheetName
                                                  ? String.Format("'{0}'!{1}:{2}", sheetName, column1, column2)
                                                  : String.Format("{0}:{1}", column1, column2));
                                }
                                else if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                                         matchRange.RangeAddress.FirstAddress.ColumnNumber)
                                {
                                    if (rangeAddress.Contains(':'))
                                    {
                                        if (useSheetName)
                                        {
                                            sb.Append(String.Format("'{0}'!{1}:{2}",
                                                                    sheetName,
                                                                    new XLAddress(_worksheet,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.RowNumber,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.ColumnNumber +
                                                                                  columnsShifted,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedColumn),
                                                                    new XLAddress(_worksheet,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.RowNumber,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.ColumnNumber +
                                                                                  columnsShifted,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.FixedColumn)));
                                        }
                                        else
                                        {
                                            sb.Append(String.Format("{0}:{1}",
                                                                    new XLAddress(_worksheet,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.RowNumber,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.ColumnNumber +
                                                                                  columnsShifted,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedColumn),
                                                                    new XLAddress(_worksheet,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.RowNumber,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.ColumnNumber +
                                                                                  columnsShifted,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.FixedColumn)));
                                        }
                                    }
                                    else
                                    {
                                        if (useSheetName)
                                        {
                                            sb.Append(String.Format("'{0}'!{1}",
                                                                    sheetName,
                                                                    new XLAddress(_worksheet,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.RowNumber,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.ColumnNumber +
                                                                                  columnsShifted,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedColumn)));
                                        }
                                        else
                                        {
                                            sb.Append(String.Format("{0}",
                                                                    new XLAddress(_worksheet,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.RowNumber,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.ColumnNumber +
                                                                                  columnsShifted,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedColumn)));
                                        }
                                    }
                                }
                                else
                                {
                                    if (useSheetName)
                                    {
                                        sb.Append(String.Format("'{0}'!{1}:{2}",
                                                                sheetName,
                                                                matchRange.RangeAddress.FirstAddress,
                                                                new XLAddress(_worksheet,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.RowNumber,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.ColumnNumber +
                                                                              columnsShifted,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.FixedRow,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.FixedColumn)));
                                    }
                                    else
                                    {
                                        sb.Append(String.Format("{0}:{1}",
                                                                matchRange.RangeAddress.FirstAddress,
                                                                new XLAddress(_worksheet,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.RowNumber,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.ColumnNumber +
                                                                              columnsShifted,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.FixedRow,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.FixedColumn)));
                                    }
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

            string retVal = sb.ToString();

            _formulaA1 = retVal.Substring(1, retVal.Length - 2);
        }

        // --

        #region Nested type: FormulaConversionType

        private enum FormulaConversionType
        {
            A1ToR1C1,
            R1C1ToA1
        } ;

        #endregion

        #region XLCell Above
        public XLCell CellAbove()
        {
            return CellAbove(1);
        }
        IXLCell IXLCell.CellAbove()
        {
            return CellAbove();
        }
        public XLCell CellAbove(Int32 step)
        {
            return CellShift(step * -1, 0);
        }
        IXLCell IXLCell.CellAbove(Int32 step)
        {
            return CellAbove(step);
        }
        #endregion

        #region XLCell Below
        public XLCell CellBelow()
        {
            return CellBelow(1);
        }
        IXLCell IXLCell.CellBelow()
        {
            return CellBelow();
        }
        public XLCell CellBelow(Int32 step)
        {
            return CellShift(step, 0);
        }
        IXLCell IXLCell.CellBelow(Int32 step)
        {
            return CellBelow(step);
        }
        #endregion

        #region XLCell Left
        public XLCell CellLeft()
        {
            return CellLeft(1);
        }
        IXLCell IXLCell.CellLeft()
        {
            return CellLeft();
        }
        public XLCell CellLeft(Int32 step)
        {
            return CellShift(0, step * -1);
        }
        IXLCell IXLCell.CellLeft(Int32 step)
        {
            return CellLeft(step);
        }
        #endregion

        #region XLCell Right
        public XLCell CellRight()
        {
            return CellRight(1);
        }
        IXLCell IXLCell.CellRight()
        {
            return CellRight();
        }
        public XLCell CellRight(Int32 step)
        {
            return CellShift(0, step);
        }
        IXLCell IXLCell.CellRight(Int32 step)
        {
            return CellRight(step);
        }
        #endregion

        private XLCell CellShift(Int32 rowsToShift, Int32 columnsToShift)
        {
            return Worksheet.Cell(Address.RowNumber + rowsToShift, Address.ColumnNumber + columnsToShift);
        }

        public IXLColumn WorksheetColumn()
        {
            return Worksheet.Column(Address.ColumnNumber);
        }
        public IXLRow WorksheetRow()
        {
            return Worksheet.Row(Address.RowNumber);
        }
    }
}