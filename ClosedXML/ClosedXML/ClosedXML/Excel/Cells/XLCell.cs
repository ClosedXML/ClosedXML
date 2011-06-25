namespace ClosedXML.Excel
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Data;
    using System.Linq;
    using System.Reflection;
    using System.Text;
    using System.Text.RegularExpressions;

    internal partial class XLCell : IXLCell, IXLStylized
    {
        public static readonly DateTime BaseDate = new DateTime(1899, 12, 30);
        private static Dictionary<int, string> _formatCodes;

        private static readonly Regex _a1Regex = new Regex(
            @"(?<=\W)(\$?[a-zA-Z]{1,3}\$?\d{1,7})(?=\W)" // A1
            + @"|(?<=\W)(\d{1,7}:\d{1,7})(?=\W)" // 1:1
            + @"|(?<=\W)([a-zA-Z]{1,3}:[a-zA-Z]{1,3})(?=\W)"); // A:A

        private static readonly Regex _a1SimpleRegex = new Regex(
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

        private static readonly Regex a1RowRegex = new Regex(
            @"(\d{1,7}:\d{1,7})" // 1:1
            );

        private static readonly Regex a1ColumnRegex = new Regex(
            @"([a-zA-Z]{1,3}:[a-zA-Z]{1,3})" // A:A
            );

        private static readonly Regex r1c1Regex = new Regex(
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

        public XLCell(XLWorksheet worksheet, XLAddress address, IXLStyle defaultStyle)
        {
            Address = address;
            ShareString = true;

            if (defaultStyle == null)
                m_style = new XLStyle(this, worksheet.Style);
            else
                m_style = new XLStyle(this, defaultStyle);

            _worksheet = worksheet;
        }

        #endregion

        public bool SettingHyperlink;
        public int SharedStringId;
        private string m_formulaA1;
        private string m_formulaR1C1;
        private IXLStyle m_style;

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
                else if (StringExtensions.IsNullOrWhiteSpace(_cellValue))
                    return FormulaA1;
                else
                    return _cellValue;
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

        public IXLRange AsRange()
        {
            return _worksheet.Range(Address, Address);
        }

        public IXLCell SetValue<T>(T value)
        {
            FormulaA1 = String.Empty;
            _richText = null;
            if (value is String)
            {
                _cellValue = value.ToString();
                _dataType = XLCellValues.Text;
                if (_cellValue.Contains(Environment.NewLine) && !Style.Alignment.WrapText)
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
                if (dtTest.Date == dtTest)
                    Style.NumberFormat.NumberFormatId = 14;
                else
                    Style.NumberFormat.NumberFormatId = 22;

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
                else
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

        public IXLRichText GetRichText()
        {
            return RichText;
        }

        public string GetFormattedString()
        {
            string cValue;
            if (StringExtensions.IsNullOrWhiteSpace(FormulaA1))
                cValue = _cellValue;
            else
                cValue = GetString();

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

                    if (_worksheet.Internals.Workbook.WorksheetsInternal.Any<XLWorksheet>(
                        w => w.Name.ToLower().Equals(sName.ToLower()))
                        && ExcelHelper.IsValidA1Address(cAddress)
                        )
                        return _worksheet.Internals.Workbook.Worksheet(sName).Cell(cAddress).Value;
                    return fA1;
                }

                if (_dataType == XLCellValues.Boolean)
                    return _cellValue != "0";
                else if (_dataType == XLCellValues.DateTime)
                    return DateTime.FromOADate(Double.Parse(_cellValue));
                else if (_dataType == XLCellValues.Number)
                    return Double.Parse(_cellValue);
                else if (_dataType == XLCellValues.TimeSpan)
                {
                    // return (DateTime.FromOADate(Double.Parse(cellValue)) - baseDate);
                    return TimeSpan.Parse(_cellValue);
                }
                else
                {
                    if (_richText == null)
                        return _cellValue;
                    else
                        return _richText.ToString();
                }
            }

            set
            {
                FormulaA1 = String.Empty;
                if (!SetEnumerable(value))
                {
                    if (!SetRange(value))
                    {
                        if (!SetRichText(value))
                            SetValue(value);
                    }
                }
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
                int co;
                int ro = Address.RowNumber + 1;
                int fRo = Address.RowNumber;
                bool hasTitles = false;
                int maxCo = 0;
                bool isDataTable = false;
                foreach (object m in data)
                {
                    co = Address.ColumnNumber;

                    if (m.GetType().IsPrimitive || m.GetType() == typeof(String) || m.GetType() == typeof(DateTime))
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
                            foreach (DataColumn column in (m as DataRow).Table.Columns)
                            {
                                string fieldName = StringExtensions.IsNullOrWhiteSpace(column.Caption)
                                                       ? column.ColumnName
                                                       : column.Caption;
                                SetValue(fieldName, fRo, co);
                                co++;
                            }

                            co = Address.ColumnNumber;
                            hasTitles = true;
                        }

                        foreach (object item in (m as DataRow).ItemArray)
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

                ClearMerged(ro - 1, maxCo - 1);
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

                    if (m.GetType().IsPrimitive || m.GetType() == typeof(String) || m.GetType() == typeof(DateTime))
                        SetValue(m, ro, co);
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

                ClearMerged(ro - 1, maxCo - 1);
                return _worksheet.Range(
                    Address.RowNumber, 
                    Address.ColumnNumber, 
                    Address.RowNumber + ro - 1, 
                    Address.ColumnNumber + maxCo - 1);
            }

            return null;
        }

        public IXLStyle Style
        {
            get
            {
                m_style = new XLStyle(this, m_style);
                return m_style;
            }

            set { m_style = new XLStyle(this, value); }
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
                if (_dataType != value)
                {
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

                            if (Style.NumberFormat.Format == String.Empty && Style.NumberFormat.NumberFormatId == 0)
                            {
                                if (_cellValue.Contains('.'))
                                    Style.NumberFormat.NumberFormatId = 22;
                                else
                                    Style.NumberFormat.NumberFormatId = 14;
                            }
                        }
                        else if (value == XLCellValues.TimeSpan)
                        {
                            TimeSpan tsTest;
                            if (TimeSpan.TryParse(_cellValue, out tsTest))
                            {
                                _cellValue = tsTest.ToString();
                                if (Style.NumberFormat.Format == String.Empty && Style.NumberFormat.NumberFormatId == 0)
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
                            var formatCodes = GetFormatCodes();
                            if (_dataType == XLCellValues.Boolean)
                                _cellValue = (_cellValue != "0").ToString();
                            else if (_dataType == XLCellValues.TimeSpan)
                                _cellValue = TimeSpan.Parse(_cellValue).ToString();
                            else if (_dataType == XLCellValues.Number)
                            {
                                string format;
                                if (Style.NumberFormat.NumberFormatId > 0)
                                    format = formatCodes[Style.NumberFormat.NumberFormatId];
                                else
                                    format = Style.NumberFormat.Format;

                                if (!StringExtensions.IsNullOrWhiteSpace(format) && format != "@")
                                    _cellValue = Double.Parse(_cellValue).ToString(format);
                            }
                            else if (_dataType == XLCellValues.DateTime)
                            {
                                string format;
                                if (Style.NumberFormat.NumberFormatId > 0)
                                    format = formatCodes[Style.NumberFormat.NumberFormatId];
                                else
                                    format = Style.NumberFormat.Format;
                                _cellValue = DateTime.FromOADate(Double.Parse(_cellValue)).ToString(format);
                            }
                        }
                    }

                    _dataType = value;
                }
            }
        }

        public void Clear()
        {
            _worksheet.Range(Address, Address).Clear();
        }

        public void ClearStyles()
        {
            var newStyle = new XLStyle(this, _worksheet.Style);
            newStyle.NumberFormat = Style.NumberFormat;
            Style = newStyle;
        }

        public void Delete(XLShiftDeletedCells shiftDeleteCells)
        {
            _worksheet.Range(Address, Address).Delete(shiftDeleteCells);
        }

        public string FormulaA1
        {
            get
            {
                if (StringExtensions.IsNullOrWhiteSpace(m_formulaA1))
                {
                    if (!StringExtensions.IsNullOrWhiteSpace(m_formulaR1C1))
                    {
                        m_formulaA1 = GetFormulaA1(m_formulaR1C1);
                        return FormulaA1;
                    }
                    else
                        return String.Empty;
                }
                else if (m_formulaA1.Trim()[0] == '=')
                    return m_formulaA1.Substring(1);
                else if (m_formulaA1.Trim().StartsWith("{="))
                    return "{" + m_formulaA1.Substring(2);
                else
                    return m_formulaA1;
            }

            set
            {
                m_formulaA1 = value;
                m_formulaR1C1 = String.Empty;
            }
        }

        public string FormulaR1C1
        {
            get
            {
                if (StringExtensions.IsNullOrWhiteSpace(m_formulaR1C1))
                    m_formulaR1C1 = GetFormulaR1C1(FormulaA1);

                return m_formulaR1C1;
            }

            set
            {
                m_formulaR1C1 = value;

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
                _hyperlink = value;
                _hyperlink.Worksheet = _worksheet;
                _hyperlink.Cell = this;
                if (_worksheet.Hyperlinks.Any(hl => Address.Equals(hl.Cell.Address)))
                    _worksheet.Hyperlinks.Delete(Address);

                _worksheet.Hyperlinks.Add(_hyperlink);

                if (!SettingHyperlink)
                {
                    if (Style.Font.FontColor.Equals(_worksheet.Style.Font.FontColor))
                        Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);

                    if (Style.Font.Underline == _worksheet.Style.Font.Underline)
                        Style.Font.Underline = XLFontUnderlineValues.Single;
                }
            }
        }

        public IXLDataValidation DataValidation
        {
            get { return AsRange().DataValidation; }
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
                    if (StringExtensions.IsNullOrWhiteSpace(_cellValue))
                        _richText = new XLRichText(m_style.Font);
                    else
                        _richText = new XLRichText(GetFormattedString(), m_style.Font);

                    _dataType = XLCellValues.Text;
                    if (!Style.Alignment.WrapText)
                        Style.Alignment.WrapText = true;
                }

                return _richText;
            }
        }

        public bool HasRichText
        {
            get { return _richText != null; }
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
                var retVal = new XLRanges();
                retVal.Add(AsRange());
                return retVal;
            }
        }

        #endregion

        private bool IsDateFormat()
        {
            return _dataType == XLCellValues.Number
                    && StringExtensions.IsNullOrWhiteSpace(Style.NumberFormat.Format)
                    && ((Style.NumberFormat.NumberFormatId >= 14
                         && Style.NumberFormat.NumberFormatId <= 22)
                        || (Style.NumberFormat.NumberFormatId >= 45
                            && Style.NumberFormat.NumberFormatId <= 47));
        }

        private string GetFormat()
        {
            string format;
            if (StringExtensions.IsNullOrWhiteSpace(Style.NumberFormat.Format))
            {
                var formatCodes = GetFormatCodes();
                format = formatCodes[Style.NumberFormat.NumberFormatId];
            }
            else
                format = Style.NumberFormat.Format;
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

        private bool SetRange(object rangeObject)
        {
            var asRange = rangeObject as XLRangeBase;
            if (asRange == null)
            {
                var tmp = rangeObject as XLCell;
                if (tmp != null)
                    asRange = tmp.AsRange() as XLRangeBase;
            }

            if (asRange != null)
            {
                int maxRows;
                int maxColumns;
                if (asRange is XLRow || asRange is XLColumn)
                {
                    var lastCellUsed = asRange.LastCellUsed(true);
                    maxRows = lastCellUsed.Address.RowNumber;
                    maxColumns = lastCellUsed.Address.ColumnNumber;

// if (asRange is XLRow)
                    // {
                    // worksheet.Range(Address.RowNumber, Address.ColumnNumber,  , maxColumns).Clear();
                    // }
                }
                else
                {
                    maxRows = asRange.RowCount();
                    maxColumns = asRange.ColumnCount();
                    _worksheet.Range(Address.RowNumber, Address.ColumnNumber, maxRows, maxColumns).Clear();
                }

                for (int ro = 1; ro <= maxRows; ro++)
                {
                    for (int co = 1; co <= maxColumns; co++)
                    {
                        var sourceCell = asRange.Cell(ro, co);
                        var targetCell = _worksheet.Cell(Address.RowNumber + ro - 1, Address.ColumnNumber + co - 1);
                        targetCell.CopyFrom(sourceCell);

// targetCell.Style = sourceCell.style;
                    }
                }

                var rangesToMerge = new List<IXLRange>();
                foreach (SheetRange mergedRange in asRange.Worksheet.Internals.MergedRanges)
                {
                    if (asRange.Contains(mergedRange))
                    {
                        int initialRo = Address.RowNumber +
                                        (mergedRange.FirstAddress.RowNumber -
                                         asRange.RangeAddress.FirstAddress.RowNumber);
                        int initialCo = Address.ColumnNumber +
                                        (mergedRange.FirstAddress.ColumnNumber -
                                         asRange.RangeAddress.FirstAddress.ColumnNumber);
                        rangesToMerge.Add(_worksheet.Range(initialRo, 
                                                           initialCo, 
                                                           initialRo + mergedRange.RowCount - 1, 
                                                           initialCo + mergedRange.ColumnCount - 1));
                    }
                }

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

        private void ClearMerged(int rowCount, int columnCount)
        {
            // TODO: For MDLeon: Need review why parameters is never used(see compare with revision 67871 before VF changes)
            var intersectingRanges =
                _worksheet.Internals.MergedRanges.GetIntersectingMergedRanges(Address.GetSheetPoint());
            intersectingRanges.ForEach(m => _worksheet.Internals.MergedRanges.Remove(m));
        }

        private void SetValue(object objWithValue, int ro, int co)
        {
            string str = String.Empty;
            if (objWithValue != null)
                str = objWithValue.ToString();

            _worksheet.Cell(ro, co).Value = str;
        }

        private void SetValue(object value)
        {
            FormulaA1 = String.Empty;
            string val = value.ToString();
            _richText = null;
            if (val.Length > 0)
            {
                double dTest;
                DateTime dtTest;
                bool bTest;
                TimeSpan tsTest;
                if (m_style.NumberFormat.Format == "@")
                {
                    _dataType = XLCellValues.Text;
                    if (val.Contains(Environment.NewLine) && !Style.Alignment.WrapText)
                        Style.Alignment.WrapText = true;
                }
                else if (val[0] == '\'')
                {
                    val = val.Substring(1, val.Length - 1);
                    _dataType = XLCellValues.Text;
                    if (val.Contains(Environment.NewLine) && !Style.Alignment.WrapText)
                        Style.Alignment.WrapText = true;
                }
                else if (value is TimeSpan || (TimeSpan.TryParse(val, out tsTest) && !Double.TryParse(val, out dTest)))
                {
                    _dataType = XLCellValues.TimeSpan;
                    if (Style.NumberFormat.Format == String.Empty && Style.NumberFormat.NumberFormatId == 0)
                        Style.NumberFormat.NumberFormatId = 46;
                }
                else if (Double.TryParse(val, out dTest))
                    _dataType = XLCellValues.Number;
                else if (DateTime.TryParse(val, out dtTest))
                {
                    _dataType = XLCellValues.DateTime;

                    if (Style.NumberFormat.Format == String.Empty && Style.NumberFormat.NumberFormatId == 0)
                    {
                        if (dtTest.Date == dtTest)
                            Style.NumberFormat.NumberFormatId = 14;
                        else
                            Style.NumberFormat.NumberFormatId = 22;
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
                    if (val.Contains(Environment.NewLine) && !Style.Alignment.WrapText)
                        Style.Alignment.WrapText = true;
                }
            }

            _cellValue = val;
        }

        private static Dictionary<int, string> GetFormatCodes()
        {
            if (_formatCodes == null)
            {
                var fCodes = new Dictionary<int, string>();
                fCodes.Add(0, string.Empty);
                fCodes.Add(1, "0");
                fCodes.Add(2, "0.00");
                fCodes.Add(3, "#,##0");
                fCodes.Add(4, "#,##0.00");
                fCodes.Add(9, "0%");
                fCodes.Add(10, "0.00%");
                fCodes.Add(11, "0.00E+00");
                fCodes.Add(12, "# ?/?");
                fCodes.Add(13, "# ??/??");
                fCodes.Add(14, "MM-dd-yy");
                fCodes.Add(15, "d-MMM-yy");
                fCodes.Add(16, "d-MMM");
                fCodes.Add(17, "MMM-yy");
                fCodes.Add(18, "h:mm AM/PM");
                fCodes.Add(19, "h:mm:ss AM/PM");
                fCodes.Add(20, "h:mm");
                fCodes.Add(21, "h:mm:ss");
                fCodes.Add(22, "M/d/yy h:mm");
                fCodes.Add(37, "#,##0 ;(#,##0)");
                fCodes.Add(38, "#,##0 ;[Red](#,##0)");
                fCodes.Add(39, "#,##0.00;(#,##0.00)");
                fCodes.Add(40, "#,##0.00;[Red](#,##0.00)");
                fCodes.Add(45, "mm:ss");
                fCodes.Add(46, "[h]:mm:ss");
                fCodes.Add(47, "mmss.0");
                fCodes.Add(48, "##0.0E+0");
                fCodes.Add(49, "@");
                _formatCodes = fCodes;
            }

            return _formatCodes;
        }

        private string GetFormulaR1C1(string value)
        {
            return GetFormula(value, FormulaConversionType.A1toR1C1, 0, 0);
        }

        private string GetFormulaA1(string value)
        {
            return GetFormula(value, FormulaConversionType.R1C1toA1, 0, 0);
        }

        private string GetFormula(string strValue, FormulaConversionType conversionType, int rowsToShift, 
                                  int columnsToShift)
        {
            if (StringExtensions.IsNullOrWhiteSpace(strValue))
                return String.Empty;

            string value = ">" + strValue + "<";

            var regex = conversionType == FormulaConversionType.A1toR1C1 ? _a1Regex : r1c1Regex;

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
                    if (conversionType == FormulaConversionType.A1toR1C1)
                        sb.Append(GetR1C1Address(matchString, rowsToShift, columnsToShift));
                    else
                        sb.Append(GetA1Address(matchString, rowsToShift, columnsToShift));
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

        private string GetA1Address(string r1c1Address, int rowsToShift, int columnsToShift)
        {
            string addressToUse = r1c1Address.ToUpper();

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
            else
            {
                string rowPart = addressToUse.Substring(0, addressToUse.IndexOf("C"));
                string rowToReturn = GetA1Row(rowPart, rowsToShift);

                string columnPart = addressToUse.Substring(addressToUse.IndexOf("C"));
                string columnToReturn = GetA1Column(columnPart, columnsToShift);

                string retAddress = columnToReturn + rowToReturn;
                return retAddress;
            }
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
                    columnToReturn = "$" +
                                     ExcelHelper.GetColumnLetterFromNumber(Int32.Parse(columnPart.Substring(1)) +
                                                                           columnsToShift);
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
            string columnPart = GetR1C1Column(address.ColumnNumber, address.FixedRow, columnsToShift);

            return rowPart + columnPart;
        }

        private string GetR1C1Row(int rowNumber, bool fixedRow, int rowsToShift)
        {
            string rowPart;
            rowNumber += rowsToShift;
            int rowDiff = rowNumber - Address.RowNumber;
            if (rowDiff != 0 || fixedRow)
            {
                if (fixedRow)
                    rowPart = String.Format("R{0}", rowNumber);
                else
                    rowPart = String.Format("R[{0}]", rowDiff);
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
                if (fixedColumn)
                    columnPart = String.Format("C{0}", columnNumber);
                else
                    columnPart = String.Format("C[{0}]", columnDiff);
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
            _richText = new XLRichText(source._richText, source.Style.Font);
        }

        public IXLCell CopyFrom(XLCell otherCell)
        {
            var source = otherCell;
            _cellValue = source._cellValue;
            _richText = new XLRichText(source._richText, source.Style.Font);
            _dataType = source._dataType;
            FormulaR1C1 = source.FormulaR1C1;
            m_style = new XLStyle(this, source.m_style);

            if (source._hyperlink != null)
            {
                SettingHyperlink = true;
                Hyperlink = new XLHyperlink(source.Hyperlink);
                SettingHyperlink = false;
            }

            var asRange = source.AsRange();
            if (source.Worksheet.DataValidations.Any(dv => dv.Ranges.Contains(asRange)))
                (DataValidation as XLDataValidation).CopyFrom(source.DataValidation);

            return this;
        }

        // internal void ShiftFormula(Int32 rowsToShift, Int32 columnsToShift)
        // {
        // if (!StringExtensions.IsNullOrWhiteSpace(formulaA1))
        // FormulaR1C1 = GetFormula(formulaA1, FormulaConversionType.A1toR1C1, rowsToShift, columnsToShift);
        // else if (!StringExtensions.IsNullOrWhiteSpace(formulaR1C1))
        // FormulaA1 = GetFormula(formulaR1C1, FormulaConversionType.R1C1toA1, rowsToShift, columnsToShift);
        // }

        internal void ShiftFormulaRows(XLRange shiftedRange, int rowsShifted)
        {
            if (!StringExtensions.IsNullOrWhiteSpace(FormulaA1))
            {
                string value = ">" + m_formulaA1 + "<";

                var regex = _a1SimpleRegex;

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

                        if (sheetName.ToLower().Equals(shiftedRange.Worksheet.Name.ToLower()))
                        {
                            string rangeAddress = matchString.Substring(matchString.IndexOf('!') + 1);
                            if (!a1ColumnRegex.IsMatch(rangeAddress))
                            {
                                var matchRange = _worksheet.Internals.Workbook.Worksheet(sheetName).Range(rangeAddress);
                                if (shiftedRange.RangeAddress.FirstAddress.RowNumber <=
                                    matchRange.RangeAddress.LastAddress.RowNumber
                                    &&
                                    shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                                    matchRange.RangeAddress.FirstAddress.ColumnNumber
                                    &&
                                    shiftedRange.RangeAddress.LastAddress.ColumnNumber >=
                                    matchRange.RangeAddress.LastAddress.ColumnNumber)
                                {
                                    if (a1RowRegex.IsMatch(rangeAddress))
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

                                        if (useSheetName)
                                            sb.Append(String.Format("'{0}'!{1}:{2}", sheetName, row1, row2));
                                        else
                                            sb.Append(String.Format("{0}:{1}", row1, row2));
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
                m_formulaA1 = retVal.Substring(1, retVal.Length - 2);
            }
        }

        internal void ShiftFormulaColumns(XLRange shiftedRange, int columnsShifted)
        {
            if (!StringExtensions.IsNullOrWhiteSpace(FormulaA1))
            {
                string value = ">" + m_formulaA1 + "<";

                var regex = _a1SimpleRegex;

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

                        if (sheetName.ToLower().Equals(shiftedRange.Worksheet.Name.ToLower()))
                        {
                            string rangeAddress = matchString.Substring(matchString.IndexOf('!') + 1);
                            if (!a1RowRegex.IsMatch(rangeAddress))
                            {
                                var matchRange = _worksheet.Internals.Workbook.Worksheet(sheetName).Range(rangeAddress);
                                if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                                    matchRange.RangeAddress.LastAddress.ColumnNumber
                                    &&
                                    shiftedRange.RangeAddress.FirstAddress.RowNumber <=
                                    matchRange.RangeAddress.FirstAddress.RowNumber
                                    &&
                                    shiftedRange.RangeAddress.LastAddress.RowNumber >=
                                    matchRange.RangeAddress.LastAddress.RowNumber)
                                {
                                    if (a1ColumnRegex.IsMatch(rangeAddress))
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

                                        if (useSheetName)
                                            sb.Append(String.Format("'{0}'!{1}:{2}", sheetName, column1, column2));
                                        else
                                            sb.Append(String.Format("{0}:{1}", column1, column2));
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

                m_formulaA1 = retVal.Substring(1, retVal.Length - 2);
            }
        }

        // --
        #region Nested type: FormulaConversionType

        private enum FormulaConversionType
        {
            A1toR1C1, 
            R1C1toA1
        };

        #endregion
    }
}