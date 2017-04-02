using FastMember;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel
{
    using Attributes;
    using ClosedXML.Extensions;

    internal class XLCell : IXLCell, IXLStylized
    {
        public static readonly DateTime BaseDate = new DateTime(1899, 12, 30);
        private static Dictionary<int, string> _formatCodes;

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
            + @"\$?[a-zA-Z]{1,3}\$?\d{1,7}" // A1 Address 1
            + @"(?<RangeEnd>:\$?[a-zA-Z]{1,3}\$?\d{1,7})?" // A1 Address 2, optional
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

        private readonly XLWorksheet _worksheet;

        internal string _cellValue = String.Empty;

        private XLComment _comment;
        internal XLCellValues _dataType;
        private XLHyperlink _hyperlink;
        private XLRichText _richText;

        #endregion Fields

        #region Constructor

        private Int32 _styleCacheId;

        public XLCell(XLWorksheet worksheet, XLAddress address, Int32 styleId)
        {
            Address = address;
            ShareString = true;
            _worksheet = worksheet;
            SetStyle(styleId);
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

        private void SetStyle(Int32 styleId)
        {
            _styleCacheId = styleId;
            _style = null;
            StyleChanged = false;
        }

        #endregion Constructor

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

                return string.Empty == _cellValue ? FormulaA1 : _cellValue;
            }
        }

        public IXLDataValidation NewDataValidation
        {
            get
            {
                using (var asRange = AsRange())
                {
                    return asRange.NewDataValidation; // Call the data validation without breaking it into pieces
                }
            }
        }

        public XLDataValidation DataValidation
        {
            get
            {
                using (var asRange = AsRange())
                {
                    var dv = asRange.DataValidation; // Call the data validation to break it into pieces
                    foreach (var d in Worksheet.DataValidations)
                    {
                        var rs = d.Ranges;
                        if (rs.Count == 1)
                        {
                            var r = rs.Single();
                            var ra1 = r.RangeAddress.ToStringRelative();
                            var ra2 = asRange.RangeAddress.ToStringRelative();
                            if (ra1.Equals(ra2))
                                return d as XLDataValidation;
                        }
                    }
                }
                return null;
            }
        }

        internal XLComment Comment
        {
            get
            {
                if (_comment == null)
                {
                    // MS Excel uses Tahoma 8 Swiss no matter what current style font
                    // var style = GetStyleForRead();
                    var defaultFont = new XLFont
                    {
                        FontName = "Tahoma",
                        FontSize = 8,
                        FontFamilyNumbering = XLFontFamilyNumberingValues.Swiss
                    };
                    _comment = new XLComment(this, defaultFont);
                }

                return _comment;
            }
        }

        #region IXLCell Members

        IXLDataValidation IXLCell.DataValidation
        {
            get { return DataValidation; }
        }

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
            FormulaA1 = String.Empty;
            _richText = null;
            var style = GetStyleForRead();
            if (value is String || value is char)
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
                if (style.NumberFormat.Format == String.Empty && style.NumberFormat.NumberFormatId == 0)
                    Style.NumberFormat.NumberFormatId = 46;
            }
            else if (value is DateTime)
            {
                _dataType = XLCellValues.DateTime;
                var dtTest = (DateTime)Convert.ChangeType(value, typeof(DateTime));
                if (style.NumberFormat.Format == String.Empty && style.NumberFormat.NumberFormatId == 0)
                    Style.NumberFormat.NumberFormatId = dtTest.Date == dtTest ? 14 : 22;

                _cellValue = dtTest.ToOADate().ToInvariantString();
            }
            else if (value.GetType().IsNumber())
            {
                if ((value is double || value is float) && (Double.IsNaN((Double)Convert.ChangeType(value, typeof(Double)))
                    || Double.IsInfinity((Double)Convert.ChangeType(value, typeof(Double)))))
                {
                    _cellValue = value.ToString();
                    _dataType = XLCellValues.Text;
                }
                else
                {
                    _dataType = XLCellValues.Number;
                    _cellValue = ((Double)Convert.ChangeType(value, typeof(Double))).ToInvariantString();
                }
            }
            else if (value is Boolean)
            {
                _dataType = XLCellValues.Boolean;
                _cellValue = (Boolean)Convert.ChangeType(value, typeof(Boolean)) ? "1" : "0";
            }
            else
            {
                _cellValue = Convert.ToString(value);
                _dataType = XLCellValues.Text;
            }

            return this;
        }

        public T GetValue<T>()
        {
            T retVal;
            if (TryGetValue(out retVal))
                return retVal;

            throw new Exception("Cannot convert cell value to " + typeof(T));
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
            String cValue;
            if (FormulaA1.Length > 0)
            {
                try
                {
                    cValue = GetString();
                }
                catch
                {
                    cValue = String.Empty;
                }
            }
            else
            {
                cValue = _cellValue;
            }

            if (_dataType == XLCellValues.Boolean)
                return (cValue != "0").ToString();
            if (_dataType == XLCellValues.TimeSpan)
                return cValue;
            if (_dataType == XLCellValues.DateTime || IsDateFormat())
            {
                double dTest;
                if (Double.TryParse(cValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out dTest)
                    && dTest.IsValidOADateNumber())
                {
                    var format = GetFormat();
                    return DateTime.FromOADate(dTest).ToString(format);
                }

                return cValue;
            }

            if (_dataType == XLCellValues.Number)
            {
                double dTest;
                if (Double.TryParse(cValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out dTest))
                {
                    var format = GetFormat();
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
                var fA1 = FormulaA1;
                if (!XLHelper.IsNullOrWhiteSpace(fA1))
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
                        w => String.Compare(w.Name, sName, true) == 0)
                        && XLHelper.IsValidA1Address(cAddress)
                        )
                        return _worksheet.Workbook.Worksheet(sName).Cell(cAddress).Value;

                    var retVal = Worksheet.Evaluate(fA1);
                    var retValEnumerable = retVal as IEnumerable;

                    if (retValEnumerable != null && !(retVal is String))
                        foreach (var v in retValEnumerable)
                            return v;

                    return retVal;
                }

                var cellValue = HasRichText ? _richText.ToString() : _cellValue;

                if (_dataType == XLCellValues.Boolean)
                    return cellValue != "0";

                if (_dataType == XLCellValues.DateTime)
                {
                    Double d;
                    if (Double.TryParse(cellValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out d)
                        && d.IsValidOADateNumber())
                        return DateTime.FromOADate(d);
                }

                if (_dataType == XLCellValues.Number)
                {
                    Double d;
                    if (double.TryParse(cellValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out d))
                        return d;
                }

                if (_dataType == XLCellValues.TimeSpan)
                {
                    TimeSpan t;
                    if (TimeSpan.TryParse(cellValue, out t))
                        return t;
                }

                return cellValue;
            }

            set
            {
                FormulaA1 = String.Empty;

                if (value as XLCells != null) throw new ArgumentException("Cannot assign IXLCells object to the cell value.");

                if (SetRangeRows(value)) return;

                if (SetRangeColumns(value)) return;

                if (SetEnumerable(value)) return;

                if (SetRange(value)) return;

                if (!SetRichText(value))
                    SetValue(value);

                if (_cellValue.Length > 32767) throw new ArgumentException("Cells can only hold 32,767 characters.");
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

        public IXLTable InsertTable<T>(IEnumerable<T> data, string tableName, bool createTable)
        {
            if (data != null && data.GetType() != typeof(String))
            {
                var ro = Address.RowNumber + 1;
                var fRo = Address.RowNumber;
                var hasTitles = false;
                var maxCo = 0;
                var isDataTable = false;
                var isDataReader = false;
                var itemType = data.GetItemType();

                if (!data.Any())
                {
                    if (itemType.IsPrimitive || itemType == typeof(String) || itemType == typeof(DateTime) || itemType.IsNumber())
                        maxCo = Address.ColumnNumber + 1;
                    else
                        maxCo = Address.ColumnNumber + itemType.GetFields().Length + itemType.GetProperties().Length;
                }
                else if (itemType.IsPrimitive || itemType == typeof(String) || itemType == typeof(DateTime) || itemType.IsNumber())
                {
                    foreach (object o in data)
                    {
                        var co = Address.ColumnNumber;

                        if (!hasTitles)
                        {
                            var fieldName = XLColumnAttribute.GetHeader(itemType);
                            if (XLHelper.IsNullOrWhiteSpace(fieldName))
                                fieldName = itemType.Name;

                            SetValue(fieldName, fRo, co);
                            hasTitles = true;
                            co = Address.ColumnNumber;
                        }

                        SetValue(o, ro, co);
                        co++;

                        if (co > maxCo)
                            maxCo = co;

                        ro++;
                    }
                }
                else
                {
                    const BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.Instance;
                    var memberCache = new Dictionary<Type, IEnumerable<MemberInfo>>();
                    var accessorCache = new Dictionary<Type, TypeAccessor>();
                    IEnumerable<MemberInfo> members = null;
                    TypeAccessor accessor = null;
                    bool isPlainObject = itemType == typeof(object);

                    if (!isPlainObject)
                    {
                        members = itemType.GetFields(bindingFlags).Cast<MemberInfo>()
                             .Concat(itemType.GetProperties(bindingFlags))
                             .Where(mi => !XLColumnAttribute.IgnoreMember(mi))
                             .OrderBy(mi => XLColumnAttribute.GetOrder(mi));
                        accessor = TypeAccessor.Create(itemType);
                    }

                    foreach (T m in data)
                    {
                        if (isPlainObject)
                        {
                            // In this case data is just IEnumerable<object>, which means we have to determine the runtime type of each element
                            // This is very inefficient and we prefer type of T to be a concrete class or struct
                            var type = m.GetType();
                            if (!memberCache.ContainsKey(type))
                            {
                                var _accessor = TypeAccessor.Create(type);

                                var _members = type.GetFields(bindingFlags).Cast<MemberInfo>()
                                     .Concat(type.GetProperties(bindingFlags))
                                     .Where(mi => !XLColumnAttribute.IgnoreMember(mi))
                                     .OrderBy(mi => XLColumnAttribute.GetOrder(mi));

                                memberCache.Add(type, _members);
                                accessorCache.Add(type, _accessor);
                            }

                            members = memberCache[type];
                            accessor = accessorCache[type];
                        }

                        var co = Address.ColumnNumber;

                        if (itemType.IsArray)
                        {
                            foreach (var item in (m as Array))
                            {
                                SetValue(item, ro, co);
                                co++;
                            }
                        }
                        else if (isDataTable || m is DataRow)
                        {
                            var row = m as DataRow;
                            if (!isDataTable)
                                isDataTable = true;

                            if (!hasTitles)
                            {
                                foreach (var fieldName in from DataColumn column in row.Table.Columns
                                                          select XLHelper.IsNullOrWhiteSpace(column.Caption)
                                                                     ? column.ColumnName
                                                                     : column.Caption)
                                {
                                    SetValue(fieldName, fRo, co);
                                    co++;
                                }

                                co = Address.ColumnNumber;
                                hasTitles = true;
                            }

                            foreach (var item in row.ItemArray)
                            {
                                SetValue(item, ro, co);
                                co++;
                            }
                        }
                        else if (isDataReader || m is IDataRecord)
                        {
                            if (!isDataReader)
                                isDataReader = true;

                            var record = m as IDataRecord;

                            var fieldCount = record.FieldCount;
                            if (!hasTitles)
                            {
                                for (var i = 0; i < fieldCount; i++)
                                {
                                    SetValue(record.GetName(i), fRo, co);
                                    co++;
                                }

                                co = Address.ColumnNumber;
                                hasTitles = true;
                            }

                            for (var i = 0; i < fieldCount; i++)
                            {
                                SetValue(record[i], ro, co);
                                co++;
                            }
                        }
                        else
                        {
                            if (!hasTitles)
                            {
                                foreach (var mi in members)
                                {
                                    if ((mi as IEnumerable) == null)
                                    {
                                        var fieldName = XLColumnAttribute.GetHeader(mi);
                                        if (XLHelper.IsNullOrWhiteSpace(fieldName))
                                            fieldName = mi.Name;

                                        SetValue(fieldName, fRo, co);
                                    }

                                    co++;
                                }

                                co = Address.ColumnNumber;
                                hasTitles = true;
                            }

                            foreach (var mi in members)
                            {
                                SetValue(accessor[m, mi.Name], ro, co);
                                co++;
                            }
                        }

                        if (co > maxCo)
                            maxCo = co;

                        ro++;
                    }
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

        public IXLTable InsertTable(DataTable data)
        {
            return InsertTable(data, null, true);
        }

        public IXLTable InsertTable(DataTable data, bool createTable)
        {
            return InsertTable(data, null, createTable);
        }

        public IXLTable InsertTable(DataTable data, string tableName)
        {
            return InsertTable(data, tableName, true);
        }

        public IXLTable InsertTable(DataTable data, string tableName, bool createTable)
        {
            if (data == null) return null;

            if (data.Rows.Count > 0) return InsertTable(data.AsEnumerable(), tableName, createTable);

            var ro = Address.RowNumber;
            var co = Address.ColumnNumber;

            foreach (DataColumn col in data.Columns)
            {
                SetValue(col.ColumnName, ro, co);
                co++;
            }

            ClearMerged();
            var range = _worksheet.Range(
                Address.RowNumber,
                Address.ColumnNumber,
                ro,
                co - 1);

            if (createTable) return tableName == null ? range.CreateTable() : range.CreateTable(tableName);

            return tableName == null ? range.AsTable() : range.AsTable(tableName);
        }

        public IXLRange InsertData(IEnumerable data)
        {
            if (data != null && data.GetType() != typeof(String))
            {
                var ro = Address.RowNumber;
                var maxCo = 0;
                var isDataTable = false;
                var isDataReader = false;

                const BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.Instance;
                var memberCache = new Dictionary<Type, IEnumerable<MemberInfo>>();
                var accessorCache = new Dictionary<Type, TypeAccessor>();
                IEnumerable<MemberInfo> members = null;
                TypeAccessor accessor = null;

                foreach (var m in data)
                {
                    var itemType = m.GetType();
                    if (!memberCache.ContainsKey(itemType))
                    {
                        var _accessor = TypeAccessor.Create(itemType);

                        var _members = itemType.GetFields(bindingFlags).Cast<MemberInfo>()
                             .Concat(itemType.GetProperties(bindingFlags))
                             .Where(mi => !XLColumnAttribute.IgnoreMember(mi))
                             .OrderBy(mi => XLColumnAttribute.GetOrder(mi));

                        memberCache.Add(itemType, _members);
                        accessorCache.Add(itemType, _accessor);
                    }

                    members = memberCache[itemType];
                    accessor = accessorCache[itemType];

                    var co = Address.ColumnNumber;

                    if (itemType.IsPrimitive || itemType == typeof(String) || itemType == typeof(DateTime) || itemType.IsNumber())
                    {
                        SetValue(m, ro, co);
                        co++;
                    }
                    else if (itemType.IsArray)
                    {
                        // dynamic arr = m;
                        foreach (var item in (Array)m)
                        {
                            SetValue(item, ro, co);
                            co++;
                        }
                    }
                    else if (isDataTable || m is DataRow)
                    {
                        if (!isDataTable)
                            isDataTable = true;

                        foreach (var item in (m as DataRow).ItemArray)
                        {
                            SetValue(item, ro, co);
                            co++;
                        }
                    }
                    else if (isDataReader || m is IDataRecord)
                    {
                        if (!isDataReader)
                            isDataReader = true;

                        var record = m as IDataRecord;

                        var fieldCount = record.FieldCount;
                        for (var i = 0; i < fieldCount; i++)
                        {
                            SetValue(record[i], ro, co);
                            co++;
                        }
                    }
                    else
                    {
                        foreach (var mi in members)
                        {
                            SetValue(accessor[m, mi.Name], ro, co);
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

        public IXLStyle Style
        {
            get { return GetStyle(); }

            set { SetStyle(value); }
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
                            _cellValue = dtTest.ToOADate().ToInvariantString();
                        else if (Double.TryParse(_cellValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out dblTest))
                            _cellValue = dblTest.ToInvariantString();
                        else
                        {
                            throw new ArgumentException(
                                string.Format(
                                    "Cannot set data type to DateTime because '{0}' is not recognized as a date.",
                                    _cellValue));
                        }
                        var style = GetStyleForRead();
                        if (style.NumberFormat.Format == String.Empty && style.NumberFormat.NumberFormatId == 0)
                            Style.NumberFormat.NumberFormatId = _cellValue.Contains('.') ? 22 : 14;
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
                                _cellValue = (DateTime.FromOADate(Double.Parse(_cellValue, XLHelper.NumberStyle, XLHelper.ParseCulture)) - BaseDate).ToString();
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
                        if (Double.TryParse(_cellValue, XLHelper.NumberStyle, CultureInfo.InvariantCulture, out dTest))
                            _cellValue = dTest.ToInvariantString();
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
                            _cellValue = BaseDate.Add(GetTimeSpan()).ToOADate().ToInvariantString();
                    }
                }

                _dataType = value;
            }
        }

        public IXLCell Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
        {
            return Clear(clearOptions, false);
        }

        internal IXLCell Clear(XLClearOptions clearOptions, bool calledFromRange)
        {
            //Note: We have to check if the cell is part of a merged range. If so we have to clear the whole range
            //Checking if called from range to avoid stack overflow
            if (IsMerged() && !calledFromRange)
            {
                using (var asRange = AsRange())
                {
                    var firstOrDefault = Worksheet.Internals.MergedRanges.FirstOrDefault(asRange.Intersects);
                    if (firstOrDefault != null)
                        firstOrDefault.Clear(clearOptions);
                }
            }
            else
            {
                if (clearOptions == XLClearOptions.Contents || clearOptions == XLClearOptions.ContentsAndFormats)
                {
                    Hyperlink = null;
                    _richText = null;
                    //_comment = null;
                    _cellValue = String.Empty;
                    FormulaA1 = String.Empty;
                }

                if (clearOptions == XLClearOptions.Formats || clearOptions == XLClearOptions.ContentsAndFormats)
                {
                    if (HasDataValidation)
                        DataValidation.Clear();

                    SetStyle(Worksheet.Style);
                }
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
                if (XLHelper.IsNullOrWhiteSpace(_formulaA1))
                {
                    if (!XLHelper.IsNullOrWhiteSpace(_formulaR1C1))
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
                _formulaA1 = XLHelper.IsNullOrWhiteSpace(value) ? null : value;

                _formulaR1C1 = null;
            }
        }

        public string FormulaR1C1
        {
            get
            {
                if (XLHelper.IsNullOrWhiteSpace(_formulaR1C1))
                    _formulaR1C1 = GetFormulaR1C1(FormulaA1);

                return _formulaR1C1;
            }

            set
            {
                _formulaR1C1 = XLHelper.IsNullOrWhiteSpace(value) ? null : value;
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

        public string ValueCached { get; internal set; }

        public IXLRichText RichText
        {
            get
            {
                if (_richText == null)
                {
                    var style = GetStyleForRead();
                    _richText = _cellValue.Length == 0
                                    ? new XLRichText(style.Font)
                                    : new XLRichText(GetFormattedString(), style.Font);
                }

                return _richText;
            }
        }

        public bool HasRichText
        {
            get { return _richText != null; }
        }

        IXLComment IXLCell.Comment
        {
            get { return Comment; }
        }

        public bool HasComment
        {
            get { return _comment != null; }
        }

        public Boolean IsMerged()
        {
            return Worksheet.Internals.MergedRanges.Any(r => r.Contains(this));
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
            {
                if (!Style.Equals(Worksheet.Style) || IsMerged() || HasComment || HasDataValidation)
                    return false;

                if (_style == null)
                {
                    XLRow row;
                    if (Worksheet.Internals.RowsCollection.TryGetValue(Address.RowNumber, out row) && !row.Style.Equals(Worksheet.Style))
                        return false;

                    XLColumn column;
                    if (Worksheet.Internals.ColumnsCollection.TryGetValue(Address.ColumnNumber, out column) && !column.Style.Equals(Worksheet.Style))
                        return false;
                }

                if (Worksheet.ConditionalFormats.Any(cf => cf.Range.Contains(this)))
                    return false;
            }
            return true;
        }

        public IXLColumn WorksheetColumn()
        {
            return Worksheet.Column(Address.ColumnNumber);
        }

        public IXLRow WorksheetRow()
        {
            return Worksheet.Row(Address.RowNumber);
        }

        public IXLCell CopyTo(IXLCell target)
        {
            (target as XLCell).CopyFrom(this, true);
            return target;
        }

        public IXLCell CopyTo(String target)
        {
            return CopyTo(GetTargetCell(target, Worksheet));
        }

        public IXLCell CopyFrom(IXLCell otherCell)
        {
            return CopyFrom(otherCell as XLCell, true);
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

        public Boolean HasDataValidation
        {
            get
            {
                using (var asRange = AsRange())
                    return Worksheet.DataValidations.Any(dv => dv.Ranges.Contains(asRange) && dv.IsDirty());
            }
        }

        public IXLDataValidation SetDataValidation()
        {
            return DataValidation;
        }

        public void Select()
        {
            AsRange().Select();
        }

        public IXLConditionalFormat AddConditionalFormat()
        {
            using (var r = AsRange())
                return r.AddConditionalFormat();
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

        public XLHyperlink GetHyperlink()
        {
            if (HasHyperlink)
                return Hyperlink;

            return Value as XLHyperlink;
        }

        public Boolean TryGetValue<T>(out T value)
        {
            var currValue = Value;

            if (currValue == null)
            {
                value = default(T);
                return true;
            }

            bool b;
            if (TryGetTimeSpanValue(out value, currValue, out b)) return b;

            if (TryGetRichStringValue(out value)) return true;

            if (TryGetStringValue(out value, currValue)) return true;

            var strValue = currValue.ToString();
            if (typeof(T) == typeof(bool)) return TryGetBasicValue<T, bool>(out value, strValue, bool.TryParse);
            if (typeof(T) == typeof(sbyte)) return TryGetBasicValue<T, sbyte>(out value, strValue, sbyte.TryParse);
            if (typeof(T) == typeof(byte)) return TryGetBasicValue<T, byte>(out value, strValue, byte.TryParse);
            if (typeof(T) == typeof(short)) return TryGetBasicValue<T, short>(out value, strValue, short.TryParse);
            if (typeof(T) == typeof(ushort)) return TryGetBasicValue<T, ushort>(out value, strValue, ushort.TryParse);
            if (typeof(T) == typeof(int)) return TryGetBasicValue<T, int>(out value, strValue, int.TryParse);
            if (typeof(T) == typeof(uint)) return TryGetBasicValue<T, uint>(out value, strValue, uint.TryParse);
            if (typeof(T) == typeof(long)) return TryGetBasicValue<T, long>(out value, strValue, long.TryParse);
            if (typeof(T) == typeof(ulong)) return TryGetBasicValue<T, ulong>(out value, strValue, ulong.TryParse);
            if (typeof(T) == typeof(float)) return TryGetBasicValue<T, float>(out value, strValue, float.TryParse);
            if (typeof(T) == typeof(double)) return TryGetBasicValue<T, double>(out value, strValue, double.TryParse);
            if (typeof(T) == typeof(decimal)) return TryGetBasicValue<T, decimal>(out value, strValue, decimal.TryParse);

            if (typeof(T) == typeof(XLHyperlink))
            {
                XLHyperlink tmp = GetHyperlink();
                if (tmp != null)
                {
                    value = (T)Convert.ChangeType(tmp, typeof(T));
                    return true;
                }

                value = default(T);
                return false;
            }

            try
            {
                value = (T)Convert.ChangeType(currValue, typeof(T));
                return true;
            }
            catch
            {
                value = default(T);
                return false;
            }
        }

        private static bool TryGetTimeSpanValue<T>(out T value, object currValue, out bool b)
        {
            if (typeof(T) == typeof(TimeSpan))
            {
                TimeSpan tmp;
                Boolean retVal = true;

                if (currValue is TimeSpan)
                {
                    tmp = (TimeSpan)currValue;
                }
                else if (!TimeSpan.TryParse(currValue.ToString(), out tmp))
                {
                    retVal = false;
                }

                value = (T)Convert.ChangeType(tmp, typeof(T));
                {
                    b = retVal;
                    return true;
                }
            }
            value = default(T);
            b = false;
            return false;
        }

        private bool TryGetRichStringValue<T>(out T value)
        {
            if (typeof(T) == typeof(IXLRichText))
            {
                value = (T)RichText;
                return true;
            }
            value = default(T);
            return false;
        }

        private static bool TryGetStringValue<T>(out T value, object currValue)
        {
            if (typeof(T) == typeof(String))
            {
                var valToUse = currValue.ToString();
                if (!utfPattern.Match(valToUse).Success)
                {
                    value = (T)Convert.ChangeType(valToUse, typeof(T));
                    return true;
                }

                var sb = new StringBuilder();
                var lastIndex = 0;
                foreach (Match match in utfPattern.Matches(valToUse))
                {
                    var matchString = match.Value;
                    var matchIndex = match.Index;
                    sb.Append(valToUse.Substring(lastIndex, matchIndex - lastIndex));

                    sb.Append((char)int.Parse(match.Groups[1].Value, NumberStyles.AllowHexSpecifier));

                    lastIndex = matchIndex + matchString.Length;
                }
                if (lastIndex < valToUse.Length)
                    sb.Append(valToUse.Substring(lastIndex));

                value = (T)Convert.ChangeType(sb.ToString(), typeof(T));
                return true;
            }
            value = default(T);
            return false;
        }

        private static Boolean TryGetBooleanValue<T>(out T value, object currValue)
        {
            if (typeof(T) == typeof(Boolean))
            {
                Boolean tmp;
                if (Boolean.TryParse(currValue.ToString(), out tmp))
                {
                    value = (T)Convert.ChangeType(tmp, typeof(T));
                    {
                        return true;
                    }
                }
            }
            value = default(T);
            return false;
        }

        private delegate Boolean Func<T>(String input, out T output);

        private static Boolean TryGetBasicValue<T, U>(out T value, String currValue, Func<U> func)
        {
            U tmp;
            if (func(currValue, out tmp))
            {
                value = (T)Convert.ChangeType(tmp, typeof(T));
                {
                    return true;
                }
            }
            value = default(T);
            return false;
        }

        #endregion IXLCell Members

        #region IXLStylized Members

        public Boolean StyleChanged { get; set; }

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
                var retVal = new XLRanges { AsRange() };
                return retVal;
            }
        }

        #endregion IXLStylized Members

        private bool SetRangeColumns(object value)
        {
            var columns = value as XLRangeColumns;
            if (columns == null)
                return SetColumns(value);

            var cell = this;
            foreach (var column in columns)
            {
                cell.SetRange(column);
                cell = cell.CellRight();
            }
            return true;
        }

        private bool SetColumns(object value)
        {
            var columns = value as XLColumns;
            if (columns == null)
                return false;

            var cell = this;
            foreach (var column in columns)
            {
                cell.SetRange(column);
                cell = cell.CellRight();
            }
            return true;
        }

        private bool SetRangeRows(object value)
        {
            var rows = value as XLRangeRows;
            if (rows == null)
                return SetRows(value);

            var cell = this;
            foreach (var row in rows)
            {
                cell.SetRange(row);
                cell = cell.CellBelow();
            }
            return true;
        }

        private bool SetRows(object value)
        {
            var rows = value as XLRows;
            if (rows == null)
                return false;

            var cell = this;
            foreach (var row in rows)
            {
                cell.SetRange(row);
                cell = cell.CellBelow();
            }
            return true;
        }

        public XLRange AsRange()
        {
            return _worksheet.Range(Address, Address);
        }

        private IXLStyle GetStyle()
        {
            if (_style != null)
                return _style;

            return _style = new XLStyle(this, Worksheet.Workbook.GetStyleById(_styleCacheId));
        }

        public void DeleteComment()
        {
            _comment = null;
        }

        private bool IsDateFormat()
        {
            var style = GetStyleForRead();
            return _dataType == XLCellValues.Number
                   && XLHelper.IsNullOrWhiteSpace(style.NumberFormat.Format)
                   && ((style.NumberFormat.NumberFormatId >= 14
                        && style.NumberFormat.NumberFormatId <= 22)
                       || (style.NumberFormat.NumberFormatId >= 45
                           && style.NumberFormat.NumberFormatId <= 47));
        }

        private string GetFormat()
        {
            var format = String.Empty;
            var style = GetStyleForRead();
            if (XLHelper.IsNullOrWhiteSpace(style.NumberFormat.Format))
            {
                var formatCodes = GetFormatCodes();
                if (formatCodes.ContainsKey(style.NumberFormat.NumberFormatId))
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
                    var maxRows = asRange.RowCount();
                    var maxColumns = asRange.ColumnCount();
                    Worksheet.Range(Address.RowNumber, Address.ColumnNumber, maxRows, maxColumns).Clear();
                }

                var minRow = asRange.RangeAddress.FirstAddress.RowNumber;
                var minColumn = asRange.RangeAddress.FirstAddress.ColumnNumber;
                foreach (var sourceCell in asRange.CellsUsed(true))
                {
                    Worksheet.Cell(
                        Address.RowNumber + sourceCell.Address.RowNumber - minRow,
                        Address.ColumnNumber + sourceCell.Address.ColumnNumber - minColumn
                        ).CopyFrom(sourceCell, true);
                }

                var rangesToMerge = (from mergedRange in (asRange.Worksheet).Internals.MergedRanges
                                     where asRange.Contains(mergedRange)
                                     let initialRo =
                                         Address.RowNumber +
                                         (mergedRange.RangeAddress.FirstAddress.RowNumber -
                                          asRange.RangeAddress.FirstAddress.RowNumber)
                                     let initialCo =
                                         Address.ColumnNumber +
                                         (mergedRange.RangeAddress.FirstAddress.ColumnNumber -
                                          asRange.RangeAddress.FirstAddress.ColumnNumber)
                                     select
                                         Worksheet.Range(initialRo, initialCo, initialRo + mergedRange.RowCount() - 1,
                                                         initialCo + mergedRange.ColumnCount() - 1)).Cast<IXLRange>().
                    ToList();
                rangesToMerge.ForEach(r => r.Merge());

                return true;
            }

            return false;
        }

        private bool SetEnumerable(object collectionObject)
        {
            // IXLRichText implements IEnumerable, but we don't want to handle this here.
            if ((collectionObject as IXLRichText) != null) return false;

            var asEnumerable = collectionObject as IEnumerable;
            return InsertData(asEnumerable) != null;
        }

        private void ClearMerged()
        {
            List<IXLRange> mergeToDelete;
            using (var asRange = AsRange())
                mergeToDelete = Worksheet.Internals.MergedRanges.Where(merge => merge.Intersects(asRange)).ToList();

            mergeToDelete.ForEach(m => Worksheet.Internals.MergedRanges.Remove(m));
        }

        private void SetValue<T>(T value, int ro, int co) where T : class
        {
            if (value == null)
                _worksheet.Cell(ro, co).SetValue(String.Empty);
            else
            {
                if (value is IConvertible)
                    _worksheet.Cell(ro, co).SetValue((T)Convert.ChangeType(value, typeof(T)));
                else
                    _worksheet.Cell(ro, co).SetValue(value);
            }
        }

        private void SetValue(object value)
        {
            FormulaA1 = String.Empty;
            string val;
            if (value == null)
                val = string.Empty;
            else if (value is DateTime)
                val = ((DateTime)value).ToString("o");
            else if (value.IsNumber())
                val = Convert.ToDecimal(value).ToInvariantString();
            else
                val = value.ToString();
            _richText = null;
            if (val.Length == 0)
                _dataType = XLCellValues.Text;
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
                else if (value is TimeSpan || (!Double.TryParse(val, XLHelper.NumberStyle, XLHelper.ParseCulture, out dTest) && TimeSpan.TryParse(val, out tsTest)))
                {
                    if (!(value is TimeSpan) && TimeSpan.TryParse(val, out tsTest))
                        val = tsTest.ToString();

                    _dataType = XLCellValues.TimeSpan;
                    if (style.NumberFormat.Format == String.Empty && style.NumberFormat.NumberFormatId == 0)
                        Style.NumberFormat.NumberFormatId = 46;
                }
                else if (val.Trim() != "NaN" && Double.TryParse(val, XLHelper.NumberStyle, XLHelper.ParseCulture, out dTest))
                    _dataType = XLCellValues.Number;
                else if (DateTime.TryParse(val, out dtTest) && dtTest >= BaseDate)
                {
                    _dataType = XLCellValues.DateTime;

                    if (style.NumberFormat.Format == String.Empty && style.NumberFormat.NumberFormatId == 0)
                        Style.NumberFormat.NumberFormatId = dtTest.Date == dtTest ? 14 : 22;
                    {
                        DateTime forMillis;
                        if (value is DateTime && (forMillis = (DateTime)value).Millisecond > 0)
                        {
                            val = forMillis.ToOADate().ToInvariantString();
                        }
                        else
                        {
                            val = dtTest.ToOADate().ToInvariantString();
                        }
                    }
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
            if (val.Length > 32767) throw new ArgumentException("Cells can only hold 32,767 characters.");
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
                        {7, "$#,##0.00_);($#,##0.00)"},
                        {9, "0%"},
                        {10, "0.00%"},
                        {11, "0.00E+00"},
                        {12, "# ?/?"},
                        {13, "# ??/??"},
                        {14, "M/d/yyyy"},
                        {15, "d-MMM-yy"},
                        {16, "d-MMM"},
                        {17, "MMM-yy"},
                        {18, "h:mm tt"},
                        {19, "h:mm:ss tt"},
                        {20, "H:mm"},
                        {21, "H:mm:ss"},
                        {22, "M/d/yyyy H:mm"},
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
            if (XLHelper.IsNullOrWhiteSpace(strValue))
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

            var rowPart = addressToUse.Substring(0, addressToUse.IndexOf("C"));
            var rowToReturn = GetA1Row(rowPart, rowsToShift);

            var columnPart = addressToUse.Substring(addressToUse.IndexOf("C"));
            var columnToReturn = GetA1Column(columnPart, columnsToShift);

            var retAddress = columnToReturn + rowToReturn;
            return retAddress;
        }

        private string GetA1Column(string columnPart, int columnsToShift)
        {
            string columnToReturn;
            if (columnPart == "C")
                columnToReturn = XLHelper.GetColumnLetterFromNumber(Address.ColumnNumber + columnsToShift);
            else
            {
                var bIndex = columnPart.IndexOf("[");
                var mIndex = columnPart.IndexOf("-");
                if (bIndex >= 0)
                {
                    columnToReturn = XLHelper.GetColumnLetterFromNumber(
                        Address.ColumnNumber +
                        Int32.Parse(columnPart.Substring(bIndex + 1, columnPart.Length - bIndex - 2)) + columnsToShift
                        );
                }
                else if (mIndex >= 0)
                {
                    columnToReturn = XLHelper.GetColumnLetterFromNumber(
                        Address.ColumnNumber + Int32.Parse(columnPart.Substring(mIndex)) + columnsToShift
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
                rowToReturn = (Address.RowNumber + rowsToShift).ToString();
            else
            {
                var bIndex = rowPart.IndexOf("[");
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
                var p1 = parts[0];
                var p2 = parts[1];
                int row1;
                if (Int32.TryParse(p1.Replace("$", string.Empty), out row1))
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

            var address = XLAddress.Create(_worksheet, a1Address);

            var rowPart = GetR1C1Row(address.RowNumber, address.FixedRow, rowsToShift);
            var columnPart = GetR1C1Column(address.ColumnNumber, address.FixedColumn, columnsToShift);

            return rowPart + columnPart;
        }

        private string GetR1C1Row(int rowNumber, bool fixedRow, int rowsToShift)
        {
            string rowPart;
            rowNumber += rowsToShift;
            var rowDiff = rowNumber - Address.RowNumber;
            if (rowDiff != 0 || fixedRow)
                rowPart = fixedRow ? String.Format("R{0}", rowNumber) : String.Format("R[{0}]", rowDiff);
            else
                rowPart = "R";

            return rowPart;
        }

        private string GetR1C1Column(int columnNumber, bool fixedColumn, int columnsToShift)
        {
            string columnPart;
            columnNumber += columnsToShift;
            var columnDiff = columnNumber - Address.ColumnNumber;
            if (columnDiff != 0 || fixedColumn)
                columnPart = fixedColumn ? String.Format("C{0}", columnNumber) : String.Format("C[{0}]", columnDiff);
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
            _comment = source._comment == null ? null : new XLComment(this, source._comment, source.Style.Font);

            if (source._hyperlink != null)
            {
                SettingHyperlink = true;
                Hyperlink = new XLHyperlink(source.Hyperlink);
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

        public IXLCell CopyFrom(IXLCell otherCell, Boolean copyDataValidations)
        {
            var source = otherCell as XLCell; // To expose GetFormulaR1C1, etc

            CopyValuesFrom(source);

            SetStyle(source._style ?? source.Worksheet.Workbook.GetStyleById(source._styleCacheId));

            var conditionalFormats = source.Worksheet.ConditionalFormats.Where(c => c.Range.Contains(source)).ToList();
            foreach (var cf in conditionalFormats)
            {
                var c = new XLConditionalFormat(cf as XLConditionalFormat) { Range = AsRange() };
                var oldValues = c.Values.Values.ToList();
                c.Values.Clear();
                foreach (var v in oldValues)
                {
                    var f = v.Value;
                    if (v.IsFormula)
                    {
                        var r1c1 = source.GetFormulaR1C1(f);
                        f = GetFormulaA1(r1c1);
                    }

                    c.Values.Add(new XLFormula { _value = f, IsFormula = v.IsFormula });
                }

                _worksheet.ConditionalFormats.Add(c);
            }

            if (copyDataValidations)
            {
                var eventTracking = Worksheet.EventTrackingEnabled;
                Worksheet.EventTrackingEnabled = false;
                if (source.HasDataValidation)
                    CopyDataValidation(source, source.DataValidation);
                else if (HasDataValidation)
                {
                    using (var asRange = AsRange())
                        Worksheet.DataValidations.Delete(asRange);
                }
                Worksheet.EventTrackingEnabled = eventTracking;
            }

            return this;
        }

        internal void CopyDataValidation(XLCell otherCell, XLDataValidation otherDv)
        {
            var thisDv = DataValidation;
            thisDv.CopyFrom(otherDv);
            thisDv.Value = GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.Value));
            thisDv.MinValue = GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.MinValue));
            thisDv.MaxValue = GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.MaxValue));
        }

        internal void ShiftFormulaRows(XLRange shiftedRange, int rowsShifted)
        {
            _formulaA1 = ShiftFormulaRows(FormulaA1, Worksheet, shiftedRange, rowsShifted);
        }

        internal static String ShiftFormulaRows(String formulaA1, XLWorksheet worksheetInAction, XLRange shiftedRange,
                                                int rowsShifted)
        {
            if (XLHelper.IsNullOrWhiteSpace(formulaA1)) return String.Empty;

            var value = formulaA1; // ">" + formulaA1 + "<";

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

                                    sb.Append(useSheetName
                                                  ? String.Format("{0}!{1}:{2}", sheetName.WrapSheetNameInQuotesIfRequired(), row1, row2)
                                                  : String.Format("{0}:{1}", row1, row2));
                                }
                                else if (shiftedRange.RangeAddress.FirstAddress.RowNumber <=
                                         matchRange.RangeAddress.FirstAddress.RowNumber)
                                {
                                    if (rangeAddress.Contains(':'))
                                    {
                                        if (useSheetName)
                                        {
                                            sb.Append(String.Format("{0}!{1}:{2}",
                                                                    sheetName.WrapSheetNameInQuotesIfRequired(),
                                                                    new XLAddress(worksheetInAction,
                                                                                  XLHelper.TrimRowNumber(matchRange.RangeAddress.FirstAddress.RowNumber + rowsShifted),
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.ColumnLetter,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedColumn),
                                                                    new XLAddress(worksheetInAction,
                                                                                  XLHelper.TrimRowNumber(matchRange.RangeAddress.LastAddress.RowNumber + rowsShifted),
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
                                                                    new XLAddress(worksheetInAction,
                                                                                  XLHelper.TrimRowNumber(matchRange.RangeAddress.FirstAddress.RowNumber + rowsShifted),
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.ColumnLetter,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedColumn),
                                                                    new XLAddress(worksheetInAction,
                                                                                  XLHelper.TrimRowNumber(matchRange.RangeAddress.LastAddress.RowNumber + rowsShifted),
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
                                            sb.Append(String.Format("{0}!{1}",
                                                                    sheetName.WrapSheetNameInQuotesIfRequired(),
                                                                    new XLAddress(worksheetInAction,
                                                                                  XLHelper.TrimRowNumber(matchRange.RangeAddress.FirstAddress.RowNumber + rowsShifted),
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
                                                                    new XLAddress(worksheetInAction,
                                                                                  XLHelper.TrimRowNumber(matchRange.RangeAddress.FirstAddress.RowNumber + rowsShifted),
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
                                        sb.Append(String.Format("{0}!{1}:{2}",
                                                                sheetName.WrapSheetNameInQuotesIfRequired(),
                                                                matchRange.RangeAddress.FirstAddress,
                                                                new XLAddress(worksheetInAction,
                                                                              XLHelper.TrimRowNumber(matchRange.RangeAddress.LastAddress.RowNumber + rowsShifted),
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
                                                                new XLAddress(worksheetInAction,
                                                                              XLHelper.TrimRowNumber(matchRange.RangeAddress.LastAddress.RowNumber + rowsShifted),
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

            return sb.ToString();

            //string retVal = sb.ToString();
            //return retVal.Substring(1, retVal.Length - 2);
        }

        internal void ShiftFormulaColumns(XLRange shiftedRange, int columnsShifted)
        {
            _formulaA1 = ShiftFormulaColumns(FormulaA1, Worksheet, shiftedRange, columnsShifted);
        }

        internal static String ShiftFormulaColumns(String formulaA1, XLWorksheet worksheetInAction, XLRange shiftedRange,
                                                   int columnsShifted)
        {
            if (XLHelper.IsNullOrWhiteSpace(formulaA1)) return String.Empty;

            var value = formulaA1; // ">" + formulaA1 + "<";

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

                                    sb.Append(useSheetName
                                                  ? String.Format("{0}!{1}:{2}", sheetName.WrapSheetNameInQuotesIfRequired(), column1, column2)
                                                  : String.Format("{0}:{1}", column1, column2));
                                }
                                else if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                                         matchRange.RangeAddress.FirstAddress.ColumnNumber)
                                {
                                    if (rangeAddress.Contains(':'))
                                    {
                                        if (useSheetName)
                                        {
                                            sb.Append(String.Format("{0}!{1}:{2}",
                                                                    sheetName.WrapSheetNameInQuotesIfRequired(),
                                                                    new XLAddress(worksheetInAction,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.RowNumber,
                                                                                  XLHelper.TrimColumnNumber(matchRange.RangeAddress.FirstAddress.ColumnNumber + columnsShifted),
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedColumn),
                                                                    new XLAddress(worksheetInAction,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.RowNumber,
                                                                                  XLHelper.TrimColumnNumber(matchRange.RangeAddress.LastAddress.ColumnNumber + columnsShifted),
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.FixedColumn)));
                                        }
                                        else
                                        {
                                            sb.Append(String.Format("{0}:{1}",
                                                                    new XLAddress(worksheetInAction,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.RowNumber,
                                                                                  XLHelper.TrimColumnNumber(matchRange.RangeAddress.FirstAddress.ColumnNumber + columnsShifted),
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedColumn),
                                                                    new XLAddress(worksheetInAction,
                                                                                  matchRange.RangeAddress.
                                                                                      LastAddress.RowNumber,
                                                                                  XLHelper.TrimColumnNumber(matchRange.RangeAddress.LastAddress.ColumnNumber + columnsShifted),
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
                                            sb.Append(String.Format("{0}!{1}",
                                                                    sheetName.WrapSheetNameInQuotesIfRequired(),
                                                                    new XLAddress(worksheetInAction,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.RowNumber,
                                                                                  XLHelper.TrimColumnNumber(matchRange.RangeAddress.FirstAddress.ColumnNumber + columnsShifted),
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedRow,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.FixedColumn)));
                                        }
                                        else
                                        {
                                            sb.Append(String.Format("{0}",
                                                                    new XLAddress(worksheetInAction,
                                                                                  matchRange.RangeAddress.
                                                                                      FirstAddress.RowNumber,
                                                                                  XLHelper.TrimColumnNumber(matchRange.RangeAddress.FirstAddress.ColumnNumber + columnsShifted),
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
                                        sb.Append(String.Format("{0}!{1}:{2}",
                                                                sheetName.WrapSheetNameInQuotesIfRequired(),
                                                                matchRange.RangeAddress.FirstAddress,
                                                                new XLAddress(worksheetInAction,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.RowNumber,
                                                                              XLHelper.TrimColumnNumber(matchRange.RangeAddress.LastAddress.ColumnNumber + columnsShifted),
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.FixedRow,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.FixedColumn)));
                                    }
                                    else
                                    {
                                        sb.Append(String.Format("{0}:{1}",
                                                                matchRange.RangeAddress.FirstAddress,
                                                                new XLAddress(worksheetInAction,
                                                                              matchRange.RangeAddress.
                                                                                  LastAddress.RowNumber,
                                                                              XLHelper.TrimColumnNumber(matchRange.RangeAddress.LastAddress.ColumnNumber + columnsShifted),
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

            return sb.ToString();

            //string retVal = sb.ToString();
            //return retVal.Substring(1, retVal.Length - 2);
        }

        // --

        private XLCell CellShift(Int32 rowsToShift, Int32 columnsToShift)
        {
            return Worksheet.Cell(Address.RowNumber + rowsToShift, Address.ColumnNumber + columnsToShift);
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

        public Boolean HasFormula { get { return !XLHelper.IsNullOrWhiteSpace(FormulaA1); } }

        public IXLRangeAddress FormulaReference { get; set; }
    }
}
