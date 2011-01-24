using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections;
using System.Data;


namespace ClosedXML.Excel
{
    internal class XLCell : IXLCell, IXLStylized
    {
        public static readonly DateTime baseDate = new DateTime(1899, 12, 30);
        XLWorksheet worksheet;
        public XLCell(IXLAddress address, IXLStyle defaultStyle, XLWorksheet worksheet)
        {
            this.Address = address;
            this.ShareString = true;

            if (defaultStyle == null) 
                style = new XLStyle(this, worksheet.Style);
            else
                style = new XLStyle(this, defaultStyle);
            this.worksheet = worksheet;
        }

        public IXLAddress Address { get; private set; }
        public String InnerText
        {
            get 
            {
                if (StringExtensions.IsNullOrWhiteSpace(cellValue))
                    return FormulaA1;
                else
                    return cellValue;
            }
        }
        public IXLRange AsRange()
        {
            return worksheet.Range(Address, Address);
        }
        public T GetValue<T>() 
        {
            if (Value is TimeSpan)
                if (typeof(T) == typeof(String))
                    return (T)Convert.ChangeType(Value.ToString(), typeof(T));
                else
                    return (T)Value;
            else
                return (T)Convert.ChangeType(Value, typeof(T));
        }
        public String GetString()
        {
            return GetValue<String>();
        }
        public Double GetDouble()
        {
            return GetValue<Double>();
        }
        public Boolean GetBoolean()
        {
            return GetValue<Boolean>();
        }
        public DateTime GetDateTime()
        {
            return GetValue<DateTime>();
        }
        public TimeSpan GetTimeSpan()
        {
            return GetValue<TimeSpan>();
        }
        public String GetFormattedString()
        {
            if (dataType == XLCellValues.Boolean)
            {
                return (cellValue != "0").ToString();
            }
            else if (dataType == XLCellValues.TimeSpan)
            {
                return cellValue;
            }
            else if (dataType == XLCellValues.DateTime || IsDateFormat())
            {
                Double dTest;
                if (Double.TryParse(cellValue, out dTest))
                {
                    String format = GetFormat();
                    return DateTime.FromOADate(dTest).ToString(format);
                }
                else
                {
                    return cellValue;
                }
            }
            else if (dataType == XLCellValues.Number)
            {
                Double dTest;
                if (Double.TryParse(cellValue, out dTest))
                {
                    String format = GetFormat();
                    return dTest.ToString(format);
                }
                else
                {
                    return cellValue;
                }
            }
            else
            {
                return cellValue;
            }
        }

        private bool IsDateFormat()
        {
            return (dataType == XLCellValues.Number 
                && StringExtensions.IsNullOrWhiteSpace(Style.NumberFormat.Format)
                && ((Style.NumberFormat.NumberFormatId >= 14
                    && Style.NumberFormat.NumberFormatId <= 22)
                || (Style.NumberFormat.NumberFormatId >= 45
                    && Style.NumberFormat.NumberFormatId <= 47)));
        }

        private String GetFormat()
        {
            String format;
            if (StringExtensions.IsNullOrWhiteSpace(Style.NumberFormat.Format))
            {
                var formatCodes = GetFormatCodes();
                format = formatCodes[Style.NumberFormat.NumberFormatId];
            }
            else
            {
                format = Style.NumberFormat.Format;
            }
            return format;
        }

        internal String cellValue = String.Empty;
        public Object Value
        {
            get
            {
                var fA1 = FormulaA1;
                if (!StringExtensions.IsNullOrWhiteSpace(fA1))
                    return fA1;

                if (dataType == XLCellValues.Boolean)
                {
                    return cellValue != "0";
                }
                else if (dataType == XLCellValues.DateTime)
                {
                    return DateTime.FromOADate(Double.Parse(cellValue));
                }
                else if (dataType == XLCellValues.Number)
                {
                    return Double.Parse(cellValue);
                }
                else if (dataType == XLCellValues.TimeSpan)
                {
                    //return (DateTime.FromOADate(Double.Parse(cellValue)) - baseDate);
                    return TimeSpan.Parse(cellValue);
                }
                else
                {
                    return cellValue;
                }
            }
            set
            {   
                FormulaA1 = String.Empty;
                if (!SetEnumerable(value))
                    if (!SetRange(value))
                        SetValue(value);
            }
        }

        private Boolean SetRange(Object rangeObject)
        {
            var asRange = rangeObject as XLRangeBase;
            if (asRange != null)
            {
                Int32 maxRows;
                Int32 maxColumns;
                if (asRange is XLRow || asRange is XLColumn)
                {
                    var lastCellUsed = asRange.LastCellUsed();
                    maxRows = lastCellUsed.Address.RowNumber;
                    maxColumns = lastCellUsed.Address.ColumnNumber;
                    //if (asRange is XLRow)
                    //    worksheet.Range(Address.RowNumber, Address.ColumnNumber,  , maxColumns).Clear();
                }
                else
                {
                    maxRows = asRange.RowCount();
                    maxColumns = asRange.ColumnCount();
                    worksheet.Range(Address.RowNumber, Address.ColumnNumber, maxRows, maxColumns).Clear();
                }
                
                for (var ro = 1; ro <= maxRows; ro++)
                {
                    for (var co = 1; co <= maxColumns; co++)
                    {
                        var sourceCell = (XLCell)asRange.Cell(ro, co);
                        var targetCell = (XLCell)worksheet.Cell(Address.RowNumber + ro - 1, Address.ColumnNumber + co - 1);
                        targetCell.CopyValues(sourceCell);
                        targetCell.Style = sourceCell.style;
                    }
                }
                var rangesToMerge = new List<IXLRange>();
                foreach (var mergedRange in asRange.Worksheet.Internals.MergedRanges)
                {
                    if (asRange.Contains(mergedRange))
                    {
                        var initialRo = Address.RowNumber + (mergedRange.RangeAddress.FirstAddress.RowNumber - asRange.RangeAddress.FirstAddress.RowNumber);
                        var initialCo = Address.ColumnNumber + (mergedRange.RangeAddress.FirstAddress.ColumnNumber - asRange.RangeAddress.FirstAddress.ColumnNumber);
                        rangesToMerge.Add(worksheet.Range(initialRo, initialCo, initialRo + mergedRange.RowCount() - 1, initialCo + mergedRange.ColumnCount() - 1));
                    }
                }
                rangesToMerge.ForEach(r => r.Merge());

                return true;
            }
            else
            {
                return false;
            }

        }

        private Boolean SetEnumerable(Object collectionObject)
        {
            var asEnumerable = collectionObject as IEnumerable;
            return InsertData(asEnumerable) != null;
        }

        public IXLTable InsertTable(IEnumerable data)
        {
            return InsertTable(data, null, true);
        }
        public IXLTable InsertTable(IEnumerable data, Boolean createTable)
        {
            return InsertTable(data, null, createTable);
        }
        public IXLTable InsertTable(IEnumerable data, String tableName)
        {
            return InsertTable(data, tableName, true);
        }
        public IXLTable InsertTable(IEnumerable data, String tableName, Boolean createTable)
        {
            if (data != null && data.GetType() != typeof(String))
            {
                Int32 co;
                Int32 ro = Address.RowNumber + 1;
                Int32 fRo = Address.RowNumber;
                Boolean hasTitles = false;
                Int32 maxCo = 0;
                Boolean isDataTable = false;
                foreach (var m in data)
                {
                    co = Address.ColumnNumber;

                    if (m.GetType().IsPrimitive || m.GetType() == typeof(String) || m.GetType() == typeof(DateTime))
                    {
                        if (!hasTitles)
                        {
                            SetValue(m.GetType().Name, fRo, co);
                            hasTitles = true;
                            co = Address.ColumnNumber;
                        }
                        SetValue(m, ro, co);
                        co++;
                    }
                    else if (m.GetType().IsArray)
                    {
                        foreach (var item in (Array)m)
                        {
                            SetValue(item, ro, co);
                            co++;
                        }
                    }
                    else if (isDataTable || (m as DataRow) != null)
                    {
                        if (!isDataTable) isDataTable = true;
                        if (!hasTitles)
                        {
                            foreach (DataColumn column in (m as DataRow).Table.Columns)
                            {
                                SetValue(column.ColumnName, fRo, co);
                                co++;
                            }
                            co = Address.ColumnNumber;
                            hasTitles = true;
                        }

                        foreach (var item in (m as DataRow).ItemArray)
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
                            foreach (var info in fieldInfo)
                            {
                                SetValue(info.Name, fRo, co);
                                co++;
                            }
                            
                            foreach (var info in propertyInfo)
                            {
                                if ((info as IEnumerable) == null)
                                    SetValue(info.Name, fRo, co);
                                co++;
                            }
                            co = Address.ColumnNumber;
                            hasTitles = true;
                        }

                        foreach (var info in fieldInfo)
                        {
                            SetValue(info.GetValue(m), ro, co);
                            co++;
                        }

                        foreach (var info in propertyInfo)
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
                var range = worksheet.Range(
                    Address.RowNumber,
                    Address.ColumnNumber,
                    ro - 1,
                    maxCo - 1);

                if (createTable)
                {
                    if (tableName == null)
                        return range.CreateTable();
                    else
                        return range.CreateTable(tableName);
                }
                else
                {
                    if (tableName == null)
                        return range.AsTable();
                    else
                        return range.AsTable(tableName);
                }
            }
            else
            {
                return null;
            }
        }

        public IXLRange InsertData(IEnumerable data)
        {
            if (data != null && data.GetType() != typeof(String))
            {
                Int32 ro = Address.RowNumber;
                Int32 maxCo = 0;
                foreach (var m in data)
                {
                    Int32 co = Address.ColumnNumber;

                    if (m.GetType().IsPrimitive || m.GetType() == typeof(String) || m.GetType() == typeof(DateTime))
                    {
                        SetValue(m, ro, co);
                    }
                    else if (m.GetType().IsArray)
                    {
                        //dynamic arr = m;
                        foreach (var item in (Array)m)
                        {
                            SetValue(item, ro, co);
                            co++;
                        }
                    }
                    else if ((m as DataRow) != null)
                    {
                        foreach (var item in (m as DataRow).ItemArray)
                        {
                            SetValue(item, ro, co);
                            co++;
                        }
                    }
                    else
                    {
                        var fieldInfo = m.GetType().GetFields();
                        foreach (var info in fieldInfo)
                        {
                            SetValue(info.GetValue(m), ro, co);
                            co++;
                        }
                        var propertyInfo = m.GetType().GetProperties();
                        foreach (var info in propertyInfo)
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
                return worksheet.Range(
                    Address.RowNumber,
                    Address.ColumnNumber,
                    Address.RowNumber + ro - 1,
                    Address.ColumnNumber + maxCo - 1);
            }
            else
            {
                return null;
            }
        }

        private void ClearMerged(Int32 rowCount, Int32 columnCount)
        {
            List<IXLRange> mergeToDelete = new List<IXLRange>();
            foreach (var merge in worksheet.Internals.MergedRanges)
            {
                if (merge.Intersects(AsRange()))
                {
                    mergeToDelete.Add(merge);
                }
            }
            mergeToDelete.ForEach(m => worksheet.Internals.MergedRanges.Remove(m));
        }

        private void SetValue(object objWithValue, int ro, int co)
        {
            String str = String.Empty;
            if (objWithValue != null)
                str = objWithValue.ToString();

            worksheet.Cell(ro, co).Value = str;
        }

        private void SetValue(Object value)
        {
            FormulaA1 = String.Empty;
            String val = value.ToString();
            Double dTest;
            DateTime dtTest;
            Boolean bTest;
            TimeSpan tsTest;
            if (val.Length > 0 && val.Substring(0, 1) == "'")
            {
                val = val.Substring(1, val.Length - 1);
                dataType = XLCellValues.Text;
            }
            else if (value is TimeSpan || (TimeSpan.TryParse(val, out tsTest) && !Double.TryParse(val, out dTest)))
            {
                //if (TimeSpan.TryParse(val, out tsTest))
                //    val = baseDate.Add(tsTest).ToOADate().ToString();
                //else
                //{
                //    TimeSpan timeSpan = (TimeSpan)value;
                //    val = baseDate.Add(timeSpan).ToOADate().ToString();
                //}
                dataType = XLCellValues.TimeSpan;
                Style.NumberFormat.NumberFormatId = 46;
            }
            else if (Double.TryParse(val, out dTest))
            {
                dataType = XLCellValues.Number;
            }
            else if (DateTime.TryParse(val, out dtTest))
            {
                dataType = XLCellValues.DateTime;

                if (dtTest.Date == dtTest)
                    Style.NumberFormat.NumberFormatId = 14;
                else
                    Style.NumberFormat.NumberFormatId = 22;

                val = dtTest.ToOADate().ToString();
            }
            else if (Boolean.TryParse(val, out bTest))
            {
                dataType = XLCellValues.Boolean;
                val = bTest ? "1" : "0";
            }
            else
            {
                dataType = XLCellValues.Text;
            }

            cellValue = val;
        }

        #region IXLStylized Members

        private IXLStyle style;
        public IXLStyle Style
        {
            get
            {
                return style;
            }
            set
            {
                style = new XLStyle(null, value);
            }
        }

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                yield return style;
                UpdatingStyle = false;
            }
        }

        public Boolean UpdatingStyle { get; set; }

        public IXLStyle InnerStyle
        {
            get { return style; }
            set { style = new XLStyle(this, value); }
        }

        #endregion

        private XLCellValues dataType;
        public XLCellValues DataType
        {
            get
            {
                return dataType;
            }
            set
            {
                if (dataType != value)
                {
                    if (cellValue.Length > 0)
                    {
                        if (value == XLCellValues.Boolean)
                        {
                            Boolean bTest;
                            if (Boolean.TryParse(cellValue, out bTest))
                                cellValue = bTest ? "1" : "0";
                            else
                                cellValue = cellValue == "0" || String.IsNullOrEmpty(cellValue) ? "0" : "1";
                        }
                        else if (value == XLCellValues.DateTime)
                        {
                            DateTime dtTest;
                            Double dblTest;
                            if (DateTime.TryParse(cellValue, out dtTest))
                            {
                                cellValue = dtTest.ToOADate().ToString();
                            }
                            else if (Double.TryParse(cellValue, out dblTest))
                            {
                                cellValue = dblTest.ToString();
                            }
                            else
                            {
                                throw new ArgumentException("Cannot set data type to DateTime because '" + cellValue + "' is not recognized as a date.");
                            }

                            if (Style.NumberFormat.Format == String.Empty && Style.NumberFormat.NumberFormatId == 0)
                                if (cellValue.Contains('.'))
                                    Style.NumberFormat.NumberFormatId = 22;
                                else
                                    Style.NumberFormat.NumberFormatId = 14;
                        }
                        else if (value == XLCellValues.TimeSpan)
                        {
                            TimeSpan tsTest;
                            if (TimeSpan.TryParse(cellValue, out tsTest))
                            {
                                cellValue = tsTest.ToString();
                                if (Style.NumberFormat.Format == String.Empty && Style.NumberFormat.NumberFormatId == 0)
                                    Style.NumberFormat.NumberFormatId = 46;
                            }
                            else
                            {
                                try
                                {
                                    cellValue = (DateTime.FromOADate(Double.Parse(cellValue)) - baseDate).ToString();
                                }
                                catch
                                {
                                    throw new ArgumentException("Cannot set data type to TimeSpan because '" + cellValue + "' is not recognized as a TimeSpan.");
                                }
                            }
                        }
                        else if (value == XLCellValues.Number)
                        {
                            Double dTest;
                            if (Double.TryParse(cellValue, out dTest))
                            {
                                cellValue = Double.Parse(cellValue).ToString();
                            }
                            else
                            {
                                throw new ArgumentException("Cannot set data type to Number because '" + cellValue + "' is not recognized as a number.");
                            }
                        }
                        else
                        {
                            var formatCodes = GetFormatCodes();
                            if (dataType == XLCellValues.Boolean)
                            {
                                cellValue = (cellValue != "0").ToString();
                            }
                            else if (dataType == XLCellValues.TimeSpan)
                            {
                                cellValue = TimeSpan.Parse(cellValue).ToString();
                            }
                            else if (dataType == XLCellValues.Number)
                            {
                                String format;
                                if (Style.NumberFormat.NumberFormatId > 0)
                                    format = formatCodes[Style.NumberFormat.NumberFormatId];
                                else
                                    format = Style.NumberFormat.Format;
                                cellValue = Double.Parse(cellValue).ToString(format);
                            }
                            else if (dataType == XLCellValues.DateTime)
                            {
                                String format;
                                if (Style.NumberFormat.NumberFormatId > 0)
                                    format = formatCodes[Style.NumberFormat.NumberFormatId];
                                else
                                    format = Style.NumberFormat.Format;
                                cellValue = DateTime.FromOADate(Double.Parse(cellValue)).ToString(format);
                            }
                        }
                    }
                    dataType = value;
                }
            }
        }

        public void Clear()
        {
            worksheet.Range(Address, Address).Clear();
        }
        public void ClearStyles()
        {
            var newStyle = new XLStyle(this, worksheet.Style);
            newStyle.NumberFormat = this.Style.NumberFormat;
            this.Style = newStyle;
        }
        public void Delete(XLShiftDeletedCells shiftDeleteCells)
        {
            worksheet.Range(Address, Address).Delete(shiftDeleteCells);
        }

        private static Dictionary<Int32, String> formatCodes;
        private static Dictionary<Int32, String> GetFormatCodes()
        {
            if (formatCodes == null)
            {
                var fCodes = new Dictionary<Int32, String>();
                fCodes.Add(0, "");
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
                formatCodes = fCodes;
            }
            return formatCodes;
        }

        private String formulaA1;
        public String FormulaA1
        {
            get 
            {
                if (StringExtensions.IsNullOrWhiteSpace(formulaA1))
                    return String.Empty;
                else if (formulaA1.Trim()[0] == '=')
                    return formulaA1.Substring(1);
                else
                    return formulaA1;
            }
            set 
            { 
                formulaA1 = value;
                formulaR1C1 = String.Empty;
            }
        }

        private String formulaR1C1;
        public String FormulaR1C1
        {
            get 
            {
                if (StringExtensions.IsNullOrWhiteSpace(formulaR1C1))
                    formulaR1C1 = GetFormulaR1C1(FormulaA1);

                return formulaR1C1; 
            }
            set 
            { 
                formulaR1C1 = value;
                FormulaA1 = GetFormulaA1(value);
            }
        }

        private String GetFormulaR1C1(String value)
        {
            return GetFormula(value, FormulaConversionType.A1toR1C1, 0, 0);
        }

        private String GetFormulaA1(String value)
        {
            return GetFormula(value, FormulaConversionType.R1C1toA1, 0, 0);
        }

        private enum FormulaConversionType { A1toR1C1, R1C1toA1 };
        private static Regex a1Regex = new Regex(
            @"(?<=\W)(\$?[a-zA-Z]{1,3}\$?\d{1,7})(?=\W)" // A1
            + @"|(?<=\W)(\d{1,7}:\d{1,7})(?=\W)" // 1:1
            + @"|(?<=\W)([a-zA-Z]{1,3}:[a-zA-Z]{1,3})(?=\W)"); // A:A

        private static Regex r1c1Regex = new Regex(
               @"(?<=\W)([Rr]\[?-?\d{0,7}\]?[Cc]\[?-?\d{0,7}\]?)(?=\W)" // R1C1
            + @"|(?<=\W)([Rr]\[?-?\d{0,7}\]?:[Rr]\[?-?\d{0,7}\]?)(?=\W)" // R:R
            + @"|(?<=\W)([Cc]\[?-?\d{0,5}\]?:[Cc]\[?-?\d{0,5}\]?)(?=\W)"); // C:C
        private String GetFormula(String strValue, FormulaConversionType conversionType, Int32 rowsToShift, Int32 columnsToShift)
        {
            if (StringExtensions.IsNullOrWhiteSpace(strValue))
                return String.Empty;

            var value = ">" + strValue + "<";

            Regex regex = conversionType == FormulaConversionType.A1toR1C1 ? a1Regex : r1c1Regex;

            var sb = new StringBuilder();
            var lastIndex = 0;

            foreach (var match in regex.Matches(value).Cast<Match>())
            {
                var matchString = match.Value;
                var matchIndex = match.Index;
                if (value.Substring(0, matchIndex).CharCount('"') % 2 == 0) // Check if the match is in between quotes
                {
                    sb.Append(value.Substring(lastIndex, matchIndex - lastIndex));
                    if (conversionType == FormulaConversionType.A1toR1C1)
                        sb.Append(GetR1C1Address(matchString, rowsToShift, columnsToShift));
                    else
                        sb.Append(GetA1Address(matchString, rowsToShift, columnsToShift));
                }
                else
                {
                    sb.Append(value.Substring(lastIndex, matchIndex - lastIndex + matchString.Length));
                }
                lastIndex = matchIndex + matchString.Length;
            }
            if (lastIndex < value.Length)
                sb.Append(value.Substring(lastIndex));

            var retVal = sb.ToString();
            return retVal.Substring(1, retVal.Length - 2);
        }

        private String GetA1Address(String r1c1Address, Int32 rowsToShift, Int32 columnsToShift)
        {
            var addressToUse = r1c1Address.ToUpper();

            if (addressToUse.Contains(':'))
            {
                var parts = addressToUse.Split(':');
                var p1 = parts[0];
                var p2 = parts[1];
                String leftPart;
                String rightPart;
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

                var rowPart = addressToUse.Substring(0, addressToUse.IndexOf("C"));
                String rowToReturn = GetA1Row(rowPart, rowsToShift);

                var columnPart = addressToUse.Substring(addressToUse.IndexOf("C"));
                String columnToReturn = GetA1Column(columnPart, columnsToShift);

                var retAddress = columnToReturn + rowToReturn;
                return retAddress;
            }
        }

        private String GetA1Column(String columnPart, Int32 columnsToShift)
        {
            String columnToReturn;
            if (columnPart == "C")
            {
                columnToReturn = XLAddress.GetColumnLetterFromNumber(Address.ColumnNumber + columnsToShift);
            }
            else
            {
                var bIndex = columnPart.IndexOf("[");
                if (bIndex >= 0)
                    columnToReturn = XLAddress.GetColumnLetterFromNumber(
                        Address.ColumnNumber + Int32.Parse(columnPart.Substring(bIndex + 1, columnPart.Length - bIndex - 2)) + columnsToShift
                        );
                else
                    columnToReturn = "$" + XLAddress.GetColumnLetterFromNumber(Int32.Parse(columnPart.Substring(1)) + columnsToShift);
            }
            return columnToReturn;
        }

        private String GetA1Row(String rowPart, Int32 rowsToShift)
        {
            String rowToReturn;
            if (rowPart == "R")
            {
                rowToReturn = (Address.RowNumber + rowsToShift).ToString();
            }
            else
            {
                var bIndex = rowPart.IndexOf("[");
                if (bIndex >= 0)
                    rowToReturn = (Address.RowNumber + Int32.Parse(rowPart.Substring(bIndex + 1, rowPart.Length - bIndex - 2)) + rowsToShift).ToString();
                else
                    rowToReturn = "$" + (Int32.Parse(rowPart.Substring(1)) + rowsToShift).ToString();
            }
            return rowToReturn;
        }

        private String GetR1C1Address(String a1Address, Int32 rowsToShift, Int32 columnsToShift)
        {
            if (a1Address.Contains(':'))
            {
                var parts = a1Address.Split(':');
                var p1 = parts[0];
                var p2 = parts[1];
                Int32 row1;
                if (Int32.TryParse(p1.Replace("$", ""), out row1))
                {
                    var row2 = Int32.Parse(p2.Replace("$", ""));
                    var leftPart = GetR1C1Row(row1, p1.Contains('$'), rowsToShift);
                    var rightPart = GetR1C1Row(row2, p2.Contains('$'), rowsToShift);
                    return leftPart + ":" + rightPart;
                }
                else
                {
                    var column1 = XLAddress.GetColumnNumberFromLetter(p1.Replace("$", ""));
                    var column2 = XLAddress.GetColumnNumberFromLetter(p2.Replace("$", ""));
                    var leftPart = GetR1C1Column(column1, p1.Contains('$'), columnsToShift);
                    var rightPart = GetR1C1Column(column2, p2.Contains('$'), columnsToShift);
                    return leftPart + ":" + rightPart;
                }
            }
            else
            {
                var address = new XLAddress(a1Address);

                String rowPart = GetR1C1Row(address.RowNumber, address.FixedRow, rowsToShift);
                String columnPart = GetR1C1Column(address.ColumnNumber, address.FixedRow, columnsToShift);

                return rowPart + columnPart;
            }
        }

        private String GetR1C1Row(Int32 rowNumber, Boolean fixedRow, Int32 rowsToShift)
        {
            String rowPart;
            rowNumber += rowsToShift;
            var rowDiff = rowNumber - Address.RowNumber;
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

        private String GetR1C1Column(Int32 columnNumber, Boolean fixedColumn, Int32 columnsToShift)
        {
            String columnPart;
            columnNumber += columnsToShift;
            var columnDiff = columnNumber - Address.ColumnNumber;
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
            this.cellValue = source.cellValue;
            this.dataType = source.dataType;
            this.formulaA1 = source.formulaA1;
            this.formulaR1C1 = source.formulaR1C1;
        }

        internal void ShiftFormula(Int32 rowsToShift, Int32 columnsToShift)
        {
            if (!StringExtensions.IsNullOrWhiteSpace(formulaA1))
                FormulaR1C1 = GetFormula(formulaA1, FormulaConversionType.A1toR1C1, rowsToShift, columnsToShift);
            else if (!StringExtensions.IsNullOrWhiteSpace(formulaA1))
                FormulaA1 = GetFormula(formulaA1, FormulaConversionType.R1C1toA1, rowsToShift, columnsToShift);
        }

        public Boolean ShareString { get; set; }
    }
}
