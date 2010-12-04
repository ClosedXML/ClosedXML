using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections;
using System.Data;


namespace ClosedXML.Excel
{
    internal class XLCell : IXLCell
    {
        public static readonly DateTime baseDate = new DateTime(1899, 12, 30);
        XLWorksheet worksheet;
        public XLCell(IXLAddress address, IXLStyle defaultStyle, XLWorksheet worksheet)
        {
            this.Address = address;
            Style = defaultStyle;
            if (Style == null) Style = worksheet.Style;
            this.worksheet = worksheet;
        }

        public IXLAddress Address { get; private set; }
        public String InnerText
        {
            get { return cellValue; }
        }

        public T GetValue<T>() 
        {
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
            else if (dataType == XLCellValues.DateTime || IsDateFormat())
            {
                TimeSpan timeSpan;
                if (TimeSpan.TryParse(cellValue, out timeSpan))
                {
                    return timeSpan.ToString();
                }
                else
                {
                    String format = GetFormat();
                    return DateTime.FromOADate(Double.Parse(cellValue)).ToString(format);
                }
            }
            else if (dataType == XLCellValues.Number)
            {
                String format = GetFormat();
                return Double.Parse(cellValue).ToString(format);
            }
            else
            {
                return cellValue;
            }
        }

        private bool IsDateFormat()
        {
            return (dataType == XLCellValues.Number 
                && String.IsNullOrWhiteSpace(Style.NumberFormat.Format)
                && ((Style.NumberFormat.NumberFormatId >= 14
                    && Style.NumberFormat.NumberFormatId <= 22)
                || (Style.NumberFormat.NumberFormatId >= 45
                    && Style.NumberFormat.NumberFormatId <= 47)));
        }

        private String GetFormat()
        {
            String format;
            if (String.IsNullOrWhiteSpace(Style.NumberFormat.Format))
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

        private Boolean initialized = false;
        private String cellValue = String.Empty;
        public Object Value
        {
            get
            {
                var fA1 = FormulaA1;
                if (!String.IsNullOrWhiteSpace(fA1))
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
                        var sourceCell = asRange.Cell(ro, co);
                        var targetCell = worksheet.Cell(Address.RowNumber + ro - 1, Address.ColumnNumber + co - 1);
                        targetCell.Style = sourceCell.Style;
                        targetCell.DataType = sourceCell.DataType;
                        targetCell.Value = sourceCell.Value;
                        targetCell.FormulaA1 = sourceCell.FormulaA1;
                    }
                }
                var rangesToMerge = new List<IXLRange>();
                foreach (var merge in asRange.Worksheet.Internals.MergedCells)
                {
                    if (asRange.ContainsRange(merge))
                    {
                        var mergedRange = worksheet.Range(merge);
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
            if (asEnumerable != null && collectionObject.GetType() != typeof(String))
            {
                Int32 ro = Address.RowNumber;
                Int32 maxCo = 0;
                foreach (var m in asEnumerable)
                {
                    Int32 co = Address.ColumnNumber;

                    if (m.GetType().IsPrimitive || m.GetType() == typeof(String) || m.GetType() == typeof(DateTime))
                    {
                        SetValue(m, ro, co);
                    }
                    else if (m.GetType().IsArray)
                    {
                        dynamic arr = m;
                        foreach (var item in arr)
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
                return true;
            } 
            else
            {
                return false;
            }
        }

        private void ClearMerged(Int32 rowCount, Int32 columnCount)
        {
            List<String> mergeToDelete = new List<String>();
            foreach (var merge in worksheet.Internals.MergedCells)
            {
                var ma = new XLRangeAddress(merge);

                if (!( // See if the two ranges intersect...
                       ma.FirstAddress.ColumnNumber > Address.ColumnNumber + columnCount
                    || ma.LastAddress.ColumnNumber < Address.ColumnNumber
                    || ma.FirstAddress.RowNumber > Address.RowNumber + rowCount
                    || ma.LastAddress.RowNumber < Address.RowNumber
                    ))
                {
                    mergeToDelete.Add(merge);
                }
            }
            mergeToDelete.ForEach(m => worksheet.Internals.MergedCells.Remove(m));
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
            if (initialized)
            {
                if (dataType == XLCellValues.Boolean)
                {
                    if (Boolean.TryParse(val, out bTest))
                        val = bTest ? "1" : "0";
                    else if (!(val == "1" || val == "0"))
                        throw new ArgumentException("'" + val + "' is not a Boolean type.");
                }
                else if (dataType == XLCellValues.DateTime)
                {
                    if (DateTime.TryParse(val, out dtTest))
                    {
                        val = dtTest.ToOADate().ToString();
                    }
                    else if (!Double.TryParse(val, out dTest))
                    {
                        throw new ArgumentException("'" + val + "' is not a DateTime type.");
                    }

                    if (Style.NumberFormat.Format == String.Empty && Style.NumberFormat.NumberFormatId == 0)
                    {
                        if (dtTest.Date == dtTest)
                            Style.NumberFormat.NumberFormatId = 14;
                        else
                            Style.NumberFormat.NumberFormatId = 22;
                    }

                }
                else if (dataType == XLCellValues.Number)
                {
                    if (value is TimeSpan)
                    {
                        var timeSpan = (TimeSpan)value;
                        val = baseDate.Add(timeSpan).ToOADate().ToString();
                        if (Style.NumberFormat.Format == String.Empty && Style.NumberFormat.NumberFormatId == 0)
                            Style.NumberFormat.NumberFormatId = 46;
                    }
                    else
                    {
                        if (!Double.TryParse(val, out dTest))
                            throw new ArgumentException("'" + val + "' is not a Numeric type.");
                    }
                }
            }
            else
            {
                if (val.Length > 0 && val.Substring(0, 1) == "'")
                {
                    val = val.Substring(1, val.Length - 1);
                    dataType = XLCellValues.Text;
                }
                else if (value is TimeSpan)
                {
                    var timeSpan = (TimeSpan)value;
                    val = baseDate.Add(timeSpan).ToOADate().ToString();
                    dataType = XLCellValues.Number;
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
                initialized = true;
                if (cellValue.Length > 0)
                {
                    TimeSpan tsTest;
                    if (value == XLCellValues.Boolean)
                    {
                        Boolean bTest;
                        if (Boolean.TryParse(cellValue, out bTest))
                            cellValue = bTest ? "1" : "0";
                        else
                            cellValue = cellValue == "0" || String.IsNullOrEmpty(cellValue) ? "0" : "1";
                    }
                    else if (value == XLCellValues.DateTime && !TimeSpan.TryParse(cellValue, out tsTest))
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
                                if (((Int32)Double.Parse(cellValue)).ToString() == cellValue)
                                    Style.NumberFormat.NumberFormatId = 14;
                                else
                                    Style.NumberFormat.NumberFormatId = 22;
                    }
                    else if (value == XLCellValues.Number)
                    {
                        if (TimeSpan.TryParse(cellValue, out tsTest))
                        {
                            cellValue = baseDate.Add(tsTest).ToOADate().ToString();
                            if (Style.NumberFormat.Format == String.Empty && Style.NumberFormat.NumberFormatId == 0)
                                Style.NumberFormat.NumberFormatId = 46;
                        }
                        else
                        {
                            cellValue = Double.Parse(cellValue).ToString();
                        }
                    }
                    else
                    {
                        var formatCodes = GetFormatCodes();
                        if (dataType == XLCellValues.Boolean)
                        {
                            cellValue = (cellValue != "0").ToString();
                        }
                        else if (dataType == XLCellValues.Number)
                        {
                            if (Style.NumberFormat.NumberFormatId == 46)
                            {
                                cellValue = (DateTime.FromOADate(Double.Parse(cellValue)) - baseDate).ToString();
                            }
                            else
                            {
                                String format;
                                if (Style.NumberFormat.NumberFormatId > 0)
                                    format = formatCodes[Style.NumberFormat.NumberFormatId];
                                else
                                    format = Style.NumberFormat.Format;
                                cellValue = Double.Parse(cellValue).ToString(format);
                            }
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

        public void Clear()
        {
            worksheet.Range(Address, Address).Clear();
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
            get { return formulaA1; }
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
                if (String.IsNullOrWhiteSpace(formulaR1C1))
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
            return GetFormula(value, FormulaConversionType.A1toR1C1);
        }

        private String GetFormulaA1(String value)
        {
            return GetFormula(value, FormulaConversionType.R1C1toA1);
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
        private String GetFormula(String strValue, FormulaConversionType conversionType)
        {
            if (String.IsNullOrWhiteSpace(strValue))
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
                        sb.Append(GetR1C1Address(matchString));
                    else
                        sb.Append(GetA1Address(matchString));
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

        private String GetA1Address(String r1c1Address)
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
                    leftPart = GetA1Row(p1);
                    rightPart = GetA1Row(p2);
                }
                else
                {
                    leftPart = GetA1Column(p1);
                    rightPart = GetA1Column(p2);
                }
                return leftPart + ":" + rightPart;
            }
            else
            {

                var rowPart = addressToUse.Substring(0, addressToUse.IndexOf("C"));
                String rowToReturn = GetA1Row(rowPart);

                var columnPart = addressToUse.Substring(addressToUse.IndexOf("C"));
                String columnToReturn = GetA1Column(columnPart);
                //var cIndex = addressToUse.IndexOf("C");
                //String columnToReturn;
                //if (cIndex == addressToUse.Length - 1)
                //{
                //    columnToReturn = Address.ColumnLetter;
                //}
                //else
                //{
                //    var columnPart = addressToUse.Substring(cIndex);
                //    var bIndex = columnPart.IndexOf("[");
                //    if (bIndex >= 0)
                //        columnToReturn = XLAddress.GetColumnLetterFromNumber(
                //            Address.ColumnNumber + Int32.Parse(columnPart.Substring(bIndex + 1, columnPart.Length - bIndex - 2)));
                //    else
                //        columnToReturn = "$" + XLAddress.GetColumnLetterFromNumber(Int32.Parse(columnPart.Substring(1)));
                //}

                var retAddress = columnToReturn + rowToReturn;
                return retAddress;
            }
        }

        private String GetA1Column(String columnPart)
        {
            String columnToReturn;
            if (columnPart == "C")
            {
                columnToReturn = Address.ColumnLetter;
            }
            else
            {
                var bIndex = columnPart.IndexOf("[");
                if (bIndex >= 0)
                    columnToReturn = XLAddress.GetColumnLetterFromNumber(
                        Address.ColumnNumber + Int32.Parse(columnPart.Substring(bIndex + 1, columnPart.Length - bIndex - 2))
                        );
                else
                    columnToReturn = "$" + XLAddress.GetColumnLetterFromNumber(Int32.Parse(columnPart.Substring(1)));
            }
            return columnToReturn;
        }

        private String GetA1Row(String rowPart)
        {
            String rowToReturn;
            if (rowPart == "R")
            {
                rowToReturn = Address.RowNumber.ToString();
            }
            else
            {
                var bIndex = rowPart.IndexOf("[");
                if (bIndex >= 0)
                    rowToReturn = (Address.RowNumber + Int32.Parse(rowPart.Substring(bIndex + 1, rowPart.Length - bIndex - 2))).ToString();
                else
                    rowToReturn = "$" + rowPart.Substring(1);
            }
            return rowToReturn;
        }

        private String GetR1C1Address(String a1Address)
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
                    var leftPart = GetR1C1Row(row1, p1.Contains('$'));
                    var rightPart = GetR1C1Row(row2, p2.Contains('$'));
                    return leftPart + ":" + rightPart;
                }
                else
                {
                    var column1 = XLAddress.GetColumnNumberFromLetter(p1.Replace("$", ""));
                    var column2 = XLAddress.GetColumnNumberFromLetter(p2.Replace("$", ""));
                    var leftPart = GetR1C1Column(column1, p1.Contains('$'));
                    var rightPart = GetR1C1Column(column2, p2.Contains('$'));
                    return leftPart + ":" + rightPart;
                }
            }
            else
            {
                var address = new XLAddress(a1Address);

                String rowPart = GetR1C1Row(address.RowNumber, address.FixedRow);
                String columnPart = GetR1C1Column(address.ColumnNumber, address.FixedRow);

                return rowPart + columnPart;
            }
        }

        private String GetR1C1Row(Int32 rowNumber, Boolean fixedRow)
        {
            String rowPart;
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

        private String GetR1C1Column(Int32 columnNumber, Boolean fixedColumn)
        {
            String columnPart;
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

    }
}
