#nullable disable

using System;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel.CalcEngine;
using Array = ClosedXML.Excel.CalcEngine.Array;

namespace ClosedXML.Excel
{
    internal enum FormulaType : byte
    {
        Normal,
        Array,
        DataTable,
        Shared // Not used
    }

    internal enum FormulaConversionType
    {
        A1ToR1C1,
        R1C1ToA1
    };

    /// <summary>
    /// A representation of a cell formula, not the formula itself (i.e. the tree).
    /// </summary>
    [DebuggerDisplay("Cell:{Range} - Type: {Type} - Formula: {A1}")]
    internal sealed class XLCellFormula
    {
        /// <summary>
        /// This is only a placeholder, so the data table formula looks like array formula for saving code.
        /// First argument is replaced by value from current row, second is replaced by value from current column.
        /// </summary>
        private const string DataTableFormulaFormat = "{{TABLE({0},{1}}}";

        private static readonly Regex A1Regex = new(
            @"(?<=\W)(\$?[a-zA-Z]{1,3}\$?\d{1,7})(?=\W)" // A1
            + @"|(?<=\W)(\$?\d{1,7}:\$?\d{1,7})(?=\W)" // 1:1
            + @"|(?<=\W)(\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})(?=\W)", RegexOptions.Compiled); // A:A

        private static readonly Regex R1C1Regex = new(
            @"(?<=\W)([Rr](?:\[-?\d{0,7}\]|\d{0,7})?[Cc](?:\[-?\d{0,7}\]|\d{0,7})?)(?=\W)" // R1C1
            + @"|(?<=\W)([Rr]\[?-?\d{0,7}\]?:[Rr]\[?-?\d{0,7}\]?)(?=\W)" // R:R
            + @"|(?<=\W)([Cc]\[?-?\d{0,5}\]?:[Cc]\[?-?\d{0,5}\]?)(?=\W)", RegexOptions.Compiled); // C:C

        private XLSheetPoint _input1;
        private XLSheetPoint _input2;
        private FormulaFlags _flags;
        private FormulaType _type;

        /// <summary>
        /// The recalculation status of the last time formula was checked whether it needs to be recalculated.
        /// </summary>
        private RecalculationStatus _lastStatus;

        /// <summary>
        /// The value of <see cref="XLWorkbook.RecalculationCounter"/> that workbook had at the moment of cell formula evaluation.
        /// If this value equals to <see cref="XLWorkbook.RecalculationCounter"/> it indicates that <see cref="XLCell.CachedValue"/> stores
        /// correct value and no re-evaluation has to be performed.
        /// </summary>
        private long _evaluatedAtVersion;

        private XLCellFormula()
        {
            // Formulas from loaded worksheets with values are treated as valid, until something changes (i.e. recalculation version increases)
            // When a formula is set directly from a cell, it invalidates formula immediately afterwards.
            _lastStatus = new(false, 0);
        }

        internal bool IsEvaluating { get; private set; }

        /// <summary>
        /// Is this formula dirty, i.e. is it potentially out of date due to changes
        /// to precedent cells?
        /// </summary>
        internal bool IsDirty { get; set; }

        /// <summary>
        /// Formula in A1 notation. Either this or <see cref="R1C1"/> must be set (potentially
        /// both, due to conversion from one notation to another).
        /// </summary>
        private string A1 { get; set; }

        /// <summary>
        /// Formula in R1C1 notation. Either this or <see cref="A1"/> must be set (potentially
        /// both, due to conversion from one notation to another).
        /// </summary>
        private string R1C1 { get; set; }

        internal FormulaType Type => _type;

        /// <summary>
        /// Range for array and data table formulas, otherwise default value.
        /// </summary>
        /// <remarks>Doesn't contain sheet, so it doesn't have to deal with
        /// sheet renames and moving formula around.</remarks>
        internal XLSheetRange Range { get; set; }

        /// <summary>
        /// True, if 1D data table formula is the row (the displayed formula in Excel is missing the second argument <c>{=TABLE(A1;)}</c>).
        /// False the 1D data table is a column. (the displayed formula in Excel is missing the first argument <c>{=TABLE(;A1)}</c>)
        /// This property is meaningless, if called for non-data-table formula.
        /// </summary>
        /// <remarks>
        /// <para>
        /// If data table is in row (i.e. the value returns <c>true</c>) that means it calculates values in a row,
        /// it takes formula from a cell from a column one less than its range and replaces the input cell with value
        /// at the intersection of current cell column and the top row of the range. When data table is a column, it works
        /// pretty much same, except axis are reversed.
        /// </para>
        /// <para>
        /// Just because data table is 1D doesn't mean its range has to be. It can be rectangular even for 1D
        /// data table. It just means that data table is applied separately to each row/column (depending on whether
        /// the data table is row or column).
        /// </para>
        /// </remarks>
        internal Boolean IsRowDataTable => _flags.HasFlag(FormulaFlags.Is1DRow);

        /// <summary>
        /// True, if data table is 2D and uses both inputs. Input1 is replaced by
        /// value from current row, input2 is replaced by a value from current column.
        /// This property is meaningless, if called for non-data-table formula.
        /// </summary>
        internal Boolean Is2DDataTable => _flags.HasFlag(FormulaFlags.Is2D);

        /// <summary>
        /// Returns a cell that data table formula uses as a variable to replace with values
        /// for the actual table. Used for 1D data table formula as a single input (row or column)
        /// and as row for 2D data table. Must be present, even if input marked as deleted.
        /// This property is meaningless, if called for non-data-table formula.
        /// </summary>
        internal XLSheetPoint Input1 => _input1;

        /// <summary>
        /// Returns a cell that 2D data table formula uses as a variable to replace with values
        /// for the actual table. The value is taken from the top of range of the current column.
        /// Must be present for 2D, even if input marked as deleted.
        /// This property is meaningless, if called for non-data-table formula.
        /// </summary>
        internal XLSheetPoint Input2 => _input2;

        /// <summary>
        /// Returns true, if data table formula has its input1 deleted.
        /// This property is meaningless, if called for non-data-table formula.
        /// </summary>
        internal Boolean Input1Deleted => _flags.HasFlag(FormulaFlags.Input1Deleted);

        /// <summary>
        /// Returns true, if data table formula has its input1 deleted.
        /// This property is meaningless, if called for non-data-table formula.
        /// </summary>
        internal Boolean Input2Deleted => _flags.HasFlag(FormulaFlags.Input2Deleted);

        /// <summary>
        /// Flag indicating that previously calculated cell value may be not valid anymore and has to be re-evaluated.
        /// </summary>
        internal bool NeedsRecalculation(XLCell cell)
        {
            var worksheet = cell.Worksheet;
            var currentVersion = worksheet.Workbook.RecalculationCounter;

            // Nothing changed since last check => answer is still the same
            if (_lastStatus.Version == currentVersion)
            {
                return _lastStatus.NeededRecalculation;
            }

            var recalculationNeeded = EvaluateStatus(cell);
            _lastStatus = new(recalculationNeeded, currentVersion);
            return recalculationNeeded;
        }

        private bool EvaluateStatus(XLCell cell)
        {
            // Cell could have been invalidated or a new formula was set
            bool cellWasModified = _evaluatedAtVersion < cell.ModifiedAtVersion;
            if (cellWasModified)
            {
                return true;
            }

            var worksheet = cell.Worksheet;
            if (!worksheet.CalcEngine.TryGetPrecedentCells(A1, worksheet, out var precedentCells))
            {
                // If we are unable to even determine precedent cells, always recalculate
                return true;
            }

            foreach (var precedentCell in precedentCells!)
            {
                // the affecting cell was modified after this one was evaluated
                // e.g. cell now has a different value than it had at the last evaluation.
                if (precedentCell.ModifiedAtVersion > _evaluatedAtVersion)
                {
                    return true;
                }

                // the affecting cell was evaluated after this one (normally this should not happen)
                if (precedentCell.Formula is not null && precedentCell.Formula._evaluatedAtVersion > _evaluatedAtVersion)
                {
                    return true;
                }

                // the affecting cell needs recalculation (recursion to walk through dependencies)
                if (precedentCell.NeedsRecalculation)
                {
                    return true;
                }
            }

            return false;
        }

        internal void Invalidate(XLWorksheet worksheet)
        {
            _lastStatus = new(true, worksheet.Workbook.RecalculationCounter);
        }

        /// <summary>
        /// Get stored formula in A1 notation. Returned formula doesn't contain equal sign.
        /// </summary>
        /// <param name="cellAddress">Address of the formula cell. Used to convert relative R1C1 to A1, if conversion is necessary.</param>
        public string GetFormulaA1(XLSheetPoint cellAddress)
        {
            if (String.IsNullOrWhiteSpace(A1))
                A1 = GetFormula(R1C1, FormulaConversionType.R1C1ToA1, cellAddress);

            if (A1.Trim()[0] == '=')
                return A1.Substring(1);

            return A1;
        }

        /// <summary>
        /// Get stored formula in R1C1 notation. Returned formula doesn't contain equal sign.
        /// </summary>
        public string GetFormulaR1C1(XLSheetPoint cellAddress)
        {
            if (String.IsNullOrWhiteSpace(R1C1))
            {
                var normalizedA1 = GetFormulaA1(cellAddress);
                R1C1 = GetFormula(normalizedA1, FormulaConversionType.A1ToR1C1, cellAddress);
            }

            return R1C1;
        }

        internal static string GetFormula(string strValue, FormulaConversionType conversionType, XLSheetPoint cellAddress)
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
                        ? GetR1C1Address(matchString, cellAddress)
                        : GetA1Address(matchString, cellAddress));
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

        private static string GetA1Address(string r1C1Address, XLSheetPoint cellAddress)
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
                    leftPart = GetA1Row(p1, cellAddress.Row);
                    rightPart = GetA1Row(p2, cellAddress.Row);
                }
                else
                {
                    leftPart = GetA1Column(p1, cellAddress.Column);
                    rightPart = GetA1Column(p2, cellAddress.Column);
                }

                return leftPart + ":" + rightPart;
            }

            try
            {
                var rowPart = addressToUse.Substring(0, addressToUse.IndexOf('C'));
                var rowToReturn = GetA1Row(rowPart, cellAddress.Row);

                var columnPart = addressToUse.Substring(addressToUse.IndexOf('C'));
                var columnToReturn = GetA1Column(columnPart, cellAddress.Column);

                var retAddress = columnToReturn + rowToReturn;
                return retAddress;
            }
            catch (ArgumentOutOfRangeException)
            {
                return "#REF!";
            }
        }

        private static string GetA1Column(string columnPartRC, int cellColumn)
        {
            string columnToReturn;
            if (columnPartRC == "C")
                columnToReturn = XLHelper.GetColumnLetterFromNumber(cellColumn);
            else
            {
                var bIndex = columnPartRC.IndexOf('[');
                var mIndex = columnPartRC.IndexOf('-');
                if (bIndex >= 0)
                {
                    columnToReturn = XLHelper.GetColumnLetterFromNumber(
                        cellColumn +
                        Int32.Parse(columnPartRC.Substring(bIndex + 1, columnPartRC.Length - bIndex - 2))
                    );
                }
                else if (mIndex >= 0)
                {
                    columnToReturn = XLHelper.GetColumnLetterFromNumber(
                        cellColumn + Int32.Parse(columnPartRC.Substring(mIndex))
                    );
                }
                else
                {
                    columnToReturn = "$" +
                                     XLHelper.GetColumnLetterFromNumber(Int32.Parse(columnPartRC.Substring(1)));
                }
            }

            return columnToReturn;
        }

        private static string GetA1Row(string rowPartRC, int cellRow)
        {
            string rowToReturn;
            if (rowPartRC == "R")
                rowToReturn = cellRow.ToString();
            else
            {
                var bIndex = rowPartRC.IndexOf('[');
                if (bIndex >= 0)
                {
                    rowToReturn =
                        (cellRow + Int32.Parse(rowPartRC.Substring(bIndex + 1, rowPartRC.Length - bIndex - 2))).ToString();
                }
                else
                    rowToReturn = "$" + (Int32.Parse(rowPartRC.Substring(1)));
            }

            return rowToReturn;
        }

        private static string GetR1C1Address(string a1Address, XLSheetPoint cellAddress)
        {
            if (a1Address.Contains(':'))
            {
                var parts = a1Address.Split(':');
                var p1 = parts[0];
                var p2 = parts[1];
                if (Int32.TryParse(p1.Replace("$", string.Empty), out Int32 row1))
                {
                    var row2 = Int32.Parse(p2.Replace("$", string.Empty));
                    var leftPart = GetR1C1Row(row1, p1.Contains('$'), cellAddress.Row);
                    var rightPart = GetR1C1Row(row2, p2.Contains('$'), cellAddress.Row);
                    return leftPart + ":" + rightPart;
                }
                else
                {
                    var column1 = XLHelper.GetColumnNumberFromLetter(p1.Replace("$", string.Empty));
                    var column2 = XLHelper.GetColumnNumberFromLetter(p2.Replace("$", string.Empty));
                    var leftPart = GetR1C1Column(column1, p1.Contains('$'), cellAddress.Column);
                    var rightPart = GetR1C1Column(column2, p2.Contains('$'), cellAddress.Column);
                    return leftPart + ":" + rightPart;
                }
            }

            var address = XLAddress.Create(a1Address);

            var rowPart = GetR1C1Row(address.RowNumber, address.FixedRow, cellAddress.Row);
            var columnPart = GetR1C1Column(address.ColumnNumber, address.FixedColumn, cellAddress.Column);

            return rowPart + columnPart;
        }

        private static string GetR1C1Row(int rowNumber, bool fixedRow, int cellRow)
        {
            string rowPart;
            var rowDiff = rowNumber - cellRow;
            if (rowDiff != 0 || fixedRow)
                rowPart = fixedRow ? "R" + rowNumber : "R[" + rowDiff + "]";
            else
                rowPart = "R";

            return rowPart;
        }

        private static string GetR1C1Column(int columnNumber, bool fixedColumn, int cellColumn)
        {
            string columnPart;
            var columnDiff = columnNumber - cellColumn;
            if (columnDiff != 0 || fixedColumn)
                columnPart = fixedColumn ? "C" + columnNumber : "C[" + columnDiff + "]";
            else
                columnPart = "C";

            return columnPart;
        }

        internal void ApplyFormula(XLCell cell)
        {
            if (IsEvaluating)
            {
                throw new InvalidOperationException($"Cell {cell.Address} is a part of circular reference.");
            }

            try
            {
                IsEvaluating = true;

                if (Type == FormulaType.Normal)
                {
                    var fA1 = GetFormulaA1(cell.SheetPoint);
                    var result = CalculateNormalFormula(fA1, cell);
                    cell.SetOnlyValue(result.ToCellValue());
                }
                else if (Type == FormulaType.Array)
                {
                    var result = CalculateArrayFormula(cell);
                    for (var rowIdx = 0; rowIdx < result.Height; ++rowIdx)
                    {
                        for (var colIdx = 0; colIdx < result.Width; ++colIdx)
                        {
                            var cellValue = result[rowIdx, colIdx];
                            var cellRow = Range.FirstPoint.Row + rowIdx;
                            var cellColumn = Range.FirstPoint.Column + colIdx;
                            cell.Worksheet.Cell(cellRow, cellColumn).SetOnlyValue(cellValue.ToCellValue());
                        }
                    }
                }
                else
                {
                    throw new NotImplementedException($"Evaluation of {Type} formula not implemented.");
                }

                var recalculationCounter = cell.Worksheet.Workbook.RecalculationCounter;
                _lastStatus = new RecalculationStatus(false, recalculationCounter);
                _evaluatedAtVersion = recalculationCounter;
            }
            finally
            {
                IsEvaluating = false;
            }
        }

        /// <summary>
        /// Calculate a value of the specified formula.
        /// </summary>
        /// <param name="fA1">Cell formula to evaluate in A1 format.</param>
        /// <param name="cell">Cell whose formula is being evaluated.</param>
        private ScalarValue CalculateNormalFormula(string fA1, XLCell cell)
        {
            var worksheet = cell.Worksheet;
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
                sName = worksheet.Name;
                cAddress = fA1;
            }

            if (worksheet.Workbook.Worksheets.Contains(sName)
                && XLHelper.IsValidA1Address(cAddress))
            {
                var referenceCell = worksheet.Workbook.Worksheet(sName).Cell(cAddress);
                if (referenceCell.IsEmpty(XLCellsUsedOptions.AllContents))
                    return 0d;
                else
                    return referenceCell.Value;
            }

            if (worksheet.Workbook.Worksheets.Contains(sName)
                && XLHelper.IsValidA1Address(cAddress))
            {
                var referenceCell = worksheet.Workbook.Worksheet(sName).Cell(cAddress);
                if (referenceCell.IsEmpty(XLCellsUsedOptions.AllContents))
                    return 0;
                else
                    return referenceCell.Value;
            }

            var retVal = worksheet.CalcEngine.EvaluateFormula(fA1, worksheet.Workbook, worksheet, cell.Address);
            return retVal;
        }

        /// <summary>
        /// Calculate array formula and return an array of the array formula range size.
        /// </summary>
        private Array CalculateArrayFormula(XLCell masterCell)
        {
            var ws = masterCell.Worksheet;
            var formula = GetFormulaA1(masterCell.SheetPoint);
            var resultArray = ws.CalcEngine.EvaluateArrayFormula(formula, masterCell);
            var rescaledArray = resultArray.Rescale(Range.Height, Range.Width);
            return rescaledArray;
        }

        /// <summary>
        /// A factory method to create a normal A1 formula. Doesn't affect recalculation version.
        /// </summary>
        /// <param name="formulaA1">Doesn't start with <c>=</c>.</param>
        internal static XLCellFormula NormalA1(string formulaA1)
        {
            return new XLCellFormula
            {
                A1 = formulaA1,
                R1C1 = null,
                _type = FormulaType.Normal,
                _flags = FormulaFlags.None
            };
        }

        /// <summary>
        /// A factory method to create a normal R1C1 formula. Doesn't affect recalculation version.
        /// </summary>
        /// <param name="formulaR1C1">Doesn't start with <c>=</c>.</param>
        internal static XLCellFormula NormalR1C1(string formulaR1C1)
        {
            return new XLCellFormula
            {
                A1 = null,
                R1C1 = formulaR1C1,
                _type = FormulaType.Normal,
                _flags = FormulaFlags.None
            };
        }

        /// <summary>
        /// A factory method to create an array formula. Doesn't affect recalculation version.
        /// </summary>
        /// <param name="arrayFormulaA1">Isn't wrapped in <c>{}</c> and doesn't start with <c>=</c>.</param>
        /// <param name="range">A range of cells that are calculated through the array formula.</param>
        /// <param name="aca">A flag for always calculate array.</param>
        internal static XLCellFormula Array(string arrayFormulaA1, XLSheetRange range, bool aca)
        {
            return new XLCellFormula
            {
                A1 = arrayFormulaA1,
                R1C1 = null,
                _type = FormulaType.Array,
                _flags = aca ? FormulaFlags.AlwaysCalculateArray : FormulaFlags.None,
                Range = range
            };
        }

        /// <summary>
        /// A factory method to create a cell formula for 1D data table formula. Doesn't affect recalculation version.
        /// </summary>
        /// <param name="range">Range of the data table formula. Even 1D table can have rectangular range.</param>
        /// <param name="input1Address">Address of the input cell that will be replaced in the data table. If input deleted, ignored and value can be anything.</param>
        /// <param name="input1Deleted">Was the original address deleted?</param>
        /// <param name="isRowDataTable">Is data table in row (<c>true</c>) or columns (<c>false</c>)?</param>
        internal static XLCellFormula DataTable1D(
            XLSheetRange range,
            XLSheetPoint input1Address,
            bool input1Deleted,
            bool isRowDataTable)
        {
            String rowInput;
            String colInput;
            if (isRowDataTable)
            {
                colInput = string.Empty;
                rowInput = input1Deleted ? "#REF!" : input1Address.ToString();
            }
            else
            {
                colInput = input1Deleted ? "#REF!" : input1Address.ToString();
                rowInput = string.Empty;
            }

            return new XLCellFormula
            {
                A1 = string.Format(DataTableFormulaFormat, rowInput, colInput),
                R1C1 = null,
                Range = range,
                _type = FormulaType.DataTable,
                _input1 = input1Address,
                _flags =
                    (isRowDataTable ? FormulaFlags.Is1DRow : FormulaFlags.None) |
                    (input1Deleted ? FormulaFlags.Input1Deleted : FormulaFlags.None)
            };
        }

        /// <summary>
        /// A factory method to create a 2D data table formula. Doesn't affect recalculation version.
        /// </summary>
        /// <param name="range">Range of the formula.</param>
        /// <param name="input1Address">Address of the input cell that will be replaced in the data table. If input deleted, ignored and value can be anything.</param>
        /// <param name="input1Deleted">Was the original address deleted?</param>
        /// <param name="input2Address">Address of the input cell that will be replaced in the data table. If input deleted, ignored and value can be anything.</param>
        /// <param name="input2Deleted">Was the original address deleted?</param>
        internal static XLCellFormula DataTable2D(
            XLSheetRange range,
            XLSheetPoint input1Address,
            bool input1Deleted,
            XLSheetPoint input2Address,
            bool input2Deleted)
        {
            var colInput = input1Deleted ? "#REF!" : input1Address.ToString();
            var rowInput = input2Deleted ? "#REF!" : input2Address.ToString();
            return new XLCellFormula
            {
                A1 = string.Format(DataTableFormulaFormat, rowInput, colInput),
                R1C1 = null,
                Range = range,
                _type = FormulaType.DataTable,
                _input1 = input1Address,
                _input2 = input2Address,
                _flags = FormulaFlags.Is2D |
                (input1Deleted ? FormulaFlags.Input1Deleted : FormulaFlags.None) |
                (input2Deleted ? FormulaFlags.Input2Deleted : FormulaFlags.None)
            };
        }

        /// <summary>
        /// An enum to efficiently store various flags for formulas (bool takes up 1-4 bytes due to alignment).
        /// Note that each type of formula uses different flags.
        /// </summary>
        [Flags]
        private enum FormulaFlags : byte
        {
            None = 0,

            /// <summary>
            /// For Array formula. Not fully clear from documentation, but seems to be some kind of dirty flag.
            /// Current excel just writes <c>ca="1"</c> to each cell of array formula for cases described in the DOC.
            /// </summary>
            AlwaysCalculateArray = 1,

            /// <summary>
            /// For data table formula. Flag whether the data table is 2D and has two inputs.
            /// </summary>
            Is2D = 2,

            /// <summary>
            /// For data table formula. If the set, the data table is in row, not column. It uses input1 in both case, but the position
            /// is interpreted differently.
            /// </summary>
            Is1DRow = 4,

            /// <summary>
            /// For data table formula. When the input 1 cell has been deleted (not content, but the row or a column where cell was),
            /// this flag is set.
            /// </summary>
            Input1Deleted = 8,

            /// <summary>
            /// For data table formula. When the input 2 cell has been deleted (not content, but the row or a column where cell was),
            /// this flag is set.
            /// </summary>
            Input2Deleted = 16,
        }

        private readonly struct RecalculationStatus
        {
            public RecalculationStatus(bool neededRecalculation, long version)
            {
                NeededRecalculation = neededRecalculation;
                Version = version;
            }

            internal bool NeededRecalculation { get; }

            /// <summary>
            /// The value of <see cref="XLWorkbook.RecalculationCounter"/> that workbook had at the moment of determining whether the cell
            /// needs re-evaluation (due to it has been edited or some of the affecting cells has). If this value equals to <see cref="XLWorkbook.RecalculationCounter"/>
            /// it indicates that <see cref="NeededRecalculation"/> stores correct value and no check has to be performed.
            /// </summary>
            internal long Version { get; }
        };

        /// <summary>
        /// Get a lazy initialized AST for the formula.
        /// </summary>
        /// <param name="engine">Engine to parse the formula into AST, if necessary.</param>
        public Formula GetAst(CalcEngine.CalcEngine engine)
        {
            // TODO: Add caching for lazy initialization.
            var a1 = GetFormulaA1(Range.FirstPoint);
            var ast = engine.Parse(a1);
            return ast;
        }

        public override string ToString()
        {
            return A1 ?? R1C1;
        }
    }
}
