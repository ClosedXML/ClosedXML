using System;
using System.Collections.Generic;
using System.Diagnostics;
using ClosedXML.Excel.CalcEngine;
using ClosedXML.Excel.CalcEngine.Visitors;
using ClosedXML.Parser;

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

        private XLSheetPoint _input1;
        private XLSheetPoint _input2;
        private FormulaFlags _flags;

        /// <summary>
        /// Is this formula dirty, i.e. is it potentially out of date due to changes
        /// to precedent cells?
        /// </summary>
        internal bool IsDirty { get; set; }

        /// <summary>
        /// Formula in A1 notation. Doesn't start with <c>=</c> sign.
        /// </summary>
        internal string A1 { get; private set; }

        internal FormulaType Type { get; private init; }

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

        private XLCellFormula(string a1)
        {
            A1 = a1;
        }

        /// <summary>
        /// Get stored formula in R1C1 notation. Returned formula doesn't contain equal sign.
        /// </summary>
        public string GetFormulaR1C1(XLSheetPoint cellAddress)
        {
            return GetFormula(A1, FormulaConversionType.A1ToR1C1, cellAddress);
        }

        internal static string GetFormula(string strValue, FormulaConversionType conversionType, XLSheetPoint cellAddress)
        {
            if (String.IsNullOrWhiteSpace(strValue))
                return String.Empty;

            // Users and some producers might prefix formula with '=', but that is not a valid
            // formula, so strip and re-add if present.
            var formula = strValue.Trim();
            if (formula.StartsWith('='))
                formula = formula[1..];

            var converted = conversionType switch
            {
                FormulaConversionType.A1ToR1C1 => FormulaConverter.ToR1C1(formula, cellAddress.Row, cellAddress.Column),
                FormulaConversionType.R1C1ToA1 => FormulaConverter.ToA1(formula, cellAddress.Row, cellAddress.Column),
                _ => throw new NotSupportedException()
            };

            if (formula.Length != strValue.Length)
                converted = strValue[..^formula.Length] + converted;

            return converted;
        }

        /// <summary>
        /// A factory method to create a normal A1 formula. Doesn't affect recalculation version.
        /// </summary>
        /// <param name="formulaA1">Formula in A1 form. Shouldn't start with <c>=</c>.</param>
        internal static XLCellFormula NormalA1(string formulaA1)
        {
            return new XLCellFormula(formulaA1)
            {
                Type = FormulaType.Normal,
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
            return new XLCellFormula(arrayFormulaA1)
            {
                Type = FormulaType.Array,
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

            var formula = string.Format(DataTableFormulaFormat, rowInput, colInput);
            return new XLCellFormula(formula)
            {
                Range = range,
                Type = FormulaType.DataTable,
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
            var formula = string.Format(DataTableFormulaFormat, rowInput, colInput);
            return new XLCellFormula(formula)
            {
                Range = range,
                Type = FormulaType.DataTable,
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

        /// <summary>
        /// Get a lazy initialized AST for the formula.
        /// </summary>
        /// <param name="engine">Engine to parse the formula into AST, if necessary.</param>
        public Formula GetAst(XLCalcEngine engine)
        {
            var ast = engine.Parse(A1);
            return ast;
        }

        public override string ToString()
        {
            return A1;
        }

        public void RenameSheet(XLSheetPoint origin, string oldSheetName, string newSheetName)
        {
            var a1 = A1;
            var res = FormulaConverter.ModifyA1(a1, origin.Row, origin.Column, new RenameRefModVisitor
            {
                Sheets = new Dictionary<string, string?> { { oldSheetName, newSheetName } }
            });

            if (res != a1)
            {
                A1 = res;
                IsDirty = true;
            }
        }

        internal XLCellFormula GetMovedTo(XLSheetPoint origin, XLSheetPoint destination)
        {
            // I could in theory swap 1x1 array or dataTable, but not worth it in this path.
            if (Type != FormulaType.Normal)
                throw new InvalidOperationException("Can only swap normal formulas.");

            var originR1C1 = FormulaConverter.ToR1C1(A1, origin.Row, origin.Column);
            var targetA1 = FormulaConverter.ToA1(originR1C1, destination.Row, destination.Column);
            var targetFormula = NormalA1(targetA1);
            targetFormula.IsDirty = true;
            return targetFormula;
        }
    }
}
