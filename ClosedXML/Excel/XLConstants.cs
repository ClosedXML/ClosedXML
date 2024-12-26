// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    //Use the class to store magic strings or variables.
    public static class XLConstants
    {
        internal const int NumberOfBuiltInStyles = 164; // But they are stored as 0-based (0 - 163)

        internal const int MaxFunctionArguments = 255; // To keep allocation sane

        internal const double ColumnWidthOffset = 0.710625;

        #region Pivot Table constants

        public static class PivotTable
        {
            internal const byte CreatedVersion = 5;
            internal const byte RefreshedVersion = 5;

            //TODO: Needs to be refactored to be more user-friendly.
            public const string ValuesSentinalLabel = "{{Values}}";
        }

        #endregion Pivot Table constants

        internal static class Comment
        {
            internal const string AlternateShapeTypeId = "_xssf_cell_comment";
            internal const string ShapeTypeId = "_x0000_t202";
        }

        /// <summary>
        /// Functions that are marked with a prefix <c>_xlfn</c> in formulas, but not GUI. Officially,
        /// they are called future functions.
        /// </summary>
        /// <remarks>
        /// Up to date for MS-XLSX 26.1 from 2024-12-11.
        /// </remarks>
        internal static readonly Lazy<IReadOnlyList<string>> FutureFunctions = new(() => new[]
        {
            "ACOT",
            "ACOTH",
            "AGGREGATE",
            "ARABIC",
            "BASE",
            "BETA.DIST",
            "BETA.INV",
            "BINOM.DIST",
            "BINOM.DIST.RANGE",
            "BINOM.INV",
            "BITAND",
            "BITLSHIFT",
            "BITOR",
            "BITRSHIFT",
            "BITXOR",
            "BYCOL",
            "BYROW",
            "CEILING.MATH",
            "CEILING.PRECISE",
            "CHISQ.DIST",
            "CHISQ.DIST.RT",
            "CHISQ.INV",
            "CHISQ.INV.RT",
            "CHISQ.TEST",
            "CHOOSECOLS",
            "CHOOSEROWS",
            "COMBINA",
            "CONCAT",
            "CONFIDENCE.NORM",
            "CONFIDENCE.T",
            "COT",
            "COTH",
            "COVARIANCE.P",
            "COVARIANCE.S",
            "CSC",
            "CSCH",
            "DAYS",
            "DECIMAL",
            "DROP",
            "ERF.PRECISE",
            "ERFC.PRECISE",
            "EXPAND",
            "EXPON.DIST",
            "F.DIST",
            "F.DIST.RT",
            "F.INV",
            "F.INV.RT",
            "F.TEST",
            "FIELDVALUE",
            "FILTERXML",
            "FLOOR.MATH",
            "FLOOR.PRECISE",
            "FORECAST.ETS",
            "FORECAST.ETS.CONFINT",
            "FORECAST.ETS.SEASONALITY",
            "FORECAST.ETS.STAT",
            "FORECAST.LINEAR",
            "FORMULATEXT",
            "GAMMA",
            "GAMMA.DIST",
            "GAMMA.INV",
            "GAMMALN.PRECISE",
            "GAUSS",
            "HSTACK",
            "HYPGEOM.DIST",
            "IFNA",
            "IFS",
            "IMCOSH",
            "IMCOT",
            "IMCSC",
            "IMCSCH",
            "IMSEC",
            "IMSECH",
            "IMSINH",
            "IMTAN",
            "ISFORMULA",
            "ISOMITTED",
            "ISOWEEKNUM",
            "LAMBDA",
            "LET",
            "LOGNORM.DIST",
            "LOGNORM.INV",
            "MAKEARRAY",
            "MAP",
            "MAXIFS",
            "MINIFS",
            "MODE.MULT",
            "MODE.SNGL",
            "MUNIT",
            "NEGBINOM.DIST",
            "NORM.DIST",
            "NORM.INV",
            "NORM.S.DIST",
            "NORM.S.INV",
            "NUMBERVALUE",
            "PDURATION",
            "PERCENTILE.EXC",
            "PERCENTILE.INC",
            "PERCENTRANK.EXC",
            "PERCENTRANK.INC",
            "PERMUTATIONA",
            "PHI",
            "POISSON.DIST",
            "PQSOURCE",
            "PYTHON_STR",
            "PYTHON_TYPE",
            "PYTHON_TYPENAME",
            "QUARTILE.EXC",
            "QUARTILE.INC",
            "QUERYSTRING",
            "RANDARRAY",
            "RANK.AVG",
            "RANK.EQ",
            "REDUCE",
            "RRI",
            "SCAN",
            "SEC",
            "SECH",
            "SEQUENCE",
            "SHEET",
            "SHEETS",
            "SKEW.P",
            "SORTBY",
            "STDEV.P",
            "STDEV.S",
            "SWITCH",
            "T.DIST",
            "T.DIST.2T",
            "T.DIST.RT",
            "T.INV",
            "T.INV.2T",
            "T.TEST",
            "TAKE",
            "TEXTAFTER",
            "TEXTBEFORE",
            "TEXTJOIN",
            "TEXTSPLIT",
            "TOCOL",
            "TOROW",
            "UNICHAR",
            "UNICODE",
            "UNIQUE",
            "VAR.P",
            "VAR.S",
            "VSTACK",
            "WEBSERVICE",
            "WEIBULL.DIST",
            "WRAPCOLS",
            "WRAPROWS",
            "XLOOKUP",
            "XOR",
            "Z.TEST",
        });

        /// <summary>
        /// Functions that are marked with a prefix <c>_xlfn._xlws</c> in formulas, but not GUI. They
        /// are part of the future functions. Unlike other future functions, they can only be used in
        /// a worksheet, but not a macro sheet. In the grammar, they are marked as
        /// <c>worksheet-only-function-list</c>.
        /// </summary>
        /// <remarks>
        /// Up to date for MS-XLSX 26.1 from 2024-12-11.
        /// </remarks>
        internal static readonly Lazy<IReadOnlyList<string>> WorksheetOnlyFunctions = new(() => new[]
        {
            "FILTER",
            "PY",
            "SORT",
        });

        /// <summary>
        /// Key: GUI name of future function, Value: prefixed name of future function. This doesn't
        /// include all future functions, only ones that need a hidden prefix (e.g. <c>ECMA.CEILING</c>
        /// is a future function without a prefix).
        /// </summary>
        internal static readonly Lazy<IReadOnlyDictionary<string, string>> FutureFunctionMap = new(() =>
        {
            var functionsMap = new Dictionary<string, string>(XLHelper.FunctionComparer);
            foreach (var futureFunction in XLConstants.FutureFunctions.Value)
                functionsMap.Add(futureFunction, "_xlfn." + futureFunction);

            foreach (var futureFunction in XLConstants.WorksheetOnlyFunctions.Value)
                functionsMap.Add(futureFunction, "_xlfn._xlws." + futureFunction);

            return functionsMap;
        });
    }
}
