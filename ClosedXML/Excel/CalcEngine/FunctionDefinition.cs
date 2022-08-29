using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Function definition class (keeps function name, parameter counts, and delegate).
    /// </summary>
    internal class FunctionDefinition
    {
        public FunctionDefinition(string name, int minParams, int maxParams, LegacyCalcEngineFunction function, AllowRange allowRanges, IReadOnlyCollection<int> markedParams)
        {
            if (allowRanges == AllowRange.None && markedParams.Any())
                throw new ArgumentException(nameof(markedParams));

            Name = name;
            MinParams = minParams;
            MaxParams = maxParams;
            AllowRanges = allowRanges;
            MarkedParams = markedParams;
            LegacyFunction = function;
        }

        public string Name { get; }

        public int MinParams { get; }

        public int MaxParams { get; }

        public FunctionFlags Flags { get; }

        public AllowRange AllowRanges { get; }

        /// <summary>
        /// Which parameters of the function are marked. The values are indexes of the function parameters, starting from 0.
        /// Used to determine which arguments allow ranges and which don't.
        /// </summary>
        public IReadOnlyCollection<int> MarkedParams { get; }

        public LegacyCalcEngineFunction LegacyFunction { get; }
    }
}
