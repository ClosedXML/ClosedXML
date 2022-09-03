using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>Which parameters of a function allow ranges. That is important for implicit intersection.</summary>
    internal enum AllowRange
    {
        /// <summary>None of parameters allow ranges.</summary>
        None,

        /// <summary>All parameters allow ranges.</summary>
        All,

        /// <summary>All parameters except marked ones allow ranges.</summary>
        Except,

        /// <summary>Only marked parameters allow ranges.</summary>
        Only,
    }

    internal class FunctionRegistry
    {
        private readonly Dictionary<string, FunctionDefinition> _func = new(StringComparer.InvariantCultureIgnoreCase);

        public bool TryGetFunc(string name, out FunctionDefinition func)
        {
            return _func.TryGetValue(name, out func);
        }

        public void RegisterFunction(string functionName, int minParams, int maxParams, CalcEngineFunction fn, FunctionFlags flags, AllowRange allowRanges = AllowRange.None, params int[] markedParams)
        {
            _func.Add(functionName, new FunctionDefinition(minParams, maxParams, fn, flags, allowRanges, markedParams));
        }

        public void RegisterFunction(string functionName, int paramCount, LegacyCalcEngineFunction fn, AllowRange allowRanges = AllowRange.None, params int[] markedParams)
        {
            RegisterFunction(functionName, paramCount, paramCount, fn, allowRanges, markedParams);
        }

        public void RegisterFunction(string functionName, int minParams, int maxParams, LegacyCalcEngineFunction fn, AllowRange allowRanges = AllowRange.None, params int[] markedParams)
        {
            _func.Add(functionName, new FunctionDefinition(minParams, maxParams, fn, allowRanges, markedParams));
        }

        public bool TryGetFunc(string name, out int paramMin, out int paramMax)
        {
            if (_func.TryGetValue(name, out var func))
            {
                paramMin = func.MinParams;
                paramMax = func.MaxParams;
                return true;
            }

            paramMin = -1;
            paramMax = -1;
            return false;
        }
    }
}
