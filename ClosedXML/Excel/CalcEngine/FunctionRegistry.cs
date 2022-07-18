using System;
using System.Collections.Generic;
using System.Reflection;

namespace ClosedXML.Excel.CalcEngine
{
    internal class FunctionRegistry
    {
        private readonly Dictionary<string, FormulaFunction> _func = new(StringComparer.InvariantCultureIgnoreCase);
        private readonly Dictionary<string, FunctionDefinition> _legacyFunc = new(StringComparer.InvariantCultureIgnoreCase);

        public void RegisterFunction(string functionName, MethodInfo fn)
        {
            _func.Add(functionName, new FormulaFunction(fn, 0, int.MaxValue));
        }

        public void RegisterFunction(string functionName, MethodInfo fn, int parmMin, int parmMax)
        {
            _func.Add(functionName, new FormulaFunction(fn, parmMin, parmMax));
        }

        public bool TryGetFunc(string name, out FormulaFunction func)
        {
            return _func.TryGetValue(name, out func);
        }

        #region Legacy registration

        public void RegisterFunction(string functionName, int parmCount, CalcEngineFunction fn)
        {
            RegisterFunction(functionName, parmCount, parmCount, fn);
        }

        public void RegisterFunction(string functionName, int parmMin, int parmMax, CalcEngineFunction fn)
        {
            _legacyFunc.Add(functionName, new FunctionDefinition(parmMin, parmMax, fn));
        }

        public bool TryGetFunc(string name, out FunctionDefinition func)
        {
            return _legacyFunc.TryGetValue(name, out func);
        }

        #endregion
    }
}
