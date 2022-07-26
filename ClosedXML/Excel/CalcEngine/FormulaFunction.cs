using OneOf;
using System;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, string, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A representation of a function that can be called in a formula. The function must satisfy following requirements:
    /// <list type="bullet">
    ///     <item>Function must be static.</item>
    ///     <item>First argument must be <see cref="CalcContext"/>.</item>
    ///     <item>For functions supporting a variable number of arguments, the second parameter must be Span&lt;FormulaType&gt;> (with FormulaType being one of formula types). No further parameters are allowed.</item>
    ///     <item>For functions with a fixed number of arguments, each argument is a separate parameter of the function.
    ///         <list type="bullet">
    ///             <item>Argument can be any formula type.</item>
    ///             <item>Argument can be a OneOf combination of several distinct formula types (in correct order).</item>
    ///             <item>If the parameter is optional, it must be marked as nullable.</item>
    ///         </list>
    ///     </item>
    ///     <item>Return value must be AnyValue.</item>
    /// </list>
    /// <example>
    ///     static OneOf&lt;double, Error&gt; Asin(CalcContext ctx, double value);
    ///     static OneOf&lt;double, Error&gt; Sum(CalcContext ctx, Span&lt;AnyValue&gt; values);
    ///     static OneOf&lt;double, Error&gt; Day(CalcContext ctx, Span&lt;double, string&gt; values);
    ///     static OneOf&lt;double, Array&gt; Row(CalcContext ctx, Reference? reference);
    /// </example>
    /// </summary>
    internal class FormulaFunction
    {
        private static readonly Type[] ValueTypes = new[] { typeof(Logical), typeof(Number1), typeof(string), typeof(Error1), typeof(Array), typeof(Reference) };
        private readonly CalcEngineFunction _method;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="method"></param>
        /// <param name="parmMin">Minimum amount of parameters, useful only for variable number of parameters function.</param>
        /// <param name="parmMax">Minimum amount of parameters, useful only for variable number of parameters function.</param>
        public FormulaFunction(CalcEngineFunction method, int parmMin, int parmMax)
        {
            _method = method;
            ParmMin = parmMin;
            ParmMax = parmMax;
        }

        public int ParmMin { get; }
        public int ParmMax { get; }

        public AnyValue CallFunction(CalcContext ctx, params AnyValue?[] args)
        {
            return _method(ctx, args);

            /*
            var convertedArgs = new object[args.Length + 1];
            convertedArgs[0] = ctx;
            for (var argIdx = 0; argIdx < args.Length; ++argIdx)
            {
                var conversionResult = ConvertArgument<Number1>(args[argIdx]);
                if (!conversionResult.TryPickT0(out var convertedArg, out var error))
                {
                    return error;
                }

                convertedArgs[argIdx + 1] = convertedArg;
            }

            try
            {
                var val = _method.Invoke(null, convertedArgs);
                return (AnyValue)val;
            }
            catch (Exception ex)
            {
                throw ex;
            }*/
        }

        /// <summary>
        /// Convert argument to method parameter type.
        /// </summary>
        /// <typeparam name="TParamType">Type of parameter expected by function (=target of conversion).</typeparam>
        /// <param name="arg">Value of argument from formula evaluation.</param>
        /// <returns><paramref name="arg"/> converted to <typeparamref name="TParamType"/> or <c>#VALUE!</c>.</returns>
        private OneOf<TParamType, Error1> ConvertArgument<TParamType>(AnyValue arg)
        {
            // What is the type of actual value we want to pass to function.
            Type argType = arg.Match(
                logical => typeof(Logical),
                number => typeof(Number1),
                text => typeof(string),
                error => typeof(Error1),
                array => typeof(Array),
                reference => typeof(Reference));

            // What is the range of type we should convert the value to?
            Type paramType = typeof(TParamType);
            if (paramType == argType)
            {
                // No idea how to say - they are of same type so just use it as the target type without boxing.
                var castedArgValue = (TParamType)arg.Value;
                return OneOf<TParamType, Error1>.FromT0(castedArgValue);
            }

            // Parameter can be simply the type of value or can be OneOf<T0, T1,...> with the types in generic arguments.
            Type[] targetTypes = ValueTypes.Contains(paramType)
                ? new[] { paramType }
                : paramType.GetGenericArguments();

            var argTypeIndex = System.Array.IndexOf(targetTypes, argType);
            if (argTypeIndex >= 0)
            {
                var factoryMethod = paramType.GetMethod($"FromT{argTypeIndex}", BindingFlags.Static | BindingFlags.Public);
                var convertedArg = factoryMethod.Invoke(null, new object[] { arg.Value });
                return (TParamType)convertedArg;
            }

            // Well, we have to try conversion.
            throw new NotImplementedException("Conversion... like from number to OneOf<Logical, Text>? Which one to pick?");
        }
    }
}
