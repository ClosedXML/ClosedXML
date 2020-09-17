using System;

namespace ClosedXML.Excel.CalcEngine.Exceptions
{
    /// <summary>
    /// This error is most often the result of specifying a
    /// mathematical operation with one or more cells that contain
    /// text.
    /// Corresponds to the #VALUE! error in Excel
    /// </summary>
    /// <seealso cref="ClosedXML.Excel.CalcEngine.Exceptions.CalcEngineException" />
    public class CellValueException : CalcEngineException
    {
        internal CellValueException()
            : base()
        { }

        internal CellValueException(string message)
            : base(message)
        { }

        internal CellValueException(string message, Exception innerException)
            : base(message, innerException)
        { }

        public static T CatchTypeConversionExceptions<T>(Func<T> func)
        {
            if (func is null)
            {
                throw new ArgumentNullException(nameof(func));
            }

            try
            {
                return func.Invoke();
            }
            catch (Exception ex)
            {
                switch (ex)
                {
                    case InvalidCastException _:
                    case FormatException _:
                    case OverflowException _:
                        throw new CellValueException("Unable to convert value to the desired type.", ex);
                    default:
                        throw;
                }
            }
        }
    }
}
