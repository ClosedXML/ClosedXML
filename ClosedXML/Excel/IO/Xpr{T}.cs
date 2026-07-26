using ClosedXML.IO;

namespace ClosedXML.Excel.IO;

/// <summary>
/// XML Parse Result. The result of a parsed element, that may contain a value from Parse* method.
/// </summary>
/// <typeparam name="T">Type representing the result of parsed element.</typeparam>
internal readonly struct Xpr<T>
{
    /// <summary>
    /// Make a fail result.
    /// </summary>
    public Xpr()
    {
        IsSuccess = false;
        Value = default;
    }

    /// <summary>
    /// Make a successful result with a value.
    /// </summary>
    internal Xpr(T value)
        : this(value, true)
    {
    }

    private Xpr(T value, bool success)
    {
        IsSuccess = success;
        Value = value;
    }

    /// <summary>
    /// Value of the element. Throws when the Xpr is a failure.
    /// </summary>
    public T Value
    {
        get
        {
            if (IsFail)
                throw PartStructureException.ExpectedElementNotFound();

            return field!;
        }
    }

    public bool IsFail => !IsSuccess;

    public bool IsSuccess { get; }
}
