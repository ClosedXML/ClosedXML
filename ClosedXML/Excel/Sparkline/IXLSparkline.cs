using System;

namespace ClosedXML.Excel;

/// <summary>
/// A representation of a one sparkline of a <see cref="IXLSparklineGroup"/>. Sparkline is
/// visual representation (line, columns, win/loss) of a series of values from
/// the <see cref="SourceData"/> cells. The style and other properties are defined on
/// the <see cref="SparklineGroup"/>.
/// </summary>
public interface IXLSparkline
{
    /// <summary>
    /// A cell that contains the sparkline.
    /// </summary>
    IXLCell Location { get; set; }

    /// <summary>
    /// Get range of source data. The returned value is null if sparkline formula for source
    /// data doesn't represent an area in the current workbook (e.g. no longer exists due to
    /// deletion or it is a reference to a different workbook, a name that is a union of
    /// several disjoined cells ect).
    /// </summary>
    /// <remarks>
    /// If source data is a single sheet reference, it must contain cells from either a single
    /// row or a single column.
    /// </remarks>
    /// <exception cref="ArgumentException">Throws when trying to set a range is from different
    ///   worksheet than the sparkline.</exception>
    IXLRange? SourceData { get; set; }

    /// <summary>
    /// Sparkline group into which this sparkline belongs to.
    /// </summary>
    IXLSparklineGroup SparklineGroup { get; }

    /// <summary>
    /// Move sparkline from current cell to a different cell. If target cell has a sparkline,
    /// it will be replaced.
    /// </summary>
    /// <param name="value">New location of the sparkline.</param>
    /// <exception cref="ArgumentException">The <paramref name="value"/> is from different
    ///   worksheet than the sparkline.</exception>
    IXLSparkline SetLocation(IXLCell value);

    /// <summary>
    /// Change the the <see cref="SourceData"/> of sparkline. 
    /// </summary>
    /// <param name="value">The range that should be used as a source data for the sparkline.</param>
    /// <exception cref="ArgumentException">The <paramref name="value"/> is not a single row or
    ///   a single column.</exception>
    IXLSparkline SetSourceData(IXLRange? value);
}
