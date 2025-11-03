namespace ClosedXmlExample.Models;

/// <summary>
/// Represents an Excel document with metadata
/// </summary>
public class Document
{
    public int Id { get; set; }
    public string FileName { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Author { get; set; } = string.Empty;
    public int SheetCount { get; set; }
    public long FileSize { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    public DateTime? LastAccessedAt { get; set; }
    
    public string DisplayName => $"{FileName} ({SheetCount} sheets)";
    public string FileSizeFormatted => FileSize < 1024 ? $"{FileSize} B" 
        : FileSize < 1024 * 1024 ? $"{FileSize / 1024} KB" 
        : $"{FileSize / (1024 * 1024)} MB";
}

