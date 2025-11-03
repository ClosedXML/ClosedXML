using ClosedXmlExample.Models;

namespace ClosedXmlExample.Services;

/// <summary>
/// Simple in-memory storage for Excel documents metadata
/// Thread-safe CRUD operations
/// </summary>
public static class DocumentStorage
{
    private static readonly List<Document> _documents = new();
    private static int _nextId = 1;
    private static readonly object _lock = new();

    /// <summary>
    /// Get all documents
    /// </summary>
    public static List<Document> GetAll()
    {
        lock (_lock)
        {
            return new List<Document>(_documents);
        }
    }

    /// <summary>
    /// Get document by ID
    /// </summary>
    public static Document? GetById(int id)
    {
        lock (_lock)
        {
            return _documents.FirstOrDefault(d => d.Id == id);
        }
    }

    /// <summary>
    /// Get document by file name
    /// </summary>
    public static Document? GetByFileName(string fileName)
    {
        lock (_lock)
        {
            return _documents.FirstOrDefault(d => 
                string.Equals(d.FileName, fileName, StringComparison.OrdinalIgnoreCase));
        }
    }

    /// <summary>
    /// Get documents by author
    /// </summary>
    public static List<Document> GetByAuthor(string author)
    {
        lock (_lock)
        {
            return _documents.Where(d => 
                string.Equals(d.Author, author, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }
    }

    /// <summary>
    /// Create a new document
    /// </summary>
    public static Document Create(Document document)
    {
        lock (_lock)
        {
            document.Id = _nextId++;
            document.CreatedAt = DateTime.UtcNow;
            document.UpdatedAt = DateTime.UtcNow;
            
            _documents.Add(document);
            return document;
        }
    }

    /// <summary>
    /// Update an existing document
    /// </summary>
    public static Document Update(Document document)
    {
        lock (_lock)
        {
            var existingDocument = _documents.FirstOrDefault(d => d.Id == document.Id);
            if (existingDocument == null)
            {
                throw new ArgumentException($"Document with ID {document.Id} not found");
            }

            existingDocument.FileName = document.FileName;
            existingDocument.Description = document.Description;
            existingDocument.Author = document.Author;
            existingDocument.SheetCount = document.SheetCount;
            existingDocument.FileSize = document.FileSize;
            existingDocument.UpdatedAt = DateTime.UtcNow;
            
            return existingDocument;
        }
    }

    /// <summary>
    /// Update last accessed time
    /// </summary>
    public static void UpdateLastAccessed(int id)
    {
        lock (_lock)
        {
            var document = _documents.FirstOrDefault(d => d.Id == id);
            if (document != null)
            {
                document.LastAccessedAt = DateTime.UtcNow;
            }
        }
    }

    /// <summary>
    /// Delete a document
    /// </summary>
    public static bool Delete(int id)
    {
        lock (_lock)
        {
            var document = _documents.FirstOrDefault(d => d.Id == id);
            if (document == null)
            {
                return false;
            }

            _documents.Remove(document);
            return true;
        }
    }

    /// <summary>
    /// Search documents by any field
    /// </summary>
    public static List<Document> Search(string searchTerm)
    {
        if (string.IsNullOrWhiteSpace(searchTerm))
        {
            return GetAll();
        }

        lock (_lock)
        {
            return _documents.Where(d =>
                d.FileName.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ||
                (d.Description ?? "").Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ||
                (d.Author ?? "").Contains(searchTerm, StringComparison.OrdinalIgnoreCase)
            ).ToList();
        }
    }

    /// <summary>
    /// Get documents created within date range
    /// </summary>
    public static List<Document> GetByDateRange(DateTime startDate, DateTime endDate)
    {
        lock (_lock)
        {
            return _documents.Where(d => 
                d.CreatedAt >= startDate && d.CreatedAt <= endDate)
                .ToList();
        }
    }

    /// <summary>
    /// Get recently accessed documents
    /// </summary>
    public static List<Document> GetRecentlyAccessed(int count = 10)
    {
        lock (_lock)
        {
            return _documents
                .Where(d => d.LastAccessedAt.HasValue)
                .OrderByDescending(d => d.LastAccessedAt)
                .Take(count)
                .ToList();
        }
    }

    /// <summary>
    /// Clear all documents (for testing)
    /// </summary>
    public static void Clear()
    {
        lock (_lock)
        {
            _documents.Clear();
            _nextId = 1;
        }
    }

    /// <summary>
    /// Get total count
    /// </summary>
    public static int Count()
    {
        lock (_lock)
        {
            return _documents.Count;
        }
    }
}

