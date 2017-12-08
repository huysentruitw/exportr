# Exportr

[![Build status](https://ci.appveyor.com/api/projects/status/e580jiu1bif8sfvw/branch/master?svg=true)](https://ci.appveyor.com/project/huysentruitw/exportr/branch/master)

Classes and interfaces for building a worksheet based export. Contains an OpenXml implementation.

## Get it on [NuGet](https://www.nuget.org/packages/Exportr.OpenXml/)

    PM> Install-Package Exportr
    PM> Install-Package Exportr.OpenXml

## Example

### ApiController

```csharp
[Authorize]
[RoutePrefix("api/export")]
public class ExportController : ApiController
{
    [HttpGet]
    [Route("books")]
    public async Task<HttpResponseMessage> Books()
    {
        // Setup export
        var exporter = new FileStreamExporter(new ExcelDocumentFactory(), new LibraryExportTask(_dataContext));

        // Write export to response stream
        var fileName = exporter.GetFileName();
        if (fileName == null) return await NotFound().ExecuteAsync(CancellationToken.None).ConfigureAwait(false);
        var stream = new MemoryStream();

        exporter.ExportToStream(stream);
        
        stream.Position = 0;
        var response = new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StreamContent(stream)
            {
                Headers =
                {
                    ContentType = new MediaTypeHeaderValue("application/octet-stream"),
                    ContentDisposition = new ContentDispositionHeaderValue("attachment") { FileName = fileName }
                }
            }
        };

        return response;
    }
}
```

### BookExportTask

```csharp
public class LibraryExportTask : IExportTask
{
    public string Name { get; } = "Library";

    public IEnumerable<ISheetExportTask> EnumSheetExportTasks()
    {
        yield return BookSheetExportTask(); // Exports a sheet that contains the list of books in the library
    }
}
```

### BookSheetExportTask

```csharp
public class Book
{
    public Guid Id { get; set; }
    public string Author { get; set; }
    public string Title { get; set; }
}

public static class Library
{
    public static IEnumerable<Book> GetAllBooks()
    {
        yield return new Book { Id = ##BOOKID-1##, Author = "Karan Mahajan", Title = "The Association of Small Bombs" };
        yield return new Book { Id = ##BOOKID-2##, Author = "Ian McGuire", Title = "The North Water" };
        yield return new Book { Id = ##BOOKID-3##, Author = "Colson Whitehead", Title = "The Underground Railroad" };
    }
}

public class BookSheetExportTask : ISheetExportTask
{
    public string Name { get; } = "Books"; // The name of the sheet containing the book list

    public string[] GetColumnLabels() => new[] { "Id", "Author", "Title" };

    public IEnumerable<object[]> EnumRowData()
    {
        foreach (var book in Library.GetAllBooks())
        {
            yield return new object[] { book.Id, book.Author, book.Title };
        }
    }
}
```

## InlineSheetExportTask

A generic implementation that can be used to create a ISheetExportTask inline.

By using this class, the above example could be rewritten as:

```csharp
public class LibraryExportTask : IExportTask
{
    public string Name { get; } = "Library";

    public IEnumerable<ISheetExportTask> EnumSheetExportTasks()
    {
        yield return InlineSheetExportTask.SingleParse<Book, BookRowData>("Books", () => Library.GetAllBooks(), ParseBookEntities);
    }

    private static BookRowData ParseBookEntities(Book book)
        => new BookRowData
        {
            Id = book.Id,
            Author = book.Author,
            Title = book.Title
        };

    private class BookRowData
    {
        [Column("Id", Order = 0)]
        public Guid Id { get; set;}
        [Column("Author", Order = 1)]
        public string Author { get; set; }
        [Column("Title", Order = 2)]
        public string Title { get; set; }
    }
}
```
