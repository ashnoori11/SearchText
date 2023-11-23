using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using SharpCompress.Archives;
using SharpCompress.Common;
using System.Text;
using Xceed.Words.NET;

namespace SearchText.Services;

public class Search : IDisposable
{
    private bool disposedValue;

    public string SearchIntoFiles(string searchTerm, string RootPath, out int filesCount)
    {
        string startFolder = RootPath;
        System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(startFolder);

        var fileList = dir.GetFiles("*.*", System.IO.SearchOption.AllDirectories).ToList();
        filesCount = fileList.Count;

        var queryMatchingFiles =
            from file in fileList
            where file.Extension == ".htm" || file.Extension == ".txt" || file.Extension == ".html"
            || file.Extension == ".cs" || file.Extension == ".cshtml"
            let fileText = GetFileText(file.FullName)
            where fileText.Contains(searchTerm)
            select file.FullName;

        StringBuilder sb = new StringBuilder();
        foreach (string filename in queryMatchingFiles)
        {
            sb.Append($"{filename}   ");
        }

        var getZipFiles = from file in fileList
                          where file.Extension == ".zip"
                          select file.FullName;

        var searchResultIntoZipfiles = SearchIntoZipFiles(getZipFiles.ToList(), searchTerm);
        foreach (var filename in searchResultIntoZipfiles)
        {
            sb.Append($"{filename}   ");
        }

        var getDocxFiles = from file in fileList
                           where file.Extension == ".docx"
                           select file.FullName;

        var searchResultIntoDocxFiles = SearchIntoDocxFiles(getDocxFiles.ToList(), searchTerm);
        foreach (var filename in searchResultIntoDocxFiles)
        {
            sb.Append($"{filename}   ");
        }

        var getPdfFiles = from file in fileList
                          where file.Extension == ".pdf"
                          select file.FullName;

        var searchResultIntoPdfFiles = SearchIntoPdfFiles(getPdfFiles.ToList(), searchTerm);
        foreach (var filename in searchResultIntoPdfFiles)
        {
            sb.Append($"{filename}   ");
        }

        return sb.ToString().TrimStart().TrimEnd();
    }

    public List<string> SearchIntoDocxFiles(List<string> docxFilePaths, string searchTerm)
    {
        List<string> matchingDocxFiles = new List<string>();
        foreach (string docxFilePath in docxFilePaths)
        {
            string fileContent = ExtractTextFromDocx(docxFilePath);

            if (fileContent.Contains(searchTerm))
            {
                matchingDocxFiles.Add(docxFilePath);
            }
        }

        return matchingDocxFiles;
    }
    public List<string> SearchIntoZipFiles(List<string> zipFilePaths, string searchTerm)
    {
        List<string> matchingZipFiles = new List<string>();
        string extractedFilePath = string.Empty;
        string fileContent = string.Empty;

        foreach (string zipFilePath in zipFilePaths)
        {
            using (var archive = ArchiveFactory.Open(zipFilePath))
            {
                foreach (var entry in archive.Entries)
                {
                    if (!IsValidFileExtention(entry.Key))
                        continue;

                    if (entry.Key.EndsWith(".txt", StringComparison.OrdinalIgnoreCase) ||
                        entry.Key.EndsWith(".cs", StringComparison.OrdinalIgnoreCase) ||
                        entry.Key.EndsWith(".html", StringComparison.OrdinalIgnoreCase) ||
                        entry.Key.EndsWith(".cshtml", StringComparison.OrdinalIgnoreCase))
                    {

                        extractedFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());
                        entry.WriteToFile(extractedFilePath, new ExtractionOptions { ExtractFullPath = false, Overwrite = true });

                        fileContent = File.ReadAllText(extractedFilePath);
                        if (fileContent.Contains(searchTerm))
                        {
                            matchingZipFiles.Add(zipFilePath);
                        }

                        File.Delete(extractedFilePath);
                    }
                    else if (entry.Key.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                    {
                        extractedFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());
                        entry.WriteToFile(extractedFilePath, new ExtractionOptions { ExtractFullPath = false, Overwrite = true });

                        fileContent = ExtractTextFromDocx(extractedFilePath);
                        if (fileContent.Contains(searchTerm))
                        {
                            matchingZipFiles.Add(zipFilePath);
                        }

                        File.Delete(extractedFilePath);
                    }
                    else if (entry.Key.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    {
                        extractedFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());
                        entry.WriteToFile(extractedFilePath, new ExtractionOptions { ExtractFullPath = false, Overwrite = true });

                        fileContent = ExtractTextFromPdf(extractedFilePath);
                        if (fileContent.Contains(searchTerm))
                        {
                            matchingZipFiles.Add(zipFilePath);
                        }

                        File.Delete(extractedFilePath);
                    }
                }
            }
        }

        return matchingZipFiles;
    }
    public List<string> SearchIntoPdfFiles(List<string> pdfFilePaths, string searchTerm)
    {
        List<string> matchingPdfFiles = new List<string>();
        foreach (string pdfFilePath in pdfFilePaths)
        {
            string fileContent = ExtractTextFromPdf(pdfFilePath);
            if (fileContent.Contains(searchTerm))
            {
                matchingPdfFiles.Add(pdfFilePath);
            }
        }

        return matchingPdfFiles;
    }

    private static string GetFileText(string name)
    {
        string fileContents = String.Empty;
        if (System.IO.File.Exists(name))
        {
            fileContents = System.IO.File.ReadAllText(name);
        }
        return fileContents;
    }
    private bool IsValidFileExtention(string fileExtention)
    {
        if (fileExtention.EndsWith(".htm", StringComparison.OrdinalIgnoreCase) ||
            fileExtention.EndsWith(".txt", StringComparison.OrdinalIgnoreCase) ||
            fileExtention.EndsWith(".html", StringComparison.OrdinalIgnoreCase) ||
            fileExtention.EndsWith(".cs", StringComparison.OrdinalIgnoreCase) ||
            fileExtention.EndsWith(".cshtml", StringComparison.OrdinalIgnoreCase) ||
            fileExtention.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ||
            fileExtention.EndsWith(".js", StringComparison.OrdinalIgnoreCase) ||
            fileExtention.EndsWith(".css", StringComparison.OrdinalIgnoreCase) ||
            fileExtention.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
        else
            return false;
    }
    private string ExtractTextFromDocxXceed(string docxFilePath)
    {
        using (DocX document = DocX.Load(docxFilePath))
        {
            return string.Join(" ", document.Paragraphs.Select(p => p.Text));
        }
    }
    private string ExtractTextFromDocx(string docxFilePath)
    {
        using (WordprocessingDocument document = WordprocessingDocument.Open(docxFilePath, false))
        {
            var body = document.MainDocumentPart.Document.Body;
            var paragraphs = body.Elements<Paragraph>();
            return string.Join(" ", paragraphs.Select(p => p.InnerText));
        }
    }
    private string ExtractTextFromPdf(string pdfFilePath)
    {
        using (PdfReader reader = new PdfReader(pdfFilePath))
        {
            string text = string.Empty;
            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                text += PdfTextExtractor.GetTextFromPage(reader, i);
            }
            return text;
        }
    }

    #region Dispose Pattern
    protected virtual void Dispose(bool disposing)
    {
        if (!disposedValue)
        {
            if (disposing)
            {
                // TODO: dispose managed state (managed objects)
            }

            // TODO: free unmanaged resources (unmanaged objects) and override finalizer
            // TODO: set large fields to null
            disposedValue = true;
        }
    }

    // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
    // ~SearchService()
    // {
    //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
    //     Dispose(disposing: false);
    // }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}