using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;

// Development oracle only. Never called from src/ or production operations.
if (args.Length == 0) throw new ArgumentException("Provide explicit DOCX/PPTX/XLSX paths");
var results = new List<object>();
var failed = false;
foreach (var input in args)
{
    var path = Path.GetFullPath(input);
    var before = Convert.ToHexStringLower(SHA256.HashData(File.ReadAllBytes(path)));
    object result;
    try
    {
        using OpenXmlPackage package = Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".docx" => WordprocessingDocument.Open(path, false),
            ".pptx" => PresentationDocument.Open(path, false),
            ".xlsx" => SpreadsheetDocument.Open(path, false),
            _ => throw new ArgumentException("Unsupported extension")
        };
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package)
            .Take(201).Select(e => new { e.Id, e.Description, part = e.Part?.Uri.ToString(), xpath = e.Path?.XPath }).ToArray();
        if (errors.Length > 0) failed = true;
        var properties = package.PackageProperties;
        var coreProperties = new {
            title = properties.Title, creator = properties.Creator, subject = properties.Subject,
            description = properties.Description, keywords = properties.Keywords, category = properties.Category,
            language = properties.Language, contentStatus = properties.ContentStatus, identifier = properties.Identifier,
            lastModifiedBy = properties.LastModifiedBy, revision = properties.Revision, version = properties.Version,
            created = properties.Created?.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"),
            modified = properties.Modified?.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"),
            lastPrinted = properties.LastPrinted?.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        };
        var spreadsheetStyles = package is SpreadsheetDocument spreadsheet ? SpreadsheetStyleReader.Read(spreadsheet) : null;
        result = new { file = Path.GetFileName(path), sha256 = before, errors, truncated = errors.Length > 200, coreProperties, spreadsheetStyles };
    }
    catch (Exception e)
    {
        failed = true;
        result = new { file = Path.GetFileName(path), sha256 = before, exception = e.Message };
    }
    if (before != Convert.ToHexStringLower(SHA256.HashData(File.ReadAllBytes(path))))
        throw new InvalidOperationException("Read-only validation changed an input");
    results.Add(result);
}
Console.WriteLine(JsonSerializer.Serialize(new
{
    schemaVersion = 1,
    validator = "DocumentFormat.OpenXml",
    version = typeof(OpenXmlElement).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion,
    profile = "Office2019",
    status = failed ? "failed" : "passed",
    results
}, new JsonSerializerOptions { WriteIndented = true }));
return failed ? 1 : 0;
