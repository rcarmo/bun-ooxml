using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// Development-only readback for explicit cell XFs, not computed/rendered styles.
internal static class SpreadsheetStyleReader
{
    internal static object Read(SpreadsheetDocument document)
    {
        var workbook = document.WorkbookPart ?? throw new InvalidDataException("styles: missing workbook");
        var styles = workbook.WorkbookStylesPart?.Stylesheet ?? throw new InvalidDataException("styles: missing stylesheet");
        var formats = styles.CellFormats?.Elements<CellFormat>().ToArray() ?? [];
        var bases = styles.CellStyleFormats?.Elements<CellFormat>().ToArray() ?? [];
        var fonts = styles.Fonts?.Elements<Font>().Count() ?? 0;
        var fills = styles.Fills?.Elements<Fill>().Count() ?? 0;
        var borders = styles.Borders?.Elements<Border>().Count() ?? 0;
        var numberFormats = styles.NumberingFormats?.Elements<NumberingFormat>().ToArray() ?? [];
        var cells = new List<object>();
        void Dependency(uint index, int count, string kind)
        {
            if (index >= count) throw new InvalidDataException($"styles: {kind} index {index} out of range {count}");
        }
        void CheckFormat(CellFormat format)
        {
            Dependency(format.FontId?.Value ?? 0, fonts, "font");
            Dependency(format.FillId?.Value ?? 0, fills, "fill");
            Dependency(format.BorderId?.Value ?? 0, borders, "border");
            var id = format.NumberFormatId?.Value ?? 0;
            if (id >= 164 && numberFormats.Count(n => n.NumberFormatId?.Value == id && !string.IsNullOrEmpty(n.FormatCode?.Value)) != 1)
                throw new InvalidDataException($"styles: custom number format {id} is missing or ambiguous");
        }
        foreach (var sheet in workbook.Workbook?.Sheets?.Elements<Sheet>() ?? [])
        {
            if (sheet.Id?.Value is not string relationship)
                throw new InvalidDataException("styles: missing sheet relationship");
            if (workbook.GetPartById(relationship) is not WorksheetPart part)
                throw new InvalidDataException("styles: only worksheets supported by this reader");
            foreach (var cell in (part.Worksheet ?? throw new InvalidDataException("styles: missing worksheet")).Descendants<Cell>())
            {
                var index = cell.StyleIndex?.Value ?? 0;
                Dependency(index, formats.Length, "cell format");
                var format = formats[(int)index];
                CheckFormat(format);
                var baseIndex = format.FormatId?.Value ?? 0;
                Dependency(baseIndex, bases.Length, "base style");
                CheckFormat(bases[(int)baseIndex]);
                var value = cell.CellValue?.Text ?? "";
                if (cell.DataType?.Value == CellValues.InlineString) value = cell.InlineString?.InnerText ?? "";
                else if (cell.DataType?.Value == CellValues.SharedString)
                {
                    var strings = workbook.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToArray() ?? [];
                    if (!uint.TryParse(value, out var stringIndex) || stringIndex >= strings.Length)
                        throw new InvalidDataException("styles: invalid shared string reference");
                    value = strings[(int)stringIndex].InnerText;
                }
                cells.Add(new {
                    sheet = sheet.Name?.Value, address = cell.CellReference?.Value, value,
                    styleIndex = index, explicitStyleIndex = cell.StyleIndex != null,
                    wrapText = format.Alignment?.WrapText?.Value ?? false,
                    fontId = format.FontId?.Value ?? 0, fillId = format.FillId?.Value ?? 0,
                    borderId = format.BorderId?.Value ?? 0, numberFormatId = format.NumberFormatId?.Value ?? 0,
                    baseStyleIndex = baseIndex
                });
            }
        }
        return new { cellFormatCount = formats.Length, cells };
    }
}
