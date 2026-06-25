using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using KellyCashApp.Models;
using System.Globalization;
using System.Text.RegularExpressions;

namespace KellyCashApp.Services
{
    internal static class OirImporter
    {
        public static Dictionary<string, OirMatch> Load(string filePath)
        {
            var matches = new Dictionary<string, OirMatch>(15000, StringComparer.OrdinalIgnoreCase);

            using SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false);

            WorkbookPart workbookPart = document.WorkbookPart
                ?? throw new Exception("Could not read workbook.");

            SharedStringTable? sharedStrings =
                workbookPart.SharedStringTablePart?.SharedStringTable;

            Sheet firstSheet = workbookPart.Workbook.Sheets!
                .Elements<Sheet>()
                .First();

            WorksheetPart worksheetPart =
                (WorksheetPart)workbookPart.GetPartById(firstSheet.Id!);

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()
                ?? throw new Exception("Could not read worksheet data.");

            int headerRowNumber = -1;
            Dictionary<string, int> headers = new(StringComparer.OrdinalIgnoreCase);

            foreach (Row row in sheetData.Elements<Row>())
            {
                int rowNumber = (int)row.RowIndex!.Value;

                if (rowNumber > 25)
                    break;

                var rowValues = ReadRow(row, sharedStrings);

                if (rowValues.Any(x => x.Value.Equals("Consultant", StringComparison.OrdinalIgnoreCase)))
                {
                    headerRowNumber = rowNumber;

                    headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                    foreach (var item in rowValues)
                    {
                        string header = item.Value.Trim();

                        if (string.IsNullOrWhiteSpace(header))
                            continue;

                        if (!headers.ContainsKey(header))
                            headers.Add(header, item.Key);
                    }

                    break;
                }
            }

            if (headerRowNumber == -1)
                throw new Exception("Could not find Consultant header in Open Invoice Report.");

            int consultantColumn = GetRequiredColumn(headers, "Consultant");
            int serviceEndColumn = GetRequiredColumn(headers, "Service End");
            int documentNumberColumn = GetRequiredColumn(headers, "Document Number");
            int remainingAmountColumn = GetRequiredColumn(headers, "Remaining Amount");
            int clientProjectsColumn = GetRequiredColumn(headers, "Client Projects");

            foreach (Row row in sheetData.Elements<Row>())
            {
                int rowNumber = (int)row.RowIndex!.Value;

                if (rowNumber <= headerRowNumber)
                    continue;

                var rowValues = ReadRow(row, sharedStrings);

                string consultant = GetValue(rowValues, consultantColumn).Trim();
                string serviceEnd = FormatServiceEnd(GetValue(rowValues, serviceEndColumn));
                string documentNumber = GetValue(rowValues, documentNumberColumn).Trim();
                decimal remainingAmount = ParseDecimal(GetValue(rowValues, remainingAmountColumn));
                string clientProjects = GetValue(rowValues, clientProjectsColumn).Trim();

                if (string.IsNullOrWhiteSpace(consultant) || string.IsNullOrWhiteSpace(serviceEnd))
                    continue;

                string key = $"{consultant} {serviceEnd}".Trim();

                if (!matches.ContainsKey(key))
                    matches.Add(key, new OirMatch(documentNumber, remainingAmount, clientProjects));
            }

            return matches;
        }

        public static Dictionary<string, List<OirMatch>> LoadMultiple(string filePath)
        {
            var matches = new Dictionary<string, List<OirMatch>>(15000, StringComparer.OrdinalIgnoreCase);

            using SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false);

            WorkbookPart workbookPart = document.WorkbookPart
                ?? throw new Exception("Could not read workbook.");

            SharedStringTable? sharedStrings =
                workbookPart.SharedStringTablePart?.SharedStringTable;

            Sheet firstSheet = workbookPart.Workbook.Sheets!
                .Elements<Sheet>()
                .First();

            WorksheetPart worksheetPart =
                (WorksheetPart)workbookPart.GetPartById(firstSheet.Id!);

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()
                ?? throw new Exception("Could not read worksheet data.");

            int headerRowNumber = -1;
            Dictionary<string, int> headers = new(StringComparer.OrdinalIgnoreCase);

            foreach (Row row in sheetData.Elements<Row>())
            {
                int rowNumber = (int)row.RowIndex!.Value;

                if (rowNumber > 25)
                    break;

                var rowValues = ReadRow(row, sharedStrings);

                if (rowValues.Any(x => x.Value.Equals("Consultant", StringComparison.OrdinalIgnoreCase)))
                {
                    headerRowNumber = rowNumber;
                    headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                    foreach (var item in rowValues)
                    {
                        string header = item.Value.Trim();

                        if (string.IsNullOrWhiteSpace(header))
                            continue;

                        if (!headers.ContainsKey(header))
                            headers.Add(header, item.Key);
                    }

                    break;
                }
            }

            if (headerRowNumber == -1)
                throw new Exception("Could not find Consultant header in Open Invoice Report.");

            int consultantColumn = GetRequiredColumn(headers, "Consultant");
            int serviceEndColumn = GetRequiredColumn(headers, "Service End");
            int documentNumberColumn = GetRequiredColumn(headers, "Document Number");
            int remainingAmountColumn = GetRequiredColumn(headers, "Remaining Amount");
            int clientProjectsColumn = GetRequiredColumn(headers, "Client Projects");

            foreach (Row row in sheetData.Elements<Row>())
            {
                int rowNumber = (int)row.RowIndex!.Value;

                if (rowNumber <= headerRowNumber)
                    continue;

                var rowValues = ReadRow(row, sharedStrings);

                string consultant = GetValue(rowValues, consultantColumn).Trim();
                string serviceEnd = FormatServiceEnd(GetValue(rowValues, serviceEndColumn));
                string documentNumber = GetValue(rowValues, documentNumberColumn).Trim();
                decimal remainingAmount = ParseDecimal(GetValue(rowValues, remainingAmountColumn));
                string clientProjects = GetValue(rowValues, clientProjectsColumn).Trim();

                if (string.IsNullOrWhiteSpace(consultant) || string.IsNullOrWhiteSpace(serviceEnd))
                    continue;

                string key = $"{consultant} {serviceEnd}".Trim();

                if (!matches.ContainsKey(key))
                    matches[key] = new List<OirMatch>();

                matches[key].Add(new OirMatch(documentNumber, remainingAmount, clientProjects));
            }

            return matches;
        }

        public static Dictionary<string, List<OirMatch>> LoadByClientProject(string filePath)
        {
            var matches = new Dictionary<string, List<OirMatch>>(15000, StringComparer.OrdinalIgnoreCase);

            using SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false);

            WorkbookPart workbookPart = document.WorkbookPart
                ?? throw new Exception("Could not read workbook.");

            SharedStringTable? sharedStrings =
                workbookPart.SharedStringTablePart?.SharedStringTable;

            Sheet firstSheet = workbookPart.Workbook.Sheets!
                .Elements<Sheet>()
                .First();

            WorksheetPart worksheetPart =
                (WorksheetPart)workbookPart.GetPartById(firstSheet.Id!);

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()
                ?? throw new Exception("Could not read worksheet data.");

            int headerRowNumber = -1;
            Dictionary<string, int> headers = new(StringComparer.OrdinalIgnoreCase);

            foreach (Row row in sheetData.Elements<Row>())
            {
                int rowNumber = (int)row.RowIndex!.Value;

                if (rowNumber > 25)
                    break;

                var rowValues = ReadRow(row, sharedStrings);

                if (rowValues.Any(x => x.Value.Equals("Consultant", StringComparison.OrdinalIgnoreCase)))
                {
                    headerRowNumber = rowNumber;
                    headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                    foreach (var item in rowValues)
                    {
                        string header = item.Value.Trim();

                        if (!string.IsNullOrWhiteSpace(header) && !headers.ContainsKey(header))
                            headers.Add(header, item.Key);
                    }

                    break;
                }
            }

            if (headerRowNumber == -1)
                throw new Exception("Could not find Consultant header in Open Invoice Report.");

            int clientProjectsColumn = GetRequiredColumn(headers, "Client Projects");
            int documentNumberColumn = GetRequiredColumn(headers, "Document Number");
            int remainingAmountColumn = GetRequiredColumn(headers, "Remaining Amount");

            foreach (Row row in sheetData.Elements<Row>())
            {
                int rowNumber = (int)row.RowIndex!.Value;

                if (rowNumber <= headerRowNumber)
                    continue;

                var rowValues = ReadRow(row, sharedStrings);

                string clientProjects = GetValue(rowValues, clientProjectsColumn).Trim();
                string documentNumber = GetValue(rowValues, documentNumberColumn).Trim();
                decimal remainingAmount = ParseDecimal(GetValue(rowValues, remainingAmountColumn));

                if (string.IsNullOrWhiteSpace(clientProjects))
                    continue;

                if (!matches.ContainsKey(clientProjects))
                    matches[clientProjects] = new List<OirMatch>();

                matches[clientProjects].Add(new OirMatch(documentNumber, remainingAmount, clientProjects));
            }

            return matches;
        }

        private static Dictionary<int, string> ReadRow(Row row, SharedStringTable? sharedStrings)
        {
            var values = new Dictionary<int, string>();

            foreach (Cell cell in row.Elements<Cell>())
            {
                int columnIndex = GetColumnIndex(cell.CellReference?.Value);

                if (columnIndex == -1)
                    continue;

                values[columnIndex] = GetCellValue(cell, sharedStrings);
            }

            return values;
        }

        private static string GetCellValue(Cell cell, SharedStringTable? sharedStrings)
        {
            string rawValue = cell.CellValue?.Text ?? "";

            if (cell.DataType == null)
                return rawValue;

            if (cell.DataType.Value == CellValues.SharedString &&
                int.TryParse(rawValue, out int sharedStringIndex) &&
                sharedStrings != null)
            {
                return sharedStrings.ElementAt(sharedStringIndex).InnerText;
            }

            if (cell.DataType.Value == CellValues.InlineString)
                return cell.InlineString?.InnerText ?? "";

            return rawValue;
        }

        private static int GetColumnIndex(string? cellReference)
        {
            if (string.IsNullOrWhiteSpace(cellReference))
                return -1;

            string letters = Regex.Replace(cellReference.ToUpper(), @"[\d]", "");

            int columnIndex = 0;

            foreach (char letter in letters)
            {
                columnIndex *= 26;
                columnIndex += letter - 'A' + 1;
            }

            return columnIndex;
        }

        private static int GetRequiredColumn(Dictionary<string, int> headers, string headerName)
        {
            if (!headers.TryGetValue(headerName, out int column))
                throw new Exception($"Could not find required column: {headerName}");

            return column;
        }

        private static string GetValue(Dictionary<int, string> rowValues, int column)
        {
            return rowValues.TryGetValue(column, out string? value) ? value : "";
        }

        private static string FormatServiceEnd(string value)
        {
            value = value.Trim();

            if (double.TryParse(value, out double oaDate))
                return DateTime.FromOADate(oaDate).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

            if (DateTime.TryParse(value, out DateTime parsedDate))
                return parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

            return value;
        }

        private static decimal ParseDecimal(string value)
        {
            value = value
                .Replace("$", "")
                .Replace(",", "")
                .Trim();

            if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
                return result;

            return 0;
        }
    }
}