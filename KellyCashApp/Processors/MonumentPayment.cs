using ClosedXML.Excel;
<<<<<<< HEAD:KellyCashApp/Processors/MonumentPayment.cs
using KellyCashApp.Configuration;
using KellyCashApp.Models;
using KellyCashApp.Services;
using System.Globalization;
using System.Text.RegularExpressions;

namespace KellyCashApp.Processors
=======
using System.Globalization;
using System.Text.RegularExpressions;

namespace KellyCashApp
>>>>>>> fd4192c1fc973185485463c4e27a69db29297d34:KellyCashApp/MonumentPayment.cs
{
    internal class MonumentPayment
    {
        public static bool IsMonumentFormat(IXLWorksheet worksheet)
        {
            return worksheet.Cell(1, 1).GetString().Trim()
                .Equals("Invoice Number", StringComparison.OrdinalIgnoreCase)
                && worksheet.Cell(1, 4).GetString().Trim()
                .Equals("Amount Paid", StringComparison.OrdinalIgnoreCase);
        }

        public static string Process(
            XLWorkbook workbook,
            IXLWorksheet worksheet,
            string inputPath,
            Dictionary<string, OirMatch> openInvoiceMatches)
        {
            decimal totalAcrossAllSheets = 0;

            var monumentSheets = workbook.Worksheets
                .Where(sheet => IsMonumentFormat(sheet))
                .ToList();

            foreach (var sheet in monumentSheets)
            {
                decimal sheetTotal = ProcessSingleSheet(sheet, openInvoiceMatches);
                totalAcrossAllSheets += sheetTotal;
            }

            string downloadsPath = Settings.GetRemittanceSavePath();

            string processedDate = DateTime.Now.ToString("M.d.yyyy", CultureInfo.InvariantCulture);
            string formattedTotal = totalAcrossAllSheets.ToString("$#,##0.00;($#,##0.00)", CultureInfo.InvariantCulture);

            string outputPath = GetUniqueOutputPath(
                downloadsPath,
                $"{processedDate} - {formattedTotal}.xlsx"
            );

            workbook.SaveAs(outputPath);

            Analytics.LogRemittanceRun($"Monument - {formattedTotal}");

            return outputPath;
        }

        private static ParsedDetails ParseDetails(string rawDetails)
        {
            rawDetails = rawDetails.Trim();

            if (string.IsNullOrWhiteSpace(rawDetails))
                return new ParsedDetails();

            string credit = "";
            string remainingDetails = rawDetails;

            if (rawDetails.StartsWith("Credit ", StringComparison.OrdinalIgnoreCase))
            {
                string[] parts = rawDetails.Split(':', 2);

                if (parts.Length == 2)
                {
                    credit = parts[0].Trim();
                    remainingDetails = parts[1].Trim();
                }
            }

            Match dateMatch = Regex.Match(
                remainingDetails,
                @"(?<date>\d{1,2}/\d{1,2}/\d{2,4})$"
            );

            string weekEndingDate = "";

            if (dateMatch.Success)
            {
                weekEndingDate = FormatDate(dateMatch.Groups["date"].Value);
                remainingDetails = remainingDetails.Substring(0, dateMatch.Index).Trim();
            }

            string name = FormatName(remainingDetails);

            bool isContractor =
                remainingDetails.Contains(",")
                && !string.IsNullOrWhiteSpace(weekEndingDate);

            return new ParsedDetails
            {
                Name = name,
                WeekEndingDate = weekEndingDate,
                Credit = credit,
                IsContractor = isContractor
            };
        }

        private static string GetUniqueWorksheetName(XLWorkbook workbook, string baseName)
        {
            string name = baseName;
            int counter = 1;

            while (workbook.Worksheets.Any(ws =>
                ws.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                string suffix = $" ({counter})";
                int maxBaseLength = 31 - suffix.Length;

                string trimmedBase = baseName.Length > maxBaseLength
                    ? baseName.Substring(0, maxBaseLength)
                    : baseName;

                name = trimmedBase + suffix;
                counter++;
            }

            return name;
        }

        private static string MakeSafeWorksheetName(string sheetName)
        {
            char[] invalidChars = { ':', '\\', '/', '?', '*', '[', ']' };

            foreach (char invalidChar in invalidChars)
            {
                sheetName = sheetName.Replace(invalidChar, '-');
            }

            if (sheetName.Length > 31)
                sheetName = sheetName.Substring(0, 31);

            return sheetName;
        }

        private static string FormatName(string input)
        {
            input = input.Trim();

            if (!input.Contains(","))
                return input;

            string[] parts = input.Split(',', 2);

            string last = ToTitle(parts[0].Trim());
            string first = ToTitle(parts[1].Trim());

            return $"{first} {last}".Trim();
        }

        private static string ToTitle(string value)
        {
            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(value.ToLower());
        }

        private static string FormatDate(string value)
        {
            if (DateTime.TryParse(value, out DateTime parsed))
                return parsed.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

            return value;
        }

        private static decimal ProcessSingleSheet(
            IXLWorksheet worksheet,
            Dictionary<string, OirMatch> openInvoiceMatches)
        {
            int headerRow = 1;

            int monumentInvoiceColumn = FindColumn(worksheet, headerRow, "Invoice Number");
            int detailAmountColumn = FindColumn(worksheet, headerRow, "Detail Amount");
            int detailsColumn = FindColumn(worksheet, headerRow, "Details");

            if (monumentInvoiceColumn == -1 || detailAmountColumn == -1 || detailsColumn == -1)
                throw new Exception($"Missing required Monument columns on sheet: {worksheet.Name}");

            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            var outputRows = new List<MonumentOutputRow>();

            var nameChanges = Rename.LoadNameChanges();

            string currentMonumentInvoice = "";

            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string invoiceFromCurrentRow = worksheet.Cell(row, monumentInvoiceColumn).GetString().Trim();

                if (!string.IsNullOrWhiteSpace(invoiceFromCurrentRow))
                    currentMonumentInvoice = invoiceFromCurrentRow;

                string rawDetails = worksheet.Cell(row, detailsColumn).GetString().Trim();
                decimal amountPaid = GetDecimalValue(worksheet.Cell(row, detailAmountColumn));

                if (
                    rawDetails.Contains("Check Total", StringComparison.OrdinalIgnoreCase)
                    || invoiceFromCurrentRow.Contains("Check Total", StringComparison.OrdinalIgnoreCase)
)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(currentMonumentInvoice)
                    && string.IsNullOrWhiteSpace(rawDetails)
                    && amountPaid == 0)
                    continue;

                ParsedDetails parsed = ParseDetails(rawDetails);

                parsed.Name = Rename.ApplyNameChange(parsed.Name, nameChanges);

                string concat = "";

                if (!string.IsNullOrWhiteSpace(parsed.Name) && !string.IsNullOrWhiteSpace(parsed.WeekEndingDate))
                    concat = $"{parsed.Name} {parsed.WeekEndingDate}".Trim();

                string invoice = "";
                decimal amountDue = 0;

                if (!string.IsNullOrWhiteSpace(concat)
                    && openInvoiceMatches.TryGetValue(concat, out OirMatch match))
                {
                    invoice = match.DocumentNumber;
                    amountDue = match.RemainingAmount;
                }

                outputRows.Add(new MonumentOutputRow
                {
                    WeekEndingDate = parsed.WeekEndingDate,
                    Name = parsed.Name,
                    Invoice = invoice,
                    AmountDue = amountDue,
                    AggregateAmountPaid = amountPaid,
                    Credit = parsed.Credit,
                    Notes = "",
                    Concat = concat,
                    MonumentInvoice = currentMonumentInvoice,
                    IsContractor = parsed.IsContractor
                });
            }

            var seenContractorKeys = new HashSet<string>();
            var aggregatedOutputRows = new List<MonumentOutputRow>();

            foreach (var row in outputRows)
            {
                if (!row.IsContractor)
                {
                    aggregatedOutputRows.Add(row);
                    continue;
                }

                string key = $"{row.Invoice}|{row.Concat}|{row.Name}|{row.WeekEndingDate}";

                if (seenContractorKeys.Contains(key))
                    continue;

                var matchingRows = outputRows
                    .Where(other =>
                        other.IsContractor
                        && other.Invoice == row.Invoice
                        && other.Concat == row.Concat
                        && other.Name == row.Name
                        && other.WeekEndingDate == row.WeekEndingDate)
                    .ToList();

                row.AggregateAmountPaid = matchingRows.Sum(other => other.AggregateAmountPaid);

                row.Credit = string.Join(", ",
                    matchingRows.Select(other => other.Credit)
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Distinct());

                row.MonumentInvoice = string.Join(", ",
                    matchingRows.Select(other => other.MonumentInvoice)
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Distinct());

                aggregatedOutputRows.Add(row);
                seenContractorKeys.Add(key);
            }

            outputRows = aggregatedOutputRows;

            worksheet.Clear();

            string[] headers =
            {
        "Week Ending Date",
        "Name",
        "Invoice",
        "Amount Due",
        "Aggregate Amount Paid",
        "Credit",
        "Notes",
        "Concat",
        "Monument Invoice"
    };

            for (int col = 1; col <= headers.Length; col++)
                worksheet.Cell(1, col).Value = headers[col - 1];

            for (int i = 0; i < outputRows.Count; i++)
            {
                int row = i + 2;
                MonumentOutputRow item = outputRows[i];

                worksheet.Cell(row, 1).Value = item.WeekEndingDate;
                worksheet.Cell(row, 2).Value = item.Name;
                worksheet.Cell(row, 3).Value = item.Invoice;

                if (item.AmountDue != 0)
                    worksheet.Cell(row, 4).Value = item.AmountDue;

                worksheet.Cell(row, 5).Value = item.AggregateAmountPaid;
                worksheet.Cell(row, 6).Value = item.Credit;
                worksheet.Cell(row, 7).Value = item.Notes;
                worksheet.Cell(row, 8).Value = item.Concat;
                worksheet.Cell(row, 9).Value = item.MonumentInvoice;

                if (!item.IsContractor)
                {
                    worksheet.Range(row, 1, row, 9).Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#F2F2F2"); // White Background 1, Darker 5%
                }
            }

            ApplyCleanFormatting(worksheet, outputRows.Count + 1, headers.Length);

            worksheet.Columns().AdjustToContents();

            decimal sheetTotal = outputRows.Sum(r => r.AggregateAmountPaid);

            string formattedSheetTotal = sheetTotal.ToString("$#,##0.00;-$#,##0.00", CultureInfo.InvariantCulture);

            worksheet.Name = GetUniqueWorksheetName(
            worksheet.Workbook,
            MakeSafeWorksheetName($"{formattedSheetTotal} Check Total")
                );

            return sheetTotal;
        }

        private static void ApplyCleanFormatting(IXLWorksheet worksheet, int lastRow, int lastColumn)
        {
            var range = worksheet.Range(1, 1, lastRow, lastColumn);

            range.Style.Font.FontName = "Arial";
            range.Style.Font.FontSize = 8;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            worksheet.Row(1).Style.Font.FontName = "Arial";
            worksheet.Row(1).Style.Font.FontSize = 9;
            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Row(1).Style.Fill.BackgroundColor = XLColor.FromHtml("#FCE4D6");

            for (int row = 2; row <= lastRow; row++)
            {
                worksheet.Row(row).Height = 13;
            }

            worksheet.Row(1).Height = 15;

            worksheet.Column(4).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
            worksheet.Column(5).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";

            for (int row = 2; row <= lastRow; row++)
            {
                decimal aggregateAmountPaid = GetDecimalValue(worksheet.Cell(row, 5));

                if (aggregateAmountPaid <= 0)
                    worksheet.Cell(row, 5).Style.Font.FontColor = XLColor.Red;
            }

            worksheet.Range(1, 1, lastRow, lastColumn).SetAutoFilter();
        }

        private static int FindColumn(IXLWorksheet worksheet, int headerRow, string headerName)
        {
            int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 50;

            for (int col = 1; col <= lastColumn; col++)
            {
                string headerText = worksheet.Cell(headerRow, col).GetString().Trim();

                if (headerText.Equals(headerName, StringComparison.OrdinalIgnoreCase))
                    return col;
            }

            return -1;
        }

        private static decimal GetDecimalValue(IXLCell cell)
        {
            string raw = cell.Value.ToString()
                .Replace("$", "")
                .Replace(",", "")
                .Replace("(", "-")
                .Replace(")", "")
                .Trim();

            if (decimal.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
                return result;

            return 0;
        }

        private static string GetUniqueOutputPath(string folderPath, string fileName)
        {
            string name = Path.GetFileNameWithoutExtension(fileName);
            string ext = Path.GetExtension(fileName);

            string path = Path.Combine(folderPath, fileName);
            int counter = 1;

            while (File.Exists(path))
            {
                path = Path.Combine(folderPath, $"{name} ({counter}){ext}");
                counter++;
            }

            return path;
        }

        private class ParsedDetails
        {
            public string Name { get; set; } = "";
            public string WeekEndingDate { get; set; } = "";
            public string Credit { get; set; } = "";
            public bool IsContractor { get; set; }
        }

        private class MonumentOutputRow
        {
            public string WeekEndingDate { get; set; } = "";
            public string Name { get; set; } = "";
            public string Invoice { get; set; } = "";
            public decimal AmountDue { get; set; }
            public decimal AggregateAmountPaid { get; set; }
            public string Credit { get; set; } = "";
            public string Notes { get; set; } = "";
            public string Concat { get; set; } = "";
            public string MonumentInvoice { get; set; } = "";
            public bool IsContractor { get; set; }
        }
    }
}