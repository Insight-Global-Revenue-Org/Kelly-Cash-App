using ClosedXML.Excel;
using KellyCashApp.Configuration;
using KellyCashApp.Models;
using KellyCashApp.Services;
using System.Globalization;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace KellyCashApp.Processors.Randstad
{
    internal class RandstadPayment
    {
        public static bool IsRandstadFormat(string inputPath)
        {
            return Path.GetExtension(inputPath)
                .Equals(".pdf", StringComparison.OrdinalIgnoreCase);
        }

        public static string Process(
            string inputPath,
            Dictionary<string, OirMatch> openInvoiceMatches)
        {
            var outputRows = new List<RandstadOutputRow>();

            // Open using PDF Library
            using PdfDocument document = PdfDocument.Open(inputPath);

            // Small loop through each page
            foreach (var page in document.GetPages())
            {
                // Extract all plain text from the PDF! Parsing logic is below
                string text = page.Text;

                string amountPattern = @"(?:-?\s*\$?[\d,]+\.\d{2}|\(\s*\$?[\d,]+\.\d{2}\s*\))";

                // REGEX Parser
                var matches = Regex.Matches(
                    text,
                    $@"Other\s+(?<invoice>\d+)\s+(?<date>\d{{1,2}}/\d{{1,2}}/\d{{2,4}})\s+(?<name>.+?)\s+(?<gross>{amountPattern})(?:\s+(?<discount>{amountPattern}))?\s+(?<paid>{amountPattern})(?=\s*Other\s+\d+|\s*Deposit Totals|$)");

                // And for each matched row, extract everything
                foreach (Match match in matches)
                {
                    string invoiceNumber = match.Groups["invoice"].Value.Trim();
                    string formattedDate = FormatDate(match.Groups["date"].Value.Trim());
                    string name = FormatName(match.Groups["name"].Value.Trim());
                    decimal paidAmount = GetDecimalValue(match.Groups["paid"].Value);

                    string concat = $"{name} {formattedDate}";

                    string invoice = "";
                    decimal amountDue = 0;

                    if (TryMatchWithDateTolerance(name, formattedDate, openInvoiceMatches, out OirMatch oirMatch))
                    {
                        invoice = oirMatch.DocumentNumber;
                        amountDue = oirMatch.RemainingAmount;
                    }

                    outputRows.Add(new RandstadOutputRow
                    {
                        WeekEndingDate = formattedDate,
                        Name = name,
                        Invoice = invoice,
                        AmountDue = amountDue,
                        AggregateAmountPaid = paidAmount,
                        Notes = "",
                        Concat = concat,
                        InvoiceNumber = invoiceNumber
                    });
                }
            }

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Randstad Payment");

            string[] headers =
            {
                "Week Ending Date",
                "Name",
                "Invoice",
                "Amount Due",
                "Aggregate Amount Paid",
                "Notes",
                "Concat",
                "Invoice Number"
            };

            for (int col = 1; col <= headers.Length; col++)
                worksheet.Cell(1, col).Value = headers[col - 1];

            for (int i = 0; i < outputRows.Count; i++)
            {
                int row = i + 2;
                var item = outputRows[i];

                worksheet.Cell(row, 1).Value = item.WeekEndingDate;
                worksheet.Cell(row, 2).Value = item.Name;
                worksheet.Cell(row, 3).Value = item.Invoice;
                worksheet.Cell(row, 4).Value = item.AmountDue;
                worksheet.Cell(row, 5).Value = item.AggregateAmountPaid;
                worksheet.Cell(row, 6).Value = item.Notes;
                worksheet.Cell(row, 7).Value = item.Concat;
                worksheet.Cell(row, 8).Value = item.InvoiceNumber;
            }

            ApplyFormatting(worksheet, outputRows.Count + 1, headers.Length);

            string downloadsPath = Settings.GetRemittanceSavePath();
            decimal total = outputRows.Sum(x => x.AggregateAmountPaid);
            string formattedTotal = total.ToString("$#,##0.00;($#,##0.00)", CultureInfo.InvariantCulture);
            string processedDate = DateTime.Now.ToString("M.d.yyyy", CultureInfo.InvariantCulture);

            string outputPath = GetUniqueOutputPath(downloadsPath, $"Randstad {processedDate} - {formattedTotal}.xlsx");

            workbook.SaveAs(outputPath);

            Analytics.LogRemittanceRun($"Randstad - {formattedTotal}");

            return outputPath;
        }

        private static void ApplyFormatting(IXLWorksheet worksheet, int lastRow, int lastColumn)
        {
            var range = worksheet.Range(1, 1, lastRow, lastColumn);

            range.Style.Font.FontName = "Aptos Narrow";
            range.Style.Font.FontSize = 9;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Row(1).Style.Fill.BackgroundColor = XLColor.FromHtml("#FCE4D6");

            worksheet.Column(4).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
            worksheet.Column(5).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";

            for (int row = 2; row <= lastRow; row++)
            {
                string invoice = worksheet.Cell(row, 3).GetString().Trim();
                decimal amountDue = GetDecimalValue(worksheet.Cell(row, 4).Value.ToString());
                decimal aggregateAmountPaid = GetDecimalValue(worksheet.Cell(row, 5).Value.ToString());

                if (string.IsNullOrWhiteSpace(invoice))
                {
                    worksheet.Range(row, 1, row, lastColumn)
                        .Style.Fill.BackgroundColor = XLColor.FromHtml("#F2F2F2");
                }

                if (amountDue <= 0)
                {
                    worksheet.Cell(row, 4).Style.Font.FontColor = XLColor.Red;
                }

                if (aggregateAmountPaid <= 0)
                {
                    worksheet.Cell(row, 5).Style.Font.FontColor = XLColor.Red;
                }
            }

            worksheet.Columns().AdjustToContents();

            worksheet.Column(3).Width = 18; // Invoice
            worksheet.Column(6).Width = 30; // Notes

            worksheet.Range(1, 1, lastRow, lastColumn).SetAutoFilter();
        }

        private static string FormatDate(string value)
        {
            if (DateTime.TryParse(value, out DateTime date))
                return date.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

            return value;
        }

        private static string FormatName(string input)
        {
            if (!input.Contains(","))
                return input;

            string[] parts = input.Split(',', 2);

            string last = ToTitle(parts[0].Trim());
            string first = ToTitle(parts[1].Trim());

            return $"{first} {last}".Trim();
        }

        private static bool TryMatchWithDateTolerance(
            string name,
            string formattedDate,
            Dictionary<string, OirMatch> openInvoiceMatches,
            out OirMatch match)
        {
            match = null;

            if (!DateTime.TryParse(formattedDate, out DateTime baseDate))
                return false;

            // Primary Loop for non-exact OIR index matches
            string exactKey = $"{name} {baseDate:MM/dd/yyyy}";
            if (openInvoiceMatches.TryGetValue(exactKey, out match))
                return true;

           
            string minusOne = $"{name} {baseDate.AddDays(-1):MM/dd/yyyy}";
            if (openInvoiceMatches.TryGetValue(minusOne, out match))
                return true;

           
            string plusOne = $"{name} {baseDate.AddDays(1):MM/dd/yyyy}";
            if (openInvoiceMatches.TryGetValue(plusOne, out match))
                return true;

            string plusTwo = $"{name} {baseDate.AddDays(2):MM/dd/yyyy}";
            if (openInvoiceMatches.TryGetValue(plusTwo, out match))
                return true;

            string minusTwo = $"{name} {baseDate.AddDays(-2):MM/dd/yyyy}";
            if (openInvoiceMatches.TryGetValue(minusTwo, out match))
                return true;


            return false;
        }

        private static string ToTitle(string value)
        {
            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(value.ToLower());
        }

        private static decimal GetDecimalValue(string value)
        {
            string raw = value.Replace("$", "").Replace(",", "").Trim();

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

        private class RandstadOutputRow
        {
            public string WeekEndingDate { get; set; } = "";
            public string Name { get; set; } = "";
            public string Invoice { get; set; } = "";
            public decimal AmountDue { get; set; }
            public decimal AggregateAmountPaid { get; set; }
            public string Notes { get; set; } = "";
            public string Concat { get; set; } = "";
            public string InvoiceNumber { get; set; } = "";
        }
    }
}