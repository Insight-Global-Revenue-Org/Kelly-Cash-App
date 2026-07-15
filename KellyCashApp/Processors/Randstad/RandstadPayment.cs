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
        // Primary Process Tree for the Randstad payment workflow - Extract PDF text, format from CSV and generating the Excel output file.
        public static bool IsRandstadFormat(string inputPath)
        {
            if (!Path.GetExtension(inputPath)
                .Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            try
            {
                using PdfDocument document = PdfDocument.Open(inputPath);

                string amountPattern =
                    @"(?:-?\s*\$?[\d,]+\.\d{2}|\(\s*\$?[\d,]+\.\d{2}\s*\))";

                string rowPattern =
                    $@"Other\s+(?<invoice>\d+)\s+" +
                    @"(?<date>\d{1,2}/\d{1,2}/\d{2,4})\s+" +
                    @"(?<name>.+?)\s+" +
                    $@"(?<gross>{amountPattern})" +
                    $@"(?:\s+(?<discount>{amountPattern}))?\s+" +
                    $@"(?<paid>{amountPattern})";

                foreach (var page in document.GetPages())
                {
                    string text = Regex.Replace(
                        page.Text,
                        @"\s+",
                        " ");

                    if (Regex.IsMatch(
                        text,
                        rowPattern,
                        RegexOptions.IgnoreCase))
                    {
                        return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        public static string Process(
                string inputPath,
                 Dictionary<string, OirMatch> openInvoiceMatches,
                 Dictionary<string, List<OirMatch>> openInvoiceMatchesByClientProject)
        {
            // Initialize a list to hold the output rows
            var outputRows = new List<RandstadOutputRow>();

            // Get the path to the Nike Tracker board file from settings
            string nikeTrackerPath = Settings.GetNikeTrackerBoardFilePath();

            // Open using PDF Library
            using PdfDocument document = PdfDocument.Open(inputPath);

            // Small loop through each page
            foreach (var page in document.GetPages())
            {
                // Extract all plain text from the PDF! Parsing logic is below
                string text = page.Text;

                // Regex pattern to match amounts in the format of $1,234.56 or ($1,234.56)
                string amountPattern = @"(?:-?\s*\$?[\d,]+\.\d{2}|\(\s*\$?[\d,]+\.\d{2}\s*\))";

                // REGEX Parser
                // A (very) complex regex pattern to match lines in the text that start with "Other", followed by an invoice number, a date, a name, and amounts for gross, discount (optional), and paid. It captures these values into named groups which we'll pull into the output file :)
                var matches = Regex.Matches(
                    text,
                    $@"Other\s+(?<invoice>\d+)\s+(?<date>\d{{1,2}}/\d{{1,2}}/\d{{2,4}})\s+(?<name>.+?)\s+(?<gross>{amountPattern})(?:\s+(?<discount>{amountPattern}))?\s+(?<paid>{amountPattern})(?=\s*Other\s+\d+|\s*Deposit Totals|$)");

                // And for each matched row, extract everything
                foreach (Match match in matches)
                {
                    // Extract the captured groups from the regex match
                    string invoiceNumber = match.Groups["invoice"].Value.Trim();
                    string formattedDate = FormatDate(match.Groups["date"].Value.Trim());
                    string name = FormatName(match.Groups["name"].Value.Trim());
                    decimal paidAmount = GetDecimalValue(match.Groups["paid"].Value);
                    // Create a concatenated string for matching purposes
                    string concat = $"{name} {formattedDate}";
                    string beelineId = GetBeelineId(name);
                    // Initialize variables for invoice, amount due, and client project
                    string invoice = "";
                    decimal amountDue = 0;
                    string clientProject = "";

                    // Attempt to match the name and date with the open invoice matches, allowing for a date tolerance of +/- 2 days
                    if (TryMatchWithDateTolerance(name, formattedDate, openInvoiceMatches, out OirMatch oirMatch))
                    {
                        invoice = oirMatch.DocumentNumber;
                        amountDue = oirMatch.RemainingAmount;
                    }

                    // If the Beeline ID and Nike Tracker path are available, attempt to find the best match in the Nike Tracker
                    if (!string.IsNullOrWhiteSpace(beelineId) &&
                    !string.IsNullOrWhiteSpace(nikeTrackerPath))
                    {
                        NikeTrackerMatch? nikeMatch = NikeTracker.FindBestMatch(
                            nikeTrackerPath,
                            beelineId,
                            formattedDate,
                            paidAmount
                        );

                        // If a match is found and the client project name exists in the open invoice matches by client project, set the client project for the output row
                        if (nikeMatch != null &&
                            openInvoiceMatchesByClientProject.TryGetValue(nikeMatch.ClientProjectName, out List<OirMatch>? projectMatches))
                        {
                            clientProject = nikeMatch.ClientProjectName;
                        }
                    }

                    // Add the extracted and processed data to the output rows list
                    outputRows.Add(new RandstadOutputRow
                    {
                        WeekEndingDate = formattedDate,
                        Name = name,
                        Invoice = invoice,
                        AmountDue = amountDue,
                        AggregateAmountPaid = paidAmount,
                        Notes = "",
                        Concat = concat,
                        InvoiceNumber = invoiceNumber,
                        BeelineId = beelineId,
                        ClientProject = clientProject
                    });
                }
            }
            
            ApplyGroupedSowMatching(outputRows, openInvoiceMatchesByClientProject);
            outputRows = outputRows
                    .OrderBy(x => x.Name)
                    .ThenBy(x => x.WeekEndingDate)
                    .ThenBy(x => x.InvoiceNumber)
                    .ToList();
            // Generate the Excel output file using ClosedXML
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Randstad Payment");

            // Define the headers for the Excel sheet
            string[] headers =
            {
                "Week Ending Date",
                "Name",
                "Invoice",
                "Amount Due",
                "Aggregate Amount Paid",
                "Notes",
                "Concat",
                "Invoice Number",
                "Beeline ID",
                "Client Project"
            };

            // Write the headers to the first row of the worksheet
            for (int col = 1; col <= headers.Length; col++)
                worksheet.Cell(1, col).Value = headers[col - 1];

            // Write the output rows to the worksheet starting from the second row
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
                worksheet.Cell(row, 9).Value = item.BeelineId;
                worksheet.Cell(row, 10).Value = item.ClientProject;
            }
            
            ApplyFormatting(worksheet, outputRows.Count + 1, headers.Length);

            // Save the workbook to a unique output path in the downloads folder
            string downloadsPath = Settings.GetRemittanceSavePath();
            decimal total = outputRows.Sum(x => x.AggregateAmountPaid);
            string formattedTotal = total.ToString("$#,##0.00;($#,##0.00)", CultureInfo.InvariantCulture);
            string processedDate = DateTime.Now.ToString("M.d.yyyy", CultureInfo.InvariantCulture);

            // Generate a unique output path for the Excel file to avoid overwriting existing files
            string outputPath = GetUniqueOutputPath(downloadsPath, $"Randstad {processedDate} - {formattedTotal}.xlsx");

            // Now finally save it!!!
            workbook.SaveAs(outputPath);

            // Logging for the Analytics class
            Analytics.LogRemittanceRun($"Randstad - {formattedTotal}");

            return outputPath;
        }

        // Helper function to apply grouped SOW matching logic to the output rows based on client project matches
        private static void ApplyGroupedSowMatching(
                    List<RandstadOutputRow> outputRows,
                    Dictionary<string, List<OirMatch>> openInvoiceMatchesByClientProject)
        {
            var usedInvoices = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var sowGroups = outputRows
                .Where(x =>
                    !string.IsNullOrWhiteSpace(x.BeelineId) &&
                    !string.IsNullOrWhiteSpace(x.ClientProject) &&
                    string.IsNullOrWhiteSpace(x.Invoice))
                .GroupBy(x => new
                {
                    x.BeelineId,
                    x.ClientProject
                });

            // For each group of SOWs, attempt to find the best match in the open invoice matches by client project
            foreach (var group in sowGroups)
            {
                var rows = group.ToList();

                if (!openInvoiceMatchesByClientProject.TryGetValue(group.Key.ClientProject, out List<OirMatch>? projectMatches))
                    continue;

                TryApplyMatch(rows, projectMatches, usedInvoices, 0.10m);

                foreach (var combo in GetCombinations(rows.Where(x => string.IsNullOrWhiteSpace(x.Invoice)).ToList(), 2))
                    TryApplyMatch(combo, projectMatches, usedInvoices, 0.05m);

                foreach (var combo in GetCombinations(rows.Where(x => string.IsNullOrWhiteSpace(x.Invoice)).ToList(), 3))
                    TryApplyMatch(combo, projectMatches, usedInvoices, 0.05m);

                TryApplyMatch(rows.Where(x => string.IsNullOrWhiteSpace(x.Invoice)).ToList(), projectMatches, usedInvoices, 0.10m);
            }
        }

        // Helper function to attempt to apply a match between a group of rows and the available project matches based on the allowed percent difference
        private static void TryApplyMatch(
            List<RandstadOutputRow> rows,
            List<OirMatch> projectMatches,
            HashSet<string> usedInvoices,
            decimal allowedPercentDifference)
        {
            if (rows.Count == 0)
                return;

            // Calculate the total paid amount for the group of rows
            decimal groupedPaidAmount = rows.Sum(x => x.AggregateAmountPaid);

            // Find the best match in the project matches that hasn't already been used and is closest to the grouped paid amount
            OirMatch? bestOirMatch = projectMatches
                .Where(x => !usedInvoices.Contains(x.DocumentNumber))
                .OrderBy(x => Math.Abs(x.RemainingAmount - groupedPaidAmount))
                .FirstOrDefault();

            // If no suitable match is found or the best match is not within the allowed percent difference, return without applying any changes
            if (bestOirMatch == null)
                return;

            // Check if the grouped paid amount is within the allowed percent difference of the best match's remaining amount
            if (!IsWithinPercent(groupedPaidAmount, bestOirMatch.RemainingAmount, allowedPercentDifference))
                return;

            // If a suitable match is found, apply the match to all rows in the group by setting the invoice number and amount due
            foreach (var row in rows)
            {
                row.Invoice = bestOirMatch.DocumentNumber;
                row.AmountDue = bestOirMatch.RemainingAmount;
            }

            // Mark the matched invoice as used to prevent it from being matched again in future iterations
            usedInvoices.Add(bestOirMatch.DocumentNumber);
        }

        // Helper function to generate all combinations of a specified size from a list of items
        private static List<List<T>> GetCombinations<T>(List<T> items, int size)
        {
            var results = new List<List<T>>();

            void Build(int start, List<T> current)
            {
                if (current.Count == size)
                {
                    results.Add(new List<T>(current));
                    return;
                }

                // Loop through the items starting from the current index and recursively build combinations
                for (int i = start; i < items.Count; i++)
                {
                    current.Add(items[i]);
                    Build(i + 1, current);
                    current.RemoveAt(current.Count - 1);
                }
            }

            // Start the recursive building of combinations
            Build(0, new List<T>());
            return results;
        }

        // Helper function to apply formatting to the Excel worksheet, including font styles, borders, alignment, and conditional formatting based on the data
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

            // Apply conditional formatting based on the data in the worksheet
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

            // Adjust column widths to fit the contents of the cells
            worksheet.Columns().AdjustToContents();

            worksheet.Column(3).Width = 18; // Invoice
            worksheet.Column(6).Width = 30; // Notes

            worksheet.Range(1, 1, lastRow, lastColumn).SetAutoFilter();
        }

        // Helper function to format a date string into "MM/dd/yyyy" format, returning the original value if parsing fails
        private static string FormatDate(string value)
        {
            if (DateTime.TryParse(value, out DateTime date))
                return date.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

            return value;
        }

        // Helper function to format a name string, converting it to "First Last" format and applying title case
        private static string FormatName(string input)
        {
            if (!input.Contains(","))
                return input;

            string[] parts = input.Split(',', 2);

            string last = ToTitle(parts[0].Trim());
            string first = ToTitle(parts[1].Trim());

            return $"{first} {last}".Trim();
        }

        // Helper function to extract the Beeline ID from a name string using regex, returning an empty string if no match is found
        private static string GetBeelineId(string name)
        {
            var match = Regex.Match(name, @"^(?<id>\d+)_\d+\s+Sow$", RegexOptions.IgnoreCase);

            if (match.Success)
                return match.Groups["id"].Value;

            return "";
        }

        // Helper function to attempt to match a name and date with the open invoice matches, allowing for a date tolerance of +/- 2 days. It returns true if a match is found and outputs the matched OirMatch.
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

        // Helper function to convert a string to title case using the current culture's text info
        private static string ToTitle(string value)
        {
            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(value.ToLower());
        }

        // maybe a temporary helper function to check if the paid amount is within 10% of the amount due
        private static bool IsWithinPercent(decimal paidAmount, decimal amountDue, decimal allowedPercentDifference)
        {
            if (amountDue == 0)
                return false;

            decimal difference = Math.Abs(paidAmount - amountDue);
            decimal percentDifference = difference / Math.Abs(amountDue);

            return percentDifference <= allowedPercentDifference;
        }

        // Helper function to parse a string value into a decimal, removing any currency symbols and commas. Returns 0 if parsing fails.
        private static decimal GetDecimalValue(string value)
        {
            string raw = value.Replace("$", "").Replace(",", "").Trim();

            if (decimal.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
                return result;


            return 0;
        }

        // Helper function to generate a unique output path for the Excel file in the specified folder, appending a counter if a file with the same name already exists
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

        // represents a match found in our Randstad vars. initiated to null (rows of output data manipulated and stored in memory for the Excel output file)
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
            public string BeelineId { get; set; } = "";
            public string ClientProject { get; set; } = "";
        }
    }
}