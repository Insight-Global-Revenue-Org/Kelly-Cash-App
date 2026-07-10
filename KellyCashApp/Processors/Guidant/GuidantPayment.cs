using ClosedXML.Excel;
using KellyCashApp.Configuration;
using KellyCashApp.Models;
using KellyCashApp.Services;
using System.Globalization;

namespace KellyCashApp.Processors.Guidant
{
    internal class GuidantPayment
    {
        // Payment Processing Class for Guidant - The excel logic below is triggered from Program.cs through its conditional workflow check. If you need to update any of your excel logic in the future, do it here!
        // Guidant header row identifiers (Customer, Vendor, Con Invoice) - I moved this to a helper method at the bottom of this class (if it needs to change in the future, it can be done in one place)
        public static bool IsGuidantFormat(IXLWorksheet worksheet)
        {
            return FindGuidantHeaderRow(worksheet) != -1;
        }

        // Primary public method that processes all worksheets in the workbook
        public static string Process(
            XLWorkbook workbook,
            IXLWorksheet worksheet,
            string inputPath,
            Dictionary<string, OirMatch> openInvoiceMatches)
        {
            decimal totalAcrossAllSheets = 0;
            // Find all worksheets in the workbook that match the Guidant format
            var guidantSheets = workbook.Worksheets
                .Where(sheet => IsGuidantFormat(sheet))
                .ToList();

            // Process all detected Guidant sheets (as long as they meet the header criteria!)
            foreach (var sheet in guidantSheets)
            {
                decimal sheetTotal = ProcessSingleSheet(sheet, openInvoiceMatches);
                totalAcrossAllSheets += sheetTotal;
            }

            // If no Guidant sheets were found, throw an exception
            string downloadsPath = Settings.GetRemittanceSavePath();
            // If the downloads path is not set, throw an exception
            string formattedTotal = totalAcrossAllSheets.ToString("N2", CultureInfo.InvariantCulture);
            // If the downloads path is not set, throw an exception
            string outputPath = GetUniqueOutputPath(
                downloadsPath,
                $"Guidant - {formattedTotal}.xlsx"
            );

            // Saves the processed workbook to the output path
            workbook.SaveAs(outputPath);

            // Log the remittance run in analytics
            Analytics.LogRemittanceRun($"Guidant - {formattedTotal}");

            return outputPath;
        }

        // Primary Payment Processing Logic for Guidant!
        private static decimal ProcessSingleSheet(
            IXLWorksheet worksheet,
            Dictionary<string, OirMatch> openInvoiceMatches)
        {
            int headerRow = FindGuidantHeaderRow(worksheet);

            if (headerRow == -1)
            {
                throw new Exception($"Could not find Guidant header row on sheet: {worksheet.Name}");
            }

            // Find the required columns based on the header row
            int invoiceColumn = FindColumn(worksheet, headerRow, "Invoice");
            int conInvoiceColumn = FindColumn(worksheet, headerRow, "Con Invoice");

            int invoiceIdColumn = FindColumn(worksheet, headerRow, "Invoice ID");
            int timesheetColumn = FindColumn(worksheet, headerRow, "Timesheet");
            int weekEndingColumn = FindColumn(worksheet, headerRow, "WeekEnding");
            int associateColumn = FindColumn(worksheet, headerRow, "Associate");
            int hoursTypeColumn = FindColumn(worksheet, headerRow, "Hours Type");
            int billHoursColumn = FindColumn(worksheet, headerRow, "Bill Hours");
            int billAmountColumn = FindColumn(worksheet, headerRow, "Bill Amount");

            if ((invoiceColumn == -1 && conInvoiceColumn == -1) || invoiceIdColumn == -1 || timesheetColumn == -1 ||
                weekEndingColumn == -1 || associateColumn == -1 || hoursTypeColumn == -1 ||
                billHoursColumn == -1 || billAmountColumn == -1)
            {
                throw new Exception($"Missing one or more required Guidant columns on sheet: {worksheet.Name}");
            }

            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            // Load name changes from the configuration file (if any)
            var nameChanges = Rename.LoadNameChanges();
            var outputRows = new List<GuidantOutputRow>();

            // Process each row in the worksheet, starting from the row after the header
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string rawName = worksheet.Cell(row, associateColumn).GetString().Trim();

                if (string.IsNullOrWhiteSpace(rawName))
                    continue;

                // Skip rows with empty names
                string formattedName = FormatName(rawName);
                formattedName = Rename.ApplyNameChange(formattedName, nameChanges);

                // Skip rows with empty week ending dates
                string formattedWeekEnding = FormatDate(worksheet.Cell(row, weekEndingColumn));
                string concat = $"{formattedName} {formattedWeekEnding}".Trim();

                string oirInvoice = "";
                decimal amountDue = 0;

                // Determine if the row is an expense or non-expense based on the "Hours Type" column
                string hoursType = worksheet.Cell(row, hoursTypeColumn).GetString().Trim();
                decimal billAmount = GetDecimalValue(worksheet.Cell(row, billAmountColumn));

                // Attempt to match the row with an open invoice based on the hours type and date spread
                bool matched = hoursType.Equals("EXP", StringComparison.OrdinalIgnoreCase)
                    ? TryMatchExpenseWithSevenDaySpread(formattedName, formattedWeekEnding, billAmount, openInvoiceMatches, out OirMatch match)
                    : TryMatchWithOneDaySpread(formattedName, formattedWeekEnding, openInvoiceMatches, out match);

                // If a match is found, retrieve the corresponding invoice number and amount due
                if (matched)
                {
                    oirInvoice = match.DocumentNumber;
                    amountDue = match.RemainingAmount;
                }

                // Add the processed row to the output list
                outputRows.Add(new GuidantOutputRow
                {
                    WeekEndingDate = formattedWeekEnding,
                    Name = formattedName,
                    Invoice = oirInvoice,
                    AmountDue = amountDue,
                    AggregateAmountPaid = billAmount,
                    Notes = "",
                    Concat = concat,
                    GuidantInvoice = invoiceColumn != -1
                        ? worksheet.Cell(row, invoiceColumn).GetString().Trim()
                        : worksheet.Cell(row, conInvoiceColumn).GetString().Trim(),
                    InvoiceId = worksheet.Cell(row, invoiceIdColumn).GetString().Trim(),
                    TimesheetId = worksheet.Cell(row, timesheetColumn).GetString().Trim(),
                    HoursType = hoursType,
                    BillHours = GetDecimalValue(worksheet.Cell(row, billHoursColumn))
                });
            }

            // Calculate the aggregate amount paid for each unique "Concat" value
            var totalsByConcat = outputRows
                .GroupBy(r => r.Concat)
                .ToDictionary(g => g.Key, g => g.Sum(r => r.AggregateAmountPaid));

            // Update each output row with the calculated aggregate amount paid
            foreach (var row in outputRows)
            {
                row.AggregateAmountPaid = totalsByConcat[row.Concat];
            }

            // sort alphabetically :)
            outputRows = outputRows
                .OrderBy(r => r.Name)
                .ThenBy(r => r.WeekEndingDate)
                .ToList();

            worksheet.Clear();

            // Write headers to the first row of the worksheet
            string[] headers =
            {
                "Week Ending Date",
                "Name",
                "Invoice",
                "Amount Due",
                "Aggregate Amount Paid",
                "Notes",
                "Concat",
                "Hours Type",
                "Guidant Invoice",
                "Invoice ID",
                "Timesheet ID",
                "Bill Hours"
            };

            // First, write the headers to the first row of the worksheet
            for (int col = 1; col <= headers.Length; col++)
                worksheet.Cell(1, col).Value = headers[col - 1];

            // Second, write the processed output rows to the worksheet, starting from the second row
            for (int i = 0; i < outputRows.Count; i++)
            {
                int row = i + 2;
                GuidantOutputRow item = outputRows[i];

                // Write each property of the output row to the corresponding cell in the worksheet
                worksheet.Cell(row, 1).Value = item.WeekEndingDate;
                worksheet.Cell(row, 2).Value = item.Name;
                worksheet.Cell(row, 3).Value = item.Invoice;

                // Only write the Amount Due if it's not zero to avoid cluttering the output with unnecessary zeros
                if (item.AmountDue != 0)
                    worksheet.Cell(row, 4).Value = item.AmountDue;

                // Otherwise, leave the cell empty for clarity
                worksheet.Cell(row, 5).Value = item.AggregateAmountPaid;
                worksheet.Cell(row, 6).Value = item.Notes;
                worksheet.Cell(row, 7).Value = item.Concat;
                worksheet.Cell(row, 8).Value = item.HoursType;
                worksheet.Cell(row, 9).Value = item.GuidantInvoice;
                worksheet.Cell(row, 10).Value = item.InvoiceId;
                worksheet.Cell(row, 11).Value = item.TimesheetId;
                worksheet.Cell(row, 12).Value = item.BillHours;
            }

            // Apply standardized formatting to the output worksheet, including fonts, borders, number formats, and column widths (review the ApplyFormatting method for details!!!!!!!!)
            ApplyFormatting(worksheet, outputRows.Count + 1, headers.Length);
            // Return the total aggregate amount paid across all output rows for this sheet
            return outputRows.Sum(r => r.AggregateAmountPaid);
        }

        // Try to match a non-expense row with a ±1 day spread
        private static bool TryMatchWithOneDaySpread(
            string formattedName,
            string formattedWeekEnding,
            Dictionary<string, OirMatch> openInvoiceMatches,
            out OirMatch match)
        {
            match = null!;

            // Check if the formattedWeekEnding can be parsed into a DateTime object
            if (!DateTime.TryParse(formattedWeekEnding, out DateTime baseDate))
                return false;

            // Check for matches within a ±1 day spread
            for (int offset = -1; offset <= 1; offset++)
            {
                DateTime testDate = baseDate.AddDays(offset);
                string testDateString = testDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                string testKey = $"{formattedName} {testDateString}".Trim();

                // Try to find a match in the openInvoiceMatches dictionary using the constructed testKey
                if (openInvoiceMatches.TryGetValue(testKey, out match))
                    return true;
            }

            // If no match is found within the ±1 day spread, return false - the match variable will remain null
            return false;
        }


        // Formats the name from "Last, First" to "First Last" and apply title casing
        private static string FormatName(string input)
        {
            input = input.Trim();

            if (!input.Contains(","))
                return input;

            // Split the input into two parts: last name and first name, using the comma as a delimiter
            string[] parts = input.Split(',', 2);

            // Convert both parts to title case and trim any extra whitespace
            string last = ToTitle(parts[0].Trim());
            string first = ToTitle(parts[1].Trim());

            // Return the formatted name in "First Last" order, ensuring no leading or trailing whitespace
            return $"{first} {last}".Trim();
        }

        // Converts a string to title case using the current culture
        private static string ToTitle(string value)
        {
            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(value.ToLower());
        }

        // Formats the date in the cell to "MM/dd/yyyy" format, ***handles both DateTime and string representations
        private static string FormatDate(IXLCell cell)
        {
            if (cell.TryGetValue(out DateTime dateValue))
                return dateValue.ToString("M/d/yyyy", CultureInfo.InvariantCulture);

            if (cell.TryGetValue(out double serialDate))
            {
                try
                {
                    return DateTime.FromOADate(serialDate).ToString("M/d/yyyy", CultureInfo.InvariantCulture);
                }
                catch
                {
                    // Fall through and try parsing as text
                }
            }

            string rawValue = cell.GetFormattedString().Trim();

            if (DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                return parsedDate.ToString("M/d/yyyy", CultureInfo.InvariantCulture);

            rawValue = cell.GetString().Trim();

            if (DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                return parsedDate.ToString("M/d/yyyy", CultureInfo.InvariantCulture);

            return rawValue;
        }

        // Expense matching allows for a ±7 day spread and a 5% margin of error in the amount due.
        private static bool TryMatchExpenseWithSevenDaySpread(
            string formattedName,
            string formattedWeekEnding,
            decimal billAmount,
            Dictionary<string, OirMatch> openInvoiceMatches,
            out OirMatch match)
        {
            match = null!;

            if (!DateTime.TryParse(formattedWeekEnding, out DateTime baseDate))
                return false;

            // First pass: exact amount match within ±7 days
            for (int offset = -7; offset <= 7; offset++)
            {
                DateTime testDate = baseDate.AddDays(offset);
                string testDateString = testDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                string testKey = $"{formattedName} {testDateString}".Trim();

                if (openInvoiceMatches.TryGetValue(testKey, out OirMatch exactMatch)
                    && Math.Abs(exactMatch.RemainingAmount - billAmount) == 0)
                {
                    match = exactMatch;
                    return true;
                }
            }

            // Second pass: within 5% margin within ±7 days
            for (int offset = -7; offset <= 7; offset++)
            {
                DateTime testDate = baseDate.AddDays(offset);
                string testDateString = testDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                string testKey = $"{formattedName} {testDateString}".Trim();

                if (openInvoiceMatches.TryGetValue(testKey, out OirMatch marginMatch)
                    && IsWithinFivePercent(marginMatch.RemainingAmount, billAmount))
                {
                    match = marginMatch;
                    return true;
                }
            }

            return false;
        }

        // Check if the oirAmount is within 5% of the paymentAmount
        private static bool IsWithinFivePercent(decimal oirAmount, decimal paymentAmount)
        {
            if (paymentAmount == 0)
                return oirAmount == 0;

            decimal difference = Math.Abs(oirAmount - paymentAmount);
            decimal allowedDifference = Math.Abs(paymentAmount) * 0.05m;

            return difference <= allowedDifference;
        }

        // Helper to find the header row for the Guidant format by checking the first 10 rows for specific header names
        private static int FindGuidantHeaderRow(IXLWorksheet worksheet)
        {
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 10;

            for (int row = 1; row <= Math.Min(lastRow, 10); row++)
            {
                string firstHeader = worksheet.Cell(row, 1).GetString().Trim();
                string secondHeader = worksheet.Cell(row, 2).GetString().Trim();
                string thirdHeader = worksheet.Cell(row, 3).GetString().Trim();

                bool isGuidantHeader =
                    firstHeader.Equals("Customer", StringComparison.OrdinalIgnoreCase)
                    && secondHeader.Equals("Vendor", StringComparison.OrdinalIgnoreCase)
                        && new[] { "Con Invoice", "Invoice" }
                        .Contains(thirdHeader, StringComparer.OrdinalIgnoreCase);

                if (isGuidantHeader)
                    return row;
            }

            return -1;
        }

        // Helper to find the column index for a given header name in the specified header row
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

        // Helper to parse a cell's value into a decimal, handling currency formatting and parentheses for negative values
        private static decimal GetDecimalValue(IXLCell cell)
        {
            string rawValue = cell.Value.ToString()
                .Replace("$", "")
                .Replace(",", "")
                .Replace("(", "-")
                .Replace(")", "")
                .Trim();

            if (decimal.TryParse(rawValue, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
                return result;

            return 0;
        }

        // Standardized Output File Formatting: fonts, borders, number formats, and column widths
        private static void ApplyFormatting(IXLWorksheet worksheet, int lastRow, int lastColumn)
        {
            var range = worksheet.Range(1, 1, lastRow, lastColumn);

            // Main body formatting
            range.Style.Font.FontName = "Aptos Narrow";
            range.Style.Font.FontSize = 8;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Header row formatting
            worksheet.Row(1).Style.Font.FontName = "Aptos Narrow";
            worksheet.Row(1).Style.Font.FontSize = 9;
            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Row(1).Style.Fill.BackgroundColor = XLColor.FromHtml("#FCE4D6");

            // Row heights
            for (int row = 2; row <= lastRow; row++)
            {
                worksheet.Row(row).Height = 13;
            }

            worksheet.Row(1).Height = 15;

            // Number formatting
            worksheet.Column(4).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)"; // Amount Due
            worksheet.Column(5).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)"; // Aggregate Amount Paid
            worksheet.Column(12).Style.NumberFormat.Format = "0.00";                 // Bill Hours

            // Make zero/negative aggregate payments red
            for (int row = 2; row <= lastRow; row++)
            {
                decimal aggregateAmountPaid = GetDecimalValue(worksheet.Cell(row, 5));

                if (aggregateAmountPaid <= 0)
                {
                    worksheet.Cell(row, 5).Style.Font.FontColor = XLColor.Red;
                }
            }

            // Center identifier/helper columns
            worksheet.Column(8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;  // Hours Type
            worksheet.Column(9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;  // Guidant Invoice
            worksheet.Column(10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Invoice ID
            worksheet.Column(11).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Timesheet ID
            worksheet.Column(12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Bill Hours

            // Freeze top row
            worksheet.SheetView.FreezeRows(1);

            // Auto-fit first, then manually widen important columns
            worksheet.Columns().AdjustToContents();

            worksheet.Column(3).Width = 16;   // Invoice
            worksheet.Column(4).Width = 14;   // Amount Due
            worksheet.Column(5).Width = 20;   // Aggregate Amount Paid
            worksheet.Column(6).Width = 22;   // Notes
            worksheet.Column(7).Width = 28;   // Concat
            worksheet.Column(8).Width = 14;   // Hours Type
            worksheet.Column(9).Width = 18;   // Guidant Invoice
            worksheet.Column(10).Width = 18;  // Invoice ID
            worksheet.Column(11).Width = 18;  // Timesheet ID
            worksheet.Column(12).Width = 12;  // Bill Hours

            // Filter
            worksheet.Range(1, 1, lastRow, lastColumn).SetAutoFilter();
        }

        // Helper to generate a unique output file path by appending a counter if the file already exists
        private static string GetUniqueOutputPath(string folderPath, string fileName)
        {
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);

            string outputPath = Path.Combine(folderPath, fileName);
            int counter = 1;

            // If the file already exists, just append a counter to the filename until a unique name is found i guess (don't overwrite))
            while (File.Exists(outputPath))
            {
                outputPath = Path.Combine(folderPath, $"{fileNameWithoutExt} ({counter}){extension}");
                counter++;
            }

            return outputPath;
        }

        // OutputRow class - the Guidant output row structure that will be used to store processed data before writing it back to the worksheet (nullable strings are initialized to empty strings to avoid null reference issues)
        private class GuidantOutputRow
        {
            public string WeekEndingDate { get; set; } = "";
            public string Name { get; set; } = "";
            public string Invoice { get; set; } = "";
            public decimal AmountDue { get; set; }
            public decimal AggregateAmountPaid { get; set; }
            public string Notes { get; set; } = "";
            public string Concat { get; set; } = "";
            public string GuidantInvoice { get; set; } = "";
            public string InvoiceId { get; set; } = "";
            public string TimesheetId { get; set; } = "";
            public string HoursType { get; set; } = "";
            public decimal BillHours { get; set; }
        }
    }
}