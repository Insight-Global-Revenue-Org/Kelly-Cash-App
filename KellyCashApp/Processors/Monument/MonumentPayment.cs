using ClosedXML.Excel;
using KellyCashApp.Configuration;
using KellyCashApp.Models;
using KellyCashApp.Services;
using System.Globalization;
using System.Text.RegularExpressions;

namespace KellyCashApp.Processors.Monument
{
    internal class MonumentPayment
    {
        public static bool IsMonumentFormat(IXLWorksheet worksheet)
        {
            // These two identifier headers will trigger the conditional check in Program.cs (on other processor classes this may be delegated near the bottom). Update if the Monument format changes in the future.
            // Check if the first row of the worksheet contains the expected headers for Monument format. Specifically, we check if the first cell (A1) contains "Invoice Number" and the fourth cell (D1) contains "Amount Paid". The comparison is case-insensitive and trims any leading or trailing whitespace from the cell values. Update if needed.
            return worksheet.Cell(1, 1).GetString().Trim()
                .Equals("Invoice Number", StringComparison.OrdinalIgnoreCase)
                && worksheet.Cell(1, 4).GetString().Trim()
                .Equals("Amount Paid", StringComparison.OrdinalIgnoreCase);
        }

        // This method processes the Monument payment data in the provided workbook and worksheet, matching it with open invoices and saving the results to a new file.
        public static string Process(
            XLWorkbook workbook,
            IXLWorksheet worksheet,
            string inputPath,
            Dictionary<string, OirMatch> openInvoiceMatches)
        {
            decimal totalAcrossAllSheets = 0;

            // Process the current worksheet if it is in Monument format
            var monumentSheets = workbook.Worksheets
                .Where(sheet => IsMonumentFormat(sheet))
                .ToList();

            // If no Monument format sheets are found, throw an exception
            foreach (var sheet in monumentSheets)
            {
                decimal sheetTotal = ProcessSingleSheet(sheet, openInvoiceMatches);
                totalAcrossAllSheets += sheetTotal;
            }

            // Save the processed workbook to a unique output path in the downloads directory
            string downloadsPath = Settings.GetRemittanceSavePath();

            string processedDate = DateTime.Now.ToString("M.d.yyyy", CultureInfo.InvariantCulture);
            string formattedTotal = totalAcrossAllSheets.ToString("$#,##0.00;($#,##0.00)", CultureInfo.InvariantCulture);

            // Generate a unique output path for the processed file
            string outputPath = GetUniqueOutputPath(
                downloadsPath,
                $"{processedDate} - {formattedTotal}.xlsx"
            );

            // And save the workbook
            workbook.SaveAs(outputPath);

            // Logging for the Analytics class!
            Analytics.LogRemittanceRun($"Monument - {formattedTotal}");

            return outputPath;
        }

        // This method parses the raw details string from the Monument payment data and extracts relevant information such as name, week ending date, credit, and whether the entry is for a contractor.
        private static ParsedDetails ParseDetails(string rawDetails)
        {
            rawDetails = rawDetails.Trim();

            if (string.IsNullOrWhiteSpace(rawDetails))
                return new ParsedDetails();

            // Check for "Credit" prefix and split the string accordingly
            string credit = "";
            string remainingDetails = rawDetails;

            // If the raw details start with "Credit", we split the string into credit and remaining details
            if (rawDetails.StartsWith("Credit ", StringComparison.OrdinalIgnoreCase))
            {
                string[] parts = rawDetails.Split(':', 2);

                if (parts.Length == 2)
                {
                    credit = parts[0].Trim();
                    remainingDetails = parts[1].Trim();
                }
            }

            // Use regex to find a date at the end of the remaining details string
            Match dateMatch = Regex.Match(
                remainingDetails,
                @"(?<date>\d{1,2}/\d{1,2}/\d{2,4})$"
            );

            // If a date is found, we format it and remove it from the remaining details string
            string weekEndingDate = "";

            if (dateMatch.Success)
            {
                weekEndingDate = FormatDate(dateMatch.Groups["date"].Value);
                remainingDetails = remainingDetails.Substring(0, dateMatch.Index).Trim();
            }

            // Format the name from the remaining details string
            string name = FormatName(remainingDetails);

            // Determine if the entry is for a contractor based on the presence of a comma in the remaining details and a non-empty week ending date
            bool isContractor =
                remainingDetails.Contains(",")
                && !string.IsNullOrWhiteSpace(weekEndingDate);

            // Return a ParsedDetails object containing the extracted information
            return new ParsedDetails
            {
                Name = name,
                WeekEndingDate = weekEndingDate,
                Credit = credit,
                IsContractor = isContractor
            };
        }

        // This method generates a unique worksheet name based on the provided base name, ensuring that it does not conflict with existing worksheet names in the workbook.
        // It appends a numeric suffix if necessary to create a unique name.
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

        // This method sanitizes the provided worksheet name by replacing invalid characters with a hyphen and truncating it to a maximum length of 31 characters, which is the limit for Excel worksheet names.
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

        // This method formats a name string by trimming whitespace, checking for a comma to separate last and first names, and converting the names to title case. If no comma is present, it returns the input as is.
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

        // This method converts a string to title case using the current culture's text info. It first converts the string to lowercase and then applies title casing.
        private static string ToTitle(string value)
        {
            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(value.ToLower());
        }

        // This method formats a date string by attempting to parse it into a DateTime object and then returning it in the "MM/dd/yyyy" format. If parsing fails, it returns the original input string.
        private static string FormatDate(string value)
        {
            if (DateTime.TryParse(value, out DateTime parsed))
                return parsed.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

            return value;
        }

        // This method processes a single worksheet in the Monument payment data, extracting relevant information, aggregating amounts for contractors, and formatting the output. It returns the total amount paid for the sheet.
        private static decimal ProcessSingleSheet(
            IXLWorksheet worksheet,
            Dictionary<string, OirMatch> openInvoiceMatches)
        {
            // The header row is assumed to be the first row in the worksheet
            int headerRow = 1;

            // Find the required columns in the worksheet based on the header row
            int monumentInvoiceColumn = FindColumn(worksheet, headerRow, "Invoice Number");
            int detailAmountColumn = FindColumn(worksheet, headerRow, "Detail Amount");
            int detailsColumn = FindColumn(worksheet, headerRow, "Details");

            // If any of the required columns are missing, throw an exception indicating which columns are missing
            if (monumentInvoiceColumn == -1 || detailAmountColumn == -1 || detailsColumn == -1)
                throw new Exception($"Missing required Monument columns on sheet: {worksheet.Name}");

            // Determine the last row used in the worksheet, defaulting to the header row if no rows are used
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            // Create a list to hold the output rows that will be generated from processing the worksheet
            var outputRows = new List<MonumentOutputRow>();

            // Load any name changes from the Rename service to apply to the parsed names
            var nameChanges = Rename.LoadNameChanges();

            // Initialize a variable to keep track of the current Monument invoice number as we iterate through the rows
            string currentMonumentInvoice = "";

            // Iterate through the rows of the worksheet, starting from the row after the header row
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                // Get the Monument invoice number from the current row and trim any whitespace
                string invoiceFromCurrentRow = worksheet.Cell(row, monumentInvoiceColumn).GetString().Trim();

                // If the current row has a non-empty Monument invoice number, update the currentMonumentInvoice variable
                if (!string.IsNullOrWhiteSpace(invoiceFromCurrentRow))
                    currentMonumentInvoice = invoiceFromCurrentRow;
                // Get the raw details and amount paid from the current row
                string rawDetails = worksheet.Cell(row, detailsColumn).GetString().Trim();
                // Get the amount paid from the current row, converting it to a decimal value
                decimal amountPaid = GetDecimalValue(worksheet.Cell(row, detailAmountColumn));

                // Skip rows that contain "Check Total" in either the raw details or the invoice from the current row, as these are not relevant for processing
                if (
                    rawDetails.Contains("Check Total", StringComparison.OrdinalIgnoreCase)
                    || invoiceFromCurrentRow.Contains("Check Total", StringComparison.OrdinalIgnoreCase)
)
                {
                    continue;
                }

                // If the current Monument invoice, raw details, and amount paid are all empty or zero, skip this row as it does not contain relevant data
                if (string.IsNullOrWhiteSpace(currentMonumentInvoice)
                    && string.IsNullOrWhiteSpace(rawDetails)
                    && amountPaid == 0)
                    continue;

                // Parse the raw details to extract the name, week ending date, credit, and contractor status
                ParsedDetails parsed = ParseDetails(rawDetails);

                // Apply any name changes to the parsed name using the Rename service
                parsed.Name = Rename.ApplyNameChange(parsed.Name, nameChanges);

                // Create a concatenated string of the parsed name and week ending date for matching with open invoices
                string concat = "";

                // If both the parsed name and week ending date are not empty, concatenate them with a space in between and trim any whitespace
                if (!string.IsNullOrWhiteSpace(parsed.Name) && !string.IsNullOrWhiteSpace(parsed.WeekEndingDate))
                    concat = $"{parsed.Name} {parsed.WeekEndingDate}".Trim();

                // Look up the open invoice match using the concatenated string and retrieve the corresponding invoice number and amount due if a match is found
                string invoice = "";
                decimal amountDue = 0;

                // If the concatenated string is not empty and a match is found in the openInvoiceMatches dictionary, retrieve the invoice number and remaining amount from the match
                if (!string.IsNullOrWhiteSpace(concat)
                    && openInvoiceMatches.TryGetValue(concat, out OirMatch match))
                {
                    invoice = match.DocumentNumber;
                    amountDue = match.RemainingAmount;
                }

                // Create a new MonumentOutputRow object with the extracted and calculated values, and add it to the outputRows list
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

            // After processing all rows, we will aggregate the output rows for contractors to ensure that we have a single entry for each unique combination of invoice, concatenated name and week ending date.
            // This is done to avoid duplicate entries for contractors who may have multiple payments in the same week.
            var seenContractorKeys = new HashSet<string>();
            // Create a new list to hold the aggregated output rows, which will include both non-contractor rows and aggregated contractor rows
            var aggregatedOutputRows = new List<MonumentOutputRow>();

            // Iterate through the output rows and aggregate the contractor rows based on unique combinations of invoice, concatenated name, and week ending date
            foreach (var row in outputRows)
            {
                if (!row.IsContractor)
                {
                    aggregatedOutputRows.Add(row);
                    continue;
                }

                // Create a unique key for the contractor row based on the invoice, concatenated name, and week ending date. This key will be used to identify duplicate contractor entries.
                string key = $"{row.Invoice}|{row.Concat}|{row.Name}|{row.WeekEndingDate}";

                // If we have already seen this contractor key, we skip adding it to the aggregated output rows to avoid duplicates.
                if (seenContractorKeys.Contains(key))
                    continue;

                // Find all matching contractor rows in the outputRows list that have the same invoice, concatenated name, and week ending date. We will aggregate their amounts paid and credits.
                var matchingRows = outputRows
                    .Where(other =>
                        other.IsContractor
                        && other.Invoice == row.Invoice
                        && other.Concat == row.Concat
                        && other.Name == row.Name
                        && other.WeekEndingDate == row.WeekEndingDate)
                    .ToList();

                // Aggregate the amounts paid for all matching contractor rows and update the current row's AggregateAmountPaid property with the total.
                row.AggregateAmountPaid = matchingRows.Sum(other => other.AggregateAmountPaid);

                // Aggregate the credits for all matching contractor rows, ensuring that we only include non-empty and distinct credit values. We join them into a single string separated by commas.
                row.Credit = string.Join(", ",
                    matchingRows.Select(other => other.Credit)
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Distinct());
                // Aggregate the Monument invoice numbers for all matching contractor rows, ensuring that we only include non-empty and distinct invoice values. We join them into a single string separated by commas.
                row.MonumentInvoice = string.Join(", ",
                    matchingRows.Select(other => other.MonumentInvoice)
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Distinct());
                // Add the aggregated contractor row to the aggregatedOutputRows list and mark the contractor key as seen to avoid duplicates in future iterations.
                aggregatedOutputRows.Add(row);
                seenContractorKeys.Add(key);
            }

            // After aggregating the contractor rows, we replace the original outputRows list with the aggregatedOutputRows list, which now contains both non-contractor rows and aggregated contractor rows.
            outputRows = aggregatedOutputRows;

            // self-explanatory
            worksheet.Clear();

            // Write the headers to the first row of the worksheet, starting from column 1. The headers are defined in the headers array and will be used to label the columns in the output worksheet.
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

            // Write the headers to the first row of the worksheet, starting from column 1. The headers are defined in the headers array and will be used to label the columns in the output worksheet.
            for (int col = 1; col <= headers.Length; col++)
                worksheet.Cell(1, col).Value = headers[col - 1];

            // Write the output rows to the worksheet, starting from row 2. Each output row contains the relevant data extracted and aggregated from the Monument payment data.
            for (int i = 0; i < outputRows.Count; i++)
            {
                int row = i + 2;
                MonumentOutputRow item = outputRows[i];

                worksheet.Cell(row, 1).Value = item.WeekEndingDate;
                worksheet.Cell(row, 2).Value = item.Name;
                worksheet.Cell(row, 3).Value = item.Invoice;

                // If the amount due is not zero, we write it to the cell; otherwise, we leave the cell empty to avoid displaying a zero value in the output worksheet.
                if (item.AmountDue != 0)
                    worksheet.Cell(row, 4).Value = item.AmountDue;

                // If the amount due is zero, we leave the cell empty to avoid displaying a zero value in the output worksheet.
                worksheet.Cell(row, 5).Value = item.AggregateAmountPaid;
                worksheet.Cell(row, 6).Value = item.Credit;
                worksheet.Cell(row, 7).Value = item.Notes;
                worksheet.Cell(row, 8).Value = item.Concat;
                worksheet.Cell(row, 9).Value = item.MonumentInvoice;

                // Apply a light gray background color to the row if the entry is not for a contractor.
                // This helps visually distinguish contractor entries from non-contractor entries in the output worksheet.
                if (!item.IsContractor)
                {
                    worksheet.Range(row, 1, row, 9).Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#F2F2F2"); // White Background 1, Darker 5%
                }
            }

            // After writing all the output rows to the worksheet, we apply clean formatting to the entire range of data, including the headers and all output rows.
            // This method sets the font, font size, borders, and other formatting options to ensure that the output worksheet is visually appealing and easy to read.
            ApplyCleanFormatting(worksheet, outputRows.Count + 1, headers.Length);

            // Adjust the column widths to fit the contents of the cells, ensuring that the output worksheet is visually appealing and easy to read.
            // This method automatically resizes the columns based on the content in each column.
            worksheet.Columns().AdjustToContents();


            // Calculate the total amount paid for the sheet by summing the AggregateAmountPaid values from all output rows. This total will be used for naming the worksheet and displaying the total amount paid in the output.
            decimal sheetTotal = outputRows.Sum(r => r.AggregateAmountPaid);

            // Format the sheet total as a currency string using the InvariantCulture to ensure consistent formatting across different locales.
            // The formatted string will be used for naming the worksheet and displaying the total amount paid in the output.
            string formattedSheetTotal = sheetTotal.ToString("$#,##0.00;-$#,##0.00", CultureInfo.InvariantCulture);

            // Set the worksheet name to a unique name based on the formatted sheet total, ensuring that it does not conflict with existing worksheet names in the workbook.
            // The MakeSafeWorksheetName method is used to sanitize the name and ensure it is valid for Excel.
            worksheet.Name = GetUniqueWorksheetName(
            worksheet.Workbook,
            MakeSafeWorksheetName($"{formattedSheetTotal} Check Total")
                );

            return sheetTotal;
        }

        // This method applies clean formatting to the specified range of cells in the worksheet, including font settings, borders, header styling, and number formatting.
        // It also highlights negative amounts in red and sets up an auto-filter for the data range.
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

        // This method finds the column index of the specified headerName in the given headerRow of the worksheet.
        // It iterates through the columns in the header row and compares the cell values with the provided headerName, returning the column index if a match is found. If no match is found, it returns -1.
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

        // This method retrieves the decimal value from the specified cell, handling various formatting cases such as currency symbols, commas, and parentheses for negative values.
        // It attempts to parse the cleaned string representation of the cell value into a decimal. If parsing is successful, it returns the decimal value; otherwise, it returns 0.
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

        // This method generates a unique output file path by checking if a file with the specified fileName already exists in the given folderPath.
        // If a file with the same name exists, it appends a numeric suffix (e.g., " (1)", " (2)", etc.) to the base name of the file until a unique file path is found. The method returns the unique output file path.
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

        // This method clears the contents of the specified range of rows in the worksheet, starting from the given startRow and spanning the specified number of rows.
        // It iterates through each row in the range and clears the contents of all cells in that row.
        private class ParsedDetails
        {
            public string Name { get; set; } = "";
            public string WeekEndingDate { get; set; } = "";
            public string Credit { get; set; } = "";
            public bool IsContractor { get; set; }
        }

        // This class represents a single output row in the Monument payment processing, containing various properties such as:
        // week ending date, name, invoice number, amount due, aggregate amount paid, credit, notes, concatenated string for matching, Monument invoice number, and a flag indicating whether the entry is for a contractor.
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