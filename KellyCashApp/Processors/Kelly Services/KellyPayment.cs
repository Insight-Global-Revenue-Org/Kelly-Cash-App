using ClosedXML.Excel;
using KellyCashApp.Configuration;
using KellyCashApp.Models;
using KellyCashApp.Services;
using System.Globalization;

namespace KellyCashApp.Processors.Kelly_Services
{
    internal static class KellyPayment
    {
        public static string Process(
            XLWorkbook workbook,
            IXLWorksheet worksheet,
            string inputPath,
            Dictionary<string, OirMatch> openInvoiceMatches)
        {
            const int headerRow = 13;

            int weekEndingColumn = FindColumn(
                worksheet,
                headerRow,
                "Week Ending Date");

            int contractorColumn = FindColumn(
                worksheet,
                headerRow,
                "Name");

            int lineTotalColumn = FindColumn(
                worksheet,
                headerRow,
                "Line Total");

            int locationDescriptionColumn = FindColumn(
                worksheet,
                headerRow,
                "Location Description");

            int userChar1Column = FindColumn(
                worksheet,
                headerRow,
                "User Char 1");

            if (weekEndingColumn == -1 ||
                contractorColumn == -1 ||
                lineTotalColumn == -1 ||
                locationDescriptionColumn == -1 ||
                userChar1Column == -1)
            {
                throw new InvalidOperationException(
                    "Could not find Week Ending Date, Name, " +
                    "Line Total, or Location Description.");
            }

            NormalizeColumns(
                worksheet,
                headerRow,
                contractorColumn);

            InsertBlankColumn(
                worksheet,
                contractorColumn + 1,
                "Name_",
                headerRow);

            // Re-find columns after inserting and normalizing.
            contractorColumn = FindColumn(
                worksheet,
                headerRow,
                "Name");

            weekEndingColumn = FindColumn(
                worksheet,
                headerRow,
                "Week Ending Date");

            lineTotalColumn = FindColumn(
                worksheet,
                headerRow,
                "Line Total");

            int invoiceColumn = FindColumnAfter(
                worksheet,
                headerRow,
                "Invoice",
                contractorColumn);

            int amountColumn = FindColumnAfter(
                worksheet,
                headerRow,
                "Amount",
                contractorColumn);

            int aggregateColumn = FindColumn(
                worksheet,
                headerRow,
                "Aggregate Line Total");

            int notesColumn = FindColumn(
                worksheet,
                headerRow,
                "Notes");

            int nameFormattedColumn = FindColumnAfter(
                worksheet,
                headerRow,
                "Name_",
                contractorColumn);

            InsertBlankColumn(
                worksheet,
                lineTotalColumn + 1,
                "Concat",
                headerRow);

            // Re-find these because inserting Concat changes positions.
            lineTotalColumn = FindColumn(
                worksheet,
                headerRow,
                "Line Total");

            int concatColumn = FindColumnAfter(
                worksheet,
                headerRow,
                "Concat",
                lineTotalColumn);

            if (invoiceColumn == -1 ||
                amountColumn == -1 ||
                aggregateColumn == -1 ||
                notesColumn == -1 ||
                nameFormattedColumn == -1 ||
                concatColumn == -1 ||
                weekEndingColumn == -1 ||
                contractorColumn == -1 ||
                lineTotalColumn == -1)
            {
                throw new InvalidOperationException(
                    "Could not find one or more required Kelly payment columns.");
            }

            int lastRow =
                worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            PopulateFormattedNamesAndConcatValues(
                worksheet,
                headerRow,
                lastRow,
                contractorColumn,
                weekEndingColumn,
                nameFormattedColumn,
                concatColumn);

            PopulateAggregateTotalsAndInvoiceMatches(
                worksheet,
                headerRow,
                lastRow,
                contractorColumn,
                weekEndingColumn,
                lineTotalColumn,
                aggregateColumn,
                nameFormattedColumn,
                invoiceColumn,
                amountColumn,
                userChar1Column,
                openInvoiceMatches);

            string downloadsPath =
                Settings.GetRemittanceSavePath();

            string companyName = worksheet
                .Cell(headerRow + 1, locationDescriptionColumn)
                .GetString()
                .Trim();

            if (string.IsNullOrWhiteSpace(companyName))
            {
                companyName = "Kelly Remittance Payment";
            }

            companyName = MakeSafeFileName(companyName);

            decimal totalAmount = GetDecimalValue(
                worksheet.Cell(lastRow, lineTotalColumn));

            string formattedTotal = totalAmount.ToString(
                "N2",
                CultureInfo.InvariantCulture);

            string outputPath = GetUniqueOutputPath(
                downloadsPath,
                $"{companyName} - {formattedTotal}.xlsx");

            FormatWorksheet(
                worksheet,
                headerRow,
                lastRow,
                invoiceColumn,
                amountColumn,
                aggregateColumn,
                notesColumn,
                nameFormattedColumn,
                concatColumn,
                lineTotalColumn);

            workbook.SaveAs(outputPath);

            Analytics.LogRemittanceRun(
                $"{companyName} - {formattedTotal}");

            return outputPath;
        }

        private static void PopulateFormattedNamesAndConcatValues(
            IXLWorksheet worksheet,
            int headerRow,
            int lastRow,
            int contractorColumn,
            int weekEndingColumn,
            int nameFormattedColumn,
            int concatColumn)
        {
            var nameChanges = Rename.LoadNameChanges();

            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string rawName = worksheet
                    .Cell(row, contractorColumn)
                    .GetString()
                    .Trim();

                if (string.IsNullOrWhiteSpace(rawName))
                {
                    continue;
                }

                string formattedName = FormatName(rawName);

                formattedName = Rename.ApplyNameChange(
                    formattedName,
                    nameChanges);

                string formattedWeekEnding = FormatWeekEndingDate(
                    worksheet.Cell(row, weekEndingColumn));

                string concatValue =
                    $"{formattedName} {formattedWeekEnding}".Trim();

                worksheet.Cell(row, nameFormattedColumn).Value =
                    formattedName;

                worksheet.Cell(row, concatColumn).Value =
                    concatValue;
            }
        }

        private static void PopulateAggregateTotalsAndInvoiceMatches(
            IXLWorksheet worksheet,
            int headerRow,
            int lastRow,
            int contractorColumn,
            int weekEndingColumn,
            int lineTotalColumn,
            int aggregateColumn,
            int nameFormattedColumn,
            int invoiceColumn,
            int amountColumn,
            int userChar1Column,
            Dictionary<string, OirMatch> openInvoiceMatches)
        {
            var totalsByContractorAndWeek =
                new Dictionary<string, decimal>();

            var lastRowByContractorAndWeek =
                new Dictionary<string, int>();

            var matchedInvoiceNumbers =
                new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Include the final used row.
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string contractor = worksheet
                    .Cell(row, contractorColumn)
                    .GetString()
                    .Trim();

                string weekEnding = worksheet
                    .Cell(row, weekEndingColumn)
                    .GetString()
                    .Trim();

                if (string.IsNullOrWhiteSpace(contractor) ||
                    string.IsNullOrWhiteSpace(weekEnding))
                {
                    continue;
                }

                decimal lineTotal = GetDecimalValue(
                    worksheet.Cell(row, lineTotalColumn));

                string key = $"{contractor}|{weekEnding}";

                if (!totalsByContractorAndWeek.ContainsKey(key))
                {
                    totalsByContractorAndWeek[key] = 0;
                }

                totalsByContractorAndWeek[key] += lineTotal;
                lastRowByContractorAndWeek[key] = row;
            }

            foreach (var item in totalsByContractorAndWeek)
            {
                string key = item.Key;
                decimal aggregateTotal = item.Value;
                int targetRow = lastRowByContractorAndWeek[key];

                IXLCell aggregateCell =
                    worksheet.Cell(targetRow, aggregateColumn);

                aggregateCell.Value = aggregateTotal;

                aggregateCell.Style =
                    worksheet.Cell(targetRow, lineTotalColumn).Style;

                aggregateCell.Style.NumberFormat.Format =
                    "$#,##0.00;($#,##0.00)";

                if (aggregateTotal < 0)
                {
                    aggregateCell.Style.Font.FontColor =
                        XLColor.Red;
                }

                string formattedName = worksheet
                    .Cell(targetRow, nameFormattedColumn)
                    .GetString()
                    .Trim();

                string formattedWeekEnding = FormatWeekEndingDate(
                    worksheet.Cell(targetRow, weekEndingColumn));

                string userChar1 = worksheet
                    .Cell(targetRow, userChar1Column)
                    .GetString()
                    .Trim();

                bool isLabor =
                    userChar1.Trim()
                    .StartsWith("LABOR",
                 StringComparison.OrdinalIgnoreCase);

                if (!TryMatchWithDateSpread(
                    formattedName,
                    formattedWeekEnding,
                    isLabor,
                    aggregateTotal,
                    openInvoiceMatches,
                    matchedInvoiceNumbers,
                    out OirMatch? match))
                {
                    continue;
                }

                matchedInvoiceNumbers.Add(match.DocumentNumber);

                worksheet.Cell(targetRow, invoiceColumn).Value =
                    match.DocumentNumber;

                IXLCell amountCell =
                    worksheet.Cell(targetRow, amountColumn);

                amountCell.Value = match.RemainingAmount;

                amountCell.Style =
                    worksheet.Cell(targetRow, lineTotalColumn).Style;

                amountCell.Style.NumberFormat.Format =
                    "$#,##0.00;($#,##0.00)";
            }
        }

        private static void FormatWorksheet(
            IXLWorksheet worksheet,
            int headerRow,
            int lastRow,
            int invoiceColumn,
            int amountColumn,
            int aggregateColumn,
            int notesColumn,
            int nameFormattedColumn,
            int concatColumn,
            int lineTotalColumn)
        {
            int lastColumn =
                worksheet.LastColumnUsed()?.ColumnNumber()
                ?? lineTotalColumn;

            if (worksheet.AutoFilter.IsEnabled)
            {
                worksheet.AutoFilter.Clear();
            }

            worksheet
                .Range(headerRow, 1, lastRow, lastColumn)
                .SetAutoFilter();

            worksheet.Column(invoiceColumn).Width = 16;
            worksheet.Column(amountColumn).Width = 14;
            worksheet.Column(aggregateColumn).Width = 20;
            worksheet.Column(notesColumn).Width = 24;
            worksheet.Column(nameFormattedColumn).Width = 18;
            worksheet.Column(concatColumn).Width = 28;
        }

        private static int FindColumn(
            IXLWorksheet worksheet,
            int headerRow,
            string headerName)
        {
            for (int column = 1; column <= 100; column++)
            {
                string headerText = worksheet
                    .Cell(headerRow, column)
                    .GetString()
                    .Trim();

                if (headerText.Equals(
                    headerName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return column;
                }
            }

            return -1;
        }

        private static int FindColumnAfter(
            IXLWorksheet worksheet,
            int headerRow,
            string headerName,
            int afterColumn)
        {
            for (int column = afterColumn + 1;
                 column <= 100;
                 column++)
            {
                string headerText = worksheet
                    .Cell(headerRow, column)
                    .GetString()
                    .Trim();

                if (headerText.Equals(
                    headerName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return column;
                }
            }

            return -1;
        }

        private static void NormalizeColumns(
            IXLWorksheet worksheet,
            int headerRow,
            int contractorColumn)
        {
            string[] desiredColumns =
            {
                "Invoice",
                "Amount",
                "Aggregate Line Total",
                "Notes"
            };

            int insertAt = contractorColumn + 1;

            foreach (string columnName in desiredColumns)
            {
                string currentHeader = worksheet
                    .Cell(headerRow, insertAt)
                    .GetString()
                    .Trim();

                if (!currentHeader.Equals(
                    columnName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    InsertBlankColumn(
                        worksheet,
                        insertAt,
                        columnName,
                        headerRow);
                }

                insertAt++;
            }
        }

        private static void InsertBlankColumn(
            IXLWorksheet worksheet,
            int targetColumn,
            string headerName,
            int headerRow)
        {
            int lastRow =
                worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            int lastColumn =
                worksheet.LastColumnUsed()?.ColumnNumber()
                ?? targetColumn;

            for (int row = 1; row <= lastRow; row++)
            {
                for (int column = lastColumn;
                     column >= targetColumn;
                     column--)
                {
                    worksheet
                        .Cell(row, column + 1)
                        .CopyFrom(worksheet.Cell(row, column));

                    worksheet.Cell(row, column).Clear();
                }
            }

            worksheet.Cell(headerRow, targetColumn).Value =
                headerName;

            worksheet.Cell(headerRow, targetColumn).Style =
                worksheet.Cell(headerRow, targetColumn - 1).Style;

            IXLCell designCell =
                worksheet.Cell(headerRow - 1, targetColumn);

            IXLCell leftCell =
                worksheet.Cell(headerRow - 1, targetColumn - 1);

            designCell.Style = leftCell.Style;

            leftCell.Style.Border.RightBorder =
                XLBorderStyleValues.None;
        }

        private static decimal GetDecimalValue(IXLCell cell)
        {
            string rawValue = cell.Value
                .ToString()
                .Replace("$", "")
                .Replace(",", "")
                .Trim();

            if (decimal.TryParse(
                rawValue,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out decimal result))
            {
                return result;
            }

            return 0;
        }

        private static string FormatName(string input)
        {
            if (!input.Contains(','))
            {
                return input;
            }

            string[] parts = input.Split(',');

            if (parts.Length < 2)
            {
                return input;
            }

            string last = parts[0].Trim().ToLower();
            string first = parts[1].Trim().ToLower();

            TextInfo textInfo =
                CultureInfo.CurrentCulture.TextInfo;

            last = textInfo.ToTitleCase(last);
            first = textInfo.ToTitleCase(first);

            return $"{first} {last}";
        }

        private static string FormatWeekEndingDate(IXLCell cell)
        {
            if (cell.Value.IsDateTime)
            {
                return cell
                    .GetDateTime()
                    .ToString(
                        "MM/dd/yyyy",
                        CultureInfo.InvariantCulture);
            }

            string rawValue = cell.GetString().Trim();

            if (DateTime.TryParse(
                rawValue,
                out DateTime parsedDate))
            {
                return parsedDate.ToString(
                    "MM/dd/yyyy",
                    CultureInfo.InvariantCulture);
            }

            return rawValue;
        }

        private static bool TryMatchWithDateSpread(
     string formattedName,
     string formattedWeekEnding,
     bool isLabor,
     decimal aggregateAmount,
     Dictionary<string, OirMatch> openInvoiceMatches,
     HashSet<string> matchedInvoiceNumbers,
     out OirMatch? match)
        {
            match = null;

            if (!DateTime.TryParse(
                formattedWeekEnding,
                out DateTime baseDate))
            {
                return false;
            }

            // Labor keeps its existing ±2-day search and does not require
            // an amount comparison.
            if (isLabor)
            {
                return TryFindMatchInRange(
                    formattedName,
                    baseDate,
                    startOffset: -2,
                    endOffset: 2,
                    requireExactAmount: false,
                    aggregateAmount,
                    openInvoiceMatches,
                    matchedInvoiceNumbers,
                    out match);
            }

            // Expense pass 1: search within ±7 days for an exact amount.
            if (TryFindMatchInRange(
                formattedName,
                baseDate,
                startOffset: -7,
                endOffset: 7,
                requireExactAmount: true,
                aggregateAmount,
                openInvoiceMatches,
                matchedInvoiceNumbers,
                out match))
            {
                return true;
            }

            // Expense pass 2: search only the additional five days,
            // expanding the total range from ±7 to ±12.
            if (TryFindMatchInRange(
                formattedName,
                baseDate,
                startOffset: -12,
                endOffset: -8,
                requireExactAmount: true,
                aggregateAmount,
                openInvoiceMatches,
                matchedInvoiceNumbers,
                out match))
            {
                return true;
            }

            return TryFindMatchInRange(
                formattedName,
                baseDate,
                startOffset: 8,
                endOffset: 12,
                requireExactAmount: true,
                aggregateAmount,
                openInvoiceMatches,
                matchedInvoiceNumbers,
                out match);
        }

        private static bool TryFindMatchInRange(
            string formattedName,
            DateTime baseDate,
            int startOffset,
            int endOffset,
            bool requireExactAmount,
            decimal aggregateAmount,
            Dictionary<string, OirMatch> openInvoiceMatches,
            HashSet<string> matchedInvoiceNumbers,
            out OirMatch? match)
        {
            match = null;

            for (int offset = startOffset; offset <= endOffset; offset++)
            {
                DateTime testDate = baseDate.AddDays(offset);

                string testDateString = testDate.ToString(
                    "MM/dd/yyyy",
                    CultureInfo.InvariantCulture);

                string testKey =
                    $"{formattedName} {testDateString}".Trim();

                if (!openInvoiceMatches.TryGetValue(
                    testKey,
                    out OirMatch? foundMatch))
                {
                    continue;
                }

                if (matchedInvoiceNumbers.Contains(foundMatch.DocumentNumber))
                {
                    continue;
                }

                if (requireExactAmount &&
                    foundMatch.RemainingAmount != aggregateAmount)
                {
                    continue;
                }

                match = foundMatch;
                return true;
            }

            return false;
        }

        private static string MakeSafeFileName(string fileName)
        {
            foreach (char invalidCharacter
                     in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(
                    invalidCharacter,
                    '_');
            }

            return fileName.Trim();
        }

        private static string GetUniqueOutputPath(
            string folderPath,
            string fileName)
        {
            string fileNameWithoutExtension =
                Path.GetFileNameWithoutExtension(fileName);

            string extension =
                Path.GetExtension(fileName);

            string outputPath =
                Path.Combine(folderPath, fileName);

            int counter = 1;

            while (File.Exists(outputPath))
            {
                outputPath = Path.Combine(
                    folderPath,
                    $"{fileNameWithoutExtension} ({counter}){extension}");

                counter++;
            }

            return outputPath;
        }
    }
}