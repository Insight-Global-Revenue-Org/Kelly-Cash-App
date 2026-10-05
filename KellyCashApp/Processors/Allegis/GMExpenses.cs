using System.Globalization;
using Microsoft.VisualBasic.FileIO;

namespace KellyCashApp.Processors.Allegis
{
    internal static class GMExpense
    {
        public static List<GMExpenseMatch> Import(string filePath)
        {
            var matches = new List<GMExpenseMatch>();

            using var parser = new TextFieldParser(filePath);

            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");
            parser.HasFieldsEnclosedInQuotes = true;

            // ------------------------------------------------------------
            // Read header row first.
            // ------------------------------------------------------------
            if (parser.EndOfData)
                return matches;

            string[]? headers = parser.ReadFields();

            if (headers == null)
                return matches;

            // ------------------------------------------------------------
            // Dynamically locate the contractor-name column.
            //
            // Only use the column whose header CONTAINS:
            // "ProjectMilestonePayment"
            // ------------------------------------------------------------
            int contractorNameCol =
                FindHeaderContaining(
                    headers,
                    "ProjectMilestonePayment");

            // ------------------------------------------------------------
            // Dynamically locate the date column.
            //
            // Only use the column whose header CONTAINS:
            // "PaymentSubmittedDate"
            // ------------------------------------------------------------
            int paymentSubmittedDateCol =
                FindHeaderContaining(
                    headers,
                    "PaymentSubmittedDate");

            // ------------------------------------------------------------
            // Payment amount remains CSV column M.
            //
            // Excel/CSV:
            // M = column 13
            //
            // C# array index:
            // M = index 12
            // ------------------------------------------------------------
            const int paymentAmountCol = 12;

            // ------------------------------------------------------------
            // Required columns must exist.
            // ------------------------------------------------------------
            if (contractorNameCol == -1)
            {
                throw new Exception(
                    "GM Expense Report does not contain a header " +
                    "with 'ProjectMilestonePayment'.");
            }

            if (paymentSubmittedDateCol == -1)
            {
                throw new Exception(
                    "GM Expense Report does not contain a header " +
                    "with 'PaymentSubmittedDate'.");
            }

            // ------------------------------------------------------------
            // Read expense rows.
            // ------------------------------------------------------------
            while (!parser.EndOfData)
            {
                string[]? fields = parser.ReadFields();

                if (fields == null)
                    continue;

                int requiredMaxIndex =
                    Math.Max(
                        contractorNameCol,
                        Math.Max(
                            paymentAmountCol,
                            paymentSubmittedDateCol));

                if (fields.Length <= requiredMaxIndex)
                    continue;

                string contractorName =
                    fields[contractorNameCol].Trim();

                string rawAmount =
                    fields[paymentAmountCol].Trim();

                string rawDate =
                    fields[paymentSubmittedDateCol].Trim();

                if (string.IsNullOrWhiteSpace(contractorName))
                    continue;

                if (!TryParseAmount(
                    rawAmount,
                    out decimal amount))
                {
                    continue;
                }

                if (!DateTime.TryParse(
                    rawDate,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out DateTime date))
                {
                    if (!DateTime.TryParse(
                        rawDate,
                        out date))
                    {
                        continue;
                    }
                }

                matches.Add(
                    new GMExpenseMatch
                    {
                        ContractorName = contractorName,
                        Amount = amount,
                        Date = date.Date
                    });
            }

            return matches;
        }

        // ------------------------------------------------------------
        // Finds the first CSV column whose header contains
        // the supplied text.
        //
        // Example:
        //
        // "ProjectMilestonePayment.WorkerName"
        //
        // matches:
        //
        // "ProjectMilestonePayment"
        // ------------------------------------------------------------
        private static int FindHeaderContaining(
            string[] headers,
            string requiredText)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                string header =
                    headers[i]?.Trim() ?? "";

                if (header.Contains(
                    requiredText,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        private static bool TryParseAmount(
            string raw,
            out decimal amount)
        {
            raw = raw
                .Replace("$", "")
                .Replace(",", "")
                .Trim();

            if (raw.StartsWith("(") &&
                raw.EndsWith(")"))
            {
                raw =
                    "-" +
                    raw[1..^1];
            }

            return decimal.TryParse(
                raw,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out amount);
        }
    }

    internal class GMExpenseMatch
    {
        public string ContractorName { get; set; } = "";

        public decimal Amount { get; set; }

        public DateTime Date { get; set; }
    }
}