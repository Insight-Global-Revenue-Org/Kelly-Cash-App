using ClosedXML.Excel;
using KellyCashApp.Configuration;
using KellyCashApp.Services;

namespace KellyCashApp.Workflows
{
    public class OIR
    {
        public static void ShowNotesMenu(int menuTop)
        {
            while (true)
            {
                int selected = ShowMenu(new[]
                {
                    "Pull Forward Open Invoice Report",
                    "Pull Forward Full Cash Report",
                    "Back"
                }, 0, menuTop);

                if (selected == 2)
                    return;

                int promptTop = menuTop + 6;

                ClearArea(promptTop, 10);
                Console.SetCursorPosition(0, promptTop);

                if (selected == 0)
                {
                    RunOpenInvoicePullForward(promptTop);
                }

                if (selected == 1)
                {
                    UAC.RunFullCashReportPullForward(promptTop);
                }

                ClearArea(promptTop, 12);
            }
        }

        private static void RunOpenInvoicePullForward(int promptTop)
        {
            string? newOirPath = FileSelector.SelectFile(
                "Paste the full file path of your NEW Open Invoice Report:",
                promptTop
            ); ;

            if (string.IsNullOrWhiteSpace(newOirPath) || !File.Exists(newOirPath))
            {
                Console.WriteLine("File not found. Please check the NEW OIR path.");
                Console.WriteLine("Press any key to return...");
                Console.ReadKey(true);
                return;
            }

            string? oldOirPath = FileSelector.SelectFile(
                "Paste the full file path of your OLD / notated Open Invoice Report:",
                 promptTop
            );

            if (string.IsNullOrWhiteSpace(oldOirPath) || !File.Exists(oldOirPath))
            {
                Console.WriteLine("File not found. Please check the OLD OIR path.");
                Console.WriteLine("Press any key to return...");
                Console.ReadKey(true);
                return;
            }

            try
            {
                bool loading = true;
                List<string> fallenOffInvoices = new();

                Task spinner = Task.Run(() =>
                {
                    char[] frames = { '/', '-', '\\', '|' };
                    int i = 0;

                    while (loading)
                    {
                        Console.SetCursorPosition(0, promptTop);
                        Console.Write($"Preparing Open Invoice Report... {frames[i++ % frames.Length]}   ");
                        Thread.Sleep(120);
                    }
                });

                string outputPath = ProcessOpenInvoiceReport(newOirPath, oldOirPath, out fallenOffInvoices);
                Analytics.LogOirRun(fallenOffInvoices.Count);

                loading = false;
                spinner.Wait();

                ClearArea(promptTop, 18);
                Console.SetCursorPosition(0, promptTop);

                Console.WriteLine("Open Invoice Report prepared successfully.");
                Console.WriteLine($"Saved to: {outputPath}");
                Console.WriteLine();

                ShowFallenOffInvoiceMarquee(fallenOffInvoices, promptTop + 4);

                Console.WriteLine();
                Console.WriteLine("Press any key to return to menu...");
                Console.ReadKey(true);
            }
            catch (Exception ex)
            {
                ClearArea(promptTop, 10);
                Console.SetCursorPosition(0, promptTop);

                Console.WriteLine("Something went wrong while preparing the OIR:");
                Console.WriteLine(ex.Message);
                Console.WriteLine();
                Console.WriteLine("Press any key to return...");
                Console.ReadKey(true);
            }
        }

        private static string ProcessOpenInvoiceReport(string newOirPath, string oldOirPath, out List<string> fallenOffInvoices)
        {
            using var stream = new FileStream(newOirPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.First();

            int headerRow = FindHeaderRow(worksheet, "Consultant");

            if (headerRow == -1)
            {
                throw new Exception("Could not locate header row (Consultant not found).");
            }

            // If header is BELOW row 1, clean up rows above it
            if (headerRow > 1)
            {
                worksheet.Rows(1, headerRow - 1).Delete();
                headerRow = 1;
            }

            DeleteColumnIfExists(worksheet, headerRow, "A/R Rep");

            string[] columnsToMove =
            {
                "Consultant",
                "Service End",
                "Document Amount",
                "Remaining Amount",
                "Percent Open",
                "Document Number",
                "Aging Bucket"
            };

            MoveColumnsAfter(worksheet, headerRow, "Client ID", columnsToMove);

            string[] calculatedColumns =
{
    "AR Status",
    "Name/WE"
};

            AddColumnsToEnd(worksheet, headerRow, calculatedColumns);

            string[] oldNoteColumns = GetOldOirNoteColumns(oldOirPath);

            AddColumnsToEnd(worksheet, headerRow, oldNoteColumns);

            ApplyOpenInvoiceFormattingAndFormulas(worksheet, headerRow);

            fallenOffInvoices = PullForwardOldOirNotes(worksheet, headerRow, oldOirPath, oldNoteColumns);

            CleanBlankNoteValues(worksheet, headerRow);
            ApplyFinalOirStyling(worksheet, headerRow);

            worksheet.Columns().AdjustToContents();

            string downloadsPath = Settings.GetOpenInvoiceReportSavePath();

            string today = DateTime.Now.ToString("M.d.yyyy");
            string outputPath = GetUniqueOutputPath(downloadsPath, $"OIR {today}.xlsx");

            workbook.SaveAs(outputPath);

            return outputPath;
        }

        private static void DeleteColumnIfExists(IXLWorksheet worksheet, int headerRow, string headerName)
        {
            int column = FindColumn(worksheet, headerRow, headerName);

            if (column != -1)
                worksheet.Column(column).Delete();
        }

        private static void MoveColumnsAfter(IXLWorksheet worksheet, int headerRow, string afterHeaderName, string[] columnsToMove)
        {
            var snapshots = new List<ColumnSnapshot>();

            foreach (string columnName in columnsToMove)
            {
                int column = FindColumn(worksheet, headerRow, columnName);

                if (column == -1)
                    throw new Exception($"Could not find required column: {columnName}");

                snapshots.Add(CaptureColumn(worksheet, column, columnName));
            }

            foreach (string columnName in columnsToMove.Reverse())
            {
                int column = FindColumn(worksheet, headerRow, columnName);

                if (column != -1)
                    worksheet.Column(column).Delete();
            }

            int clientIdColumn = FindColumn(worksheet, headerRow, afterHeaderName);

            if (clientIdColumn == -1)
                throw new Exception($"Could not find column: {afterHeaderName}");

            worksheet.Column(clientIdColumn).InsertColumnsAfter(snapshots.Count);

            for (int i = 0; i < snapshots.Count; i++)
            {
                int targetColumn = clientIdColumn + 1 + i;
                RestoreColumn(worksheet, targetColumn, snapshots[i]);
            }
        }

        private static ColumnSnapshot CaptureColumn(IXLWorksheet worksheet, int columnNumber, string headerName)
        {
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            var values = new List<string>();

            for (int row = 1; row <= lastRow; row++)
            {
                values.Add(worksheet.Cell(row, columnNumber).GetString());
            }

            return new ColumnSnapshot(headerName, values);
        }

        private static void RestoreColumn(IXLWorksheet worksheet, int columnNumber, ColumnSnapshot snapshot)
        {
            for (int row = 1; row <= snapshot.Values.Count; row++)
            {
                worksheet.Cell(row, columnNumber).Value = snapshot.Values[row - 1];
            }
        }

        private static void AddColumnsToEnd(IXLWorksheet worksheet, int headerRow, string[] newColumns)
        {
            int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            int targetColumn = lastColumn + 1;

            foreach (string columnName in newColumns)
            {
                worksheet.Cell(headerRow, targetColumn).Value = columnName;
                targetColumn++;
            }
        }

        private static int FindColumn(IXLWorksheet worksheet, int headerRow, string headerName)
        {
            int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 100;

            for (int col = 1; col <= lastColumn; col++)
            {
                string text = worksheet.Cell(headerRow, col).GetString().Trim();

                if (text.Equals(headerName, StringComparison.OrdinalIgnoreCase))
                    return col;
            }

            return -1;
        }

        private static string? PromptForFilePath(string prompt, int startLine)
        {
            Console.CursorVisible = true;

            ClearArea(startLine, 8);
            Console.SetCursorPosition(0, startLine);

            Console.WriteLine(prompt);
            Console.Write("> ");

            string? path = Console.ReadLine()?.Trim().Trim('"');

            ClearArea(startLine, 8);

            return path;
        }

        private static string GetUniqueOutputPath(string folderPath, string fileName)
        {
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);

            string outputPath = Path.Combine(folderPath, fileName);
            int counter = 1;

            while (File.Exists(outputPath))
            {
                outputPath = Path.Combine(folderPath, $"{fileNameWithoutExt} ({counter}){extension}");
                counter++;
            }

            return outputPath;
        }

        private static string[] GetOldOirNoteColumns(string oldOirPath)
        {
            using var stream = new FileStream(oldOirPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.First();

            int headerRow = 1;

            string[] excludedHeaders =
            {
        "Client ID",
        "Consultant",
        "Service End",
        "Document Amount",
        "Remaining Amount",
        "Percent Open",
        "Document Number",
        "Aging Bucket",
        "AR Status",
        "Name/WE",
        "Document Type Name",
        "GL Posting Date",
        "Date Posted",
        "Account Manager",
        "AM Email",
        "Office",
        "Client Projects"
    };

            int agingBucketColumn = FindColumn(worksheet, headerRow, "Aging Bucket");

            if (agingBucketColumn == -1)
                throw new Exception("Could not find Aging Bucket column in the old OIR.");

            int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? agingBucketColumn;

            var noteColumns = new List<string>();

            for (int col = agingBucketColumn + 1; col <= lastColumn; col++)
            {
                string header = worksheet.Cell(headerRow, col).GetString().Trim();

                if (string.IsNullOrWhiteSpace(header))
                    continue;

                if (excludedHeaders.Contains(header, StringComparer.OrdinalIgnoreCase))
                    continue;

                noteColumns.Add(header);
            }

            return noteColumns.ToArray();
        }

        private static List<string> PullForwardOldOirNotes(
     IXLWorksheet newWorksheet,
     int newHeaderRow,
     string oldOirPath,
     string[] noteColumnsToPull
 )
        {
            using var oldStream = new FileStream(oldOirPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var oldWorkbook = new XLWorkbook(oldStream);
            var oldWorksheet = oldWorkbook.Worksheets.First();

            int oldHeaderRow = FindHeaderRow(oldWorksheet, "Document Number");

            if (oldHeaderRow == -1)
                throw new Exception("Could not find Document Number header row in old OIR.");

            int oldDocumentNumberColumn = FindColumn(oldWorksheet, oldHeaderRow, "Document Number");
            int newDocumentNumberColumn = FindColumn(newWorksheet, newHeaderRow, "Document Number");

            if (oldDocumentNumberColumn == -1)
                throw new Exception("Could not find Document Number column in old OIR.");

            if (newDocumentNumberColumn == -1)
                throw new Exception("Could not find Document Number column in new OIR.");

            int oldLastRow = oldWorksheet.LastRowUsed()?.RowNumber() ?? oldHeaderRow;
            int newLastRow = newWorksheet.LastRowUsed()?.RowNumber() ?? newHeaderRow;

            var oldRowsByDocumentNumber = new Dictionary<string, int>();

            for (int oldRow = oldHeaderRow + 1; oldRow <= oldLastRow; oldRow++)
            {
                string oldDocumentNumber = NormalizeDocumentNumber(
                    oldWorksheet.Cell(oldRow, oldDocumentNumberColumn).Value.ToString()
                );

                if (string.IsNullOrWhiteSpace(oldDocumentNumber))
                    continue;

                if (!oldRowsByDocumentNumber.ContainsKey(oldDocumentNumber))
                    oldRowsByDocumentNumber.Add(oldDocumentNumber, oldRow);
            }

            for (int newRow = newHeaderRow + 1; newRow <= newLastRow; newRow++)
            {
                string newDocumentNumber = NormalizeDocumentNumber(
                    newWorksheet.Cell(newRow, newDocumentNumberColumn).Value.ToString()
                );

                if (!oldRowsByDocumentNumber.TryGetValue(newDocumentNumber, out int oldRow))
                    continue;

                foreach (string noteColumnName in noteColumnsToPull)
                {
                    int oldNoteColumn = FindColumn(oldWorksheet, oldHeaderRow, noteColumnName);
                    int newNoteColumn = FindColumn(newWorksheet, newHeaderRow, noteColumnName);

                    if (oldNoteColumn == -1 || newNoteColumn == -1)
                        continue;

                    var pulledValue = oldWorksheet.Cell(oldRow, oldNoteColumn).Value;


                    newWorksheet.Cell(newRow, newNoteColumn).Value = pulledValue;
                }
            }
            var newDocumentNumbers = new HashSet<string>();

            for (int newRow = newHeaderRow + 1; newRow <= newLastRow; newRow++)
            {
                string newDocumentNumber = NormalizeDocumentNumber(
                    newWorksheet.Cell(newRow, newDocumentNumberColumn).Value.ToString()
                );

                if (!string.IsNullOrWhiteSpace(newDocumentNumber))
                    newDocumentNumbers.Add(newDocumentNumber);
            }

            var fallenOffInvoices = oldRowsByDocumentNumber.Keys
                .Where(oldDoc => !newDocumentNumbers.Contains(oldDoc))
                .OrderBy(oldDoc => oldDoc)
                .ToList();

            return fallenOffInvoices;
        }

        private static void ShowFallenOffInvoiceMarquee(List<string> fallenOffInvoices, int startLine)
        {
            Console.SetCursorPosition(0, startLine);

            if (fallenOffInvoices.Count == 0)
            {
                Console.WriteLine("No invoices fell off the new OIR.");
                return;
            }

            Console.WriteLine($"Invoices no longer on new OIR: {fallenOffInvoices.Count}");
            Console.WriteLine("Likely applied / closed out:");
            Console.WriteLine();

            int maxToShow = Math.Min(fallenOffInvoices.Count, 6);

            for (int i = 0; i < maxToShow; i++)
            {
                Console.WriteLine($"  - {fallenOffInvoices[i]}");
            }

            if (fallenOffInvoices.Count > maxToShow)
            {
                Console.WriteLine($"  ...and {fallenOffInvoices.Count - maxToShow} more");
            }
        }

        private static string NormalizeDocumentNumber(string value)
        {
            value = value.Trim();

            if (double.TryParse(value, out double number))
                return number.ToString("0");

            return value;
        }

        private static void ApplyOpenInvoiceFormattingAndFormulas(IXLWorksheet worksheet, int headerRow)
        {
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            int consultantColumn = FindColumn(worksheet, headerRow, "Consultant");
            int serviceEndColumn = FindColumn(worksheet, headerRow, "Service End");
            if (serviceEndColumn != -1)
            {
                worksheet.Column(serviceEndColumn).Style.NumberFormat.Format = "m/d/yyyy";

                for (int row = headerRow + 1; row <= lastRow; row++)
                {
                    var cell = worksheet.Cell(row, serviceEndColumn);
                    string raw = cell.GetString().Trim();

                    if (DateTime.TryParse(raw, out DateTime parsedDate))
                    {
                        cell.Value = parsedDate.Date;
                        cell.Style.NumberFormat.Format = "m/d/yyyy";
                    }
                }
            }
            int documentAmountColumn = FindColumn(worksheet, headerRow, "Document Amount");
            int remainingAmountColumn = FindColumn(worksheet, headerRow, "Remaining Amount");
            int percentOpenColumn = FindColumn(worksheet, headerRow, "Percent Open");
            int documentNumberColumn = FindColumn(worksheet, headerRow, "Document Number");
            int arStatusColumn = FindColumn(worksheet, headerRow, "AR Status");
            int nameWeColumn = FindColumn(worksheet, headerRow, "Name/WE");

            if (documentAmountColumn != -1)
            {
                worksheet.Column(documentAmountColumn).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";

                for (int row = headerRow + 1; row <= lastRow; row++)
                {
                    string raw = worksheet.Cell(row, documentAmountColumn)
                        .GetString()
                        .Replace("$", "")
                        .Replace(",", "")
                        .Trim();

                    if (decimal.TryParse(raw, out decimal value))
                    {
                        worksheet.Cell(row, documentAmountColumn).Value = value;
                    }
                }
            }

            if (remainingAmountColumn != -1)
            {
                worksheet.Column(remainingAmountColumn).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";

                for (int row = headerRow + 1; row <= lastRow; row++)
                {
                    string raw = worksheet.Cell(row, remainingAmountColumn)
                        .GetString()
                        .Replace("$", "")
                        .Replace(",", "")
                        .Trim();

                    if (decimal.TryParse(raw, out decimal value))
                    {
                        worksheet.Cell(row, remainingAmountColumn).Value = value;
                    }
                }
            }

            if (percentOpenColumn != -1)
            {
                worksheet.Column(percentOpenColumn).Style.NumberFormat.Format = "0.00%";

                for (int row = headerRow + 1; row <= lastRow; row++)
                {
                    string raw = worksheet.Cell(row, percentOpenColumn)
                        .GetString()
                        .Replace("%", "")
                        .Trim();

                    if (decimal.TryParse(raw, out decimal value))
                    {
                        if (value > 1)
                            value /= 100m;

                        worksheet.Cell(row, percentOpenColumn).Value = value;
                    }
                }
            }

            if (documentNumberColumn != -1)
            {
                worksheet.Column(documentNumberColumn).Style.NumberFormat.Format = "0";

                for (int row = headerRow + 1; row <= lastRow; row++)
                {
                    string rawValue = worksheet.Cell(row, documentNumberColumn).GetString().Trim();

                    if (double.TryParse(rawValue, out double documentNumber))
                    {
                        worksheet.Cell(row, documentNumberColumn).Value = documentNumber;
                    }
                }
            }

            if (arStatusColumn != -1 && documentAmountColumn != -1 && remainingAmountColumn != -1)
            {
                string documentAmountLetter = worksheet.Column(documentAmountColumn).ColumnLetter();
                string remainingAmountLetter = worksheet.Column(remainingAmountColumn).ColumnLetter();

                for (int row = headerRow + 1; row <= lastRow; row++)
                {
                    worksheet.Cell(row, arStatusColumn).FormulaA1 =
                        $"=IF({documentAmountLetter}{row}={remainingAmountLetter}{row},\"Fully Applied\",\"Partially Applied\")";
                }
            }

            if (nameWeColumn != -1 && consultantColumn != -1 && serviceEndColumn != -1)
            {
                string consultantLetter = worksheet.Column(consultantColumn).ColumnLetter();
                string serviceEndLetter = worksheet.Column(serviceEndColumn).ColumnLetter();

                for (int row = headerRow + 1; row <= lastRow; row++)
                {
                    worksheet.Cell(row, nameWeColumn).FormulaA1 =
                        $"={consultantLetter}{row}&\" \"&TEXT({serviceEndLetter}{row},\"mm/dd/yyyy\")";
                }
            }
        }

        private static int FindHeaderRow(IXLWorksheet worksheet, string requiredHeader)
        {
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 100;

            for (int row = 1; row <= lastRow; row++)
            {
                for (int col = 1; col <= 100; col++)
                {
                    string cellText = worksheet.Cell(row, col).GetString().Trim();

                    if (cellText.Equals(requiredHeader, StringComparison.OrdinalIgnoreCase))
                        return row;
                }
            }

            return -1;
        }

        private static int ShowMenu(string[] options, int defaultSelected, int menuTop)
        {
            int selected = defaultSelected;
            int menuWidth = options.Max(o => o.Length) + 4;
            int neededRows = options.Length + 2;

            if (menuTop + neededRows >= Console.BufferHeight)
            {
                menuTop = Math.Max(0, Console.BufferHeight - neededRows);
            }

            Console.CursorVisible = false;
            DrawFullMenu(options, selected, menuTop, menuWidth);

            while (true)
            {
                ConsoleKey key = Console.ReadKey(true).Key;

                if (key == ConsoleKey.Enter)
                    break;

                int oldSelected = selected;

                if (key == ConsoleKey.UpArrow)
                    selected = selected == 0 ? options.Length - 1 : selected - 1;

                if (key == ConsoleKey.DownArrow)
                    selected = selected == options.Length - 1 ? 0 : selected + 1;

                if (selected != oldSelected)
                {
                    DrawFullMenu(options, selected, menuTop, menuWidth);
                }
            }

            Console.ResetColor();
            Console.CursorVisible = true;

            int afterMenu = Math.Min(menuTop + neededRows, Console.BufferHeight - 1);
            Console.SetCursorPosition(0, afterMenu);
            Console.WriteLine();

            return selected;
        }

        private static void DrawFullMenu(string[] options, int selected, int menuTop, int menuWidth)
        {
            Console.ResetColor();
            Console.SetCursorPosition(0, menuTop);
            Console.WriteLine("Select".PadRight(menuWidth));

            for (int i = 0; i < options.Length; i++)
            {
                Console.SetCursorPosition(0, menuTop + i + 1);

                string line = i == selected
                    ? $"> {options[i]}"
                    : $"  {options[i]}";

                if (i == selected)
                {
                    Console.ForegroundColor = ConsoleColor.Black;
                    Console.BackgroundColor = ConsoleColor.Gray;
                    Console.Write(line.PadRight(menuWidth));
                    Console.ResetColor();
                }
                else
                {
                    Console.Write(line.PadRight(menuWidth));
                }
            }
        }

        private static void CleanBlankNoteValues(IXLWorksheet worksheet, int headerRow)
        {
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;
            int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;

            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                for (int col = 1; col <= lastColumn; col++)
                {
                    string value = worksheet.Cell(row, col).GetString().Trim();

                    if (value.Equals("#N/A", StringComparison.OrdinalIgnoreCase) || value == "0")
                    {
                        worksheet.Cell(row, col).Clear(XLClearOptions.Contents);
                    }
                }
            }
        }

        private static void ApplyFinalOirStyling(IXLWorksheet worksheet, int headerRow)
        {
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            string[] accentHeaders =
            {
        "AR Status",
        "Name/WE",
        "Payment Number",
        "Update(pick from drop down)",
        "WOW Number",
        "CWS Number",
        "Client System Timesheet/Invoice ID",
        "Details",
        "Action(steps that still need to be taken)",
        "Tokyo Timesheet ID"
    };
            string[] lighterHeaders =
            {
        "Consultant",
        "Service End",
        "Document Amount",
        "Remaining Amount",
        "Percent Open",
        "Document Number",
        "Aging Bucket"
    };

            int agingBucketColumn = FindColumn(worksheet, headerRow, "Aging Bucket");

            if (agingBucketColumn != -1)
            {
                int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? agingBucketColumn;

                for (int column = agingBucketColumn + 1; column <= lastColumn; column++)
                {
                    var range = worksheet.Range(headerRow, column, lastRow, column);

                    range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.RightBorder = XLBorderStyleValues.Thin;

                    worksheet.Cell(headerRow, column).Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#FDE9D9");

                    worksheet.Cell(headerRow, column).Style.Font.Bold = false;

                    string header = worksheet.Cell(headerRow, column).GetString().Trim();

                    if (header.Equals("AR Status", StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet.Column(column).Width = 23;
                    }
                }
            }
            foreach (string header in lighterHeaders)
            {
                int column = FindColumn(worksheet, headerRow, header);

                if (column == -1)
                    continue;

                var range = worksheet.Range(headerRow, column, lastRow, column);

                range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                range.Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(headerRow, column).Style.Fill.BackgroundColor = XLColor.FromHtml("#F8CBAD");
                worksheet.Cell(headerRow, column).Style.Font.Bold = false;
            }
        }

        private static void ClearArea(int startLine, int numberOfLines)
        {
            for (int i = 0; i < numberOfLines; i++)
            {
                int line = startLine + i;
                if (line < 0 || line >= Console.BufferHeight) continue;

                Console.SetCursorPosition(0, line);
                Console.Write(new string(' ', Console.WindowWidth - 1));
            }

            Console.SetCursorPosition(0, startLine);
        }

        private record ColumnSnapshot(string HeaderName, List<string> Values);

    }
}