using ClosedXML.Excel;

namespace KellyCashApp
{
    public class UAC
    {
        public static void RunFullCashReportPullForward(int promptTop)
        {
            string? newUacPath = PromptForFilePath(
                "Paste the full file path of your NEW Full Cash Report:",
                promptTop
            );

            if (string.IsNullOrWhiteSpace(newUacPath) || !File.Exists(newUacPath))
            {
                Console.WriteLine("File not found. Please check the NEW Full Cash Report path.");
                Console.WriteLine("Press any key to return...");
                Console.ReadKey(true);
                return;
            }

            string? oldUacPath = PromptForFilePath(
                "Paste the full file path of your OLD / notated Full Cash Report:",
                promptTop
            );

            if (string.IsNullOrWhiteSpace(oldUacPath) || !File.Exists(oldUacPath))
            {
                Console.WriteLine("File not found. Please check the OLD Full Cash Report path.");
                Console.WriteLine("Press any key to return...");
                Console.ReadKey(true);
                return;
            }

            try
            {
                bool loading = true;

                Task spinner = Task.Run(() =>
                {
                    char[] frames = { '/', '-', '\\', '|' };
                    int i = 0;

                    while (loading)
                    {
                        Console.SetCursorPosition(0, promptTop);
                        Console.Write($"Preparing Full Cash Report... {frames[i++ % frames.Length]}   ");
                        Thread.Sleep(120);
                    }
                });

                List<string> newPayments;
                string outputPath = ProcessFullCashReport(newUacPath, oldUacPath, out newPayments);

                loading = false;
                spinner.Wait();

                ClearArea(promptTop, 12);
                Console.SetCursorPosition(0, promptTop);

                Console.WriteLine("Full Cash Report prepared successfully.");
                Console.WriteLine($"Saved to: {outputPath}");
                Console.WriteLine();
                ShowNewPaymentsLog(newPayments);
                Console.WriteLine();
                Console.WriteLine("Press any key to return to menu...");
                Console.ReadKey(true);
            }
            catch (Exception ex)
            {
                ClearArea(promptTop, 10);
                Console.SetCursorPosition(0, promptTop);

                Console.WriteLine("Something went wrong while preparing the Full Cash Report:");
                Console.WriteLine(ex.Message);
                Console.WriteLine();
                Console.WriteLine("Press any key to return...");
                Console.ReadKey(true);
            }
        }

        private static string ProcessFullCashReport(string newUacPath, string oldUacPath, out List<string> newPayments)
        {
            using var stream = new FileStream(newUacPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.First();

            int headerRow = FindHeaderRow(worksheet);

            if (headerRow == -1)
                throw new Exception("Could not determine the header row in the NEW Full Cash Report.");

            AddAmountOpenColumn(worksheet, headerRow);

            string[] missingColumns = GetMissingOldHeaders(oldUacPath, worksheet, headerRow);

            AddColumnsToEnd(worksheet, headerRow, missingColumns);
            newPayments = PullForwardOldUacNotes(worksheet, headerRow, oldUacPath, missingColumns);

            ApplyFinalUacStyling(worksheet, headerRow, missingColumns);

            worksheet.Columns().AdjustToContents();

            string downloadsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads"
            );

            string today = DateTime.Now.ToString("M.d.yyyy");
            string outputPath = GetUniqueOutputPath(downloadsPath, $"UAC {today}.xlsx");

            workbook.SaveAs(outputPath);

            return outputPath;
        }

        // Column normalization 
        private static string[] GetMissingOldHeaders(string oldUacPath, IXLWorksheet newWorksheet, int newHeaderRow)
        {
            using var stream = new FileStream(oldUacPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var workbook = new XLWorkbook(stream);
            var oldWorksheet = workbook.Worksheets.First();

            int oldHeaderRow = FindHeaderRow(oldWorksheet);

            if (oldHeaderRow == -1)
                throw new Exception("Could not determine the header row in the OLD Full Cash Report.");

            int oldLastColumn = oldWorksheet.LastColumnUsed()?.ColumnNumber() ?? 1;

            var missingHeaders = new List<string>();

            for (int col = 1; col <= oldLastColumn; col++)
            {
                string oldHeader = oldWorksheet.Cell(oldHeaderRow, col).GetString().Trim();

                if (string.IsNullOrWhiteSpace(oldHeader))
                    continue;

                int newColumn = FindColumn(newWorksheet, newHeaderRow, oldHeader);

                if (newColumn == -1 && !missingHeaders.Contains(oldHeader, StringComparer.OrdinalIgnoreCase))
                {
                    missingHeaders.Add(oldHeader);
                }
            }

            return missingHeaders.ToArray();
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

        private static void ApplyFinalUacStyling(IXLWorksheet worksheet, int headerRow, string[] addedColumns)
        {
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;
            int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;

            var usedRange = worksheet.Range(headerRow, 1, lastRow, lastColumn);

            usedRange.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            usedRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            usedRange.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            usedRange.Style.Border.RightBorder = XLBorderStyleValues.Thin;

            foreach (string header in addedColumns)
            {
                int column = FindColumn(worksheet, headerRow, header);

                if (column == -1)
                    continue;

                worksheet.Cell(headerRow, column).Style.Fill.BackgroundColor = XLColor.FromHtml("#FDE9D9");
                worksheet.Cell(headerRow, column).Style.Font.Bold = false;
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

        private static int FindHeaderRow(IXLWorksheet worksheet)
        {
            int lastRow = Math.Min(worksheet.LastRowUsed()?.RowNumber() ?? 20, 20);
            int bestRow = -1;
            int bestCount = 0;

            for (int row = 1; row <= lastRow; row++)
            {
                int usedCells = worksheet.Row(row).CellsUsed().Count();

                if (usedCells > bestCount)
                {
                    bestCount = usedCells;
                    bestRow = row;
                }
            }

            return bestCount > 1 ? bestRow : -1;
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

        private static void AddAmountOpenColumn(IXLWorksheet worksheet, int headerRow)
        {
            int originalTrxAmountColumn = FindColumn(worksheet, headerRow, "Original Trx Amount");
            int currentTrxAmountColumn = FindColumn(worksheet, headerRow, "Current Trx Amount");

            if (originalTrxAmountColumn == -1)
                throw new Exception("Could not find Original Trx Amount column.");

            if (currentTrxAmountColumn == -1)
                throw new Exception("Could not find Current Trx Amount column.");

            int existingAmountOpenColumn = FindColumn(worksheet, headerRow, "Amount Open");

            if (existingAmountOpenColumn != -1)
                worksheet.Column(existingAmountOpenColumn).Delete();

            currentTrxAmountColumn = FindColumn(worksheet, headerRow, "Current Trx Amount");
            originalTrxAmountColumn = FindColumn(worksheet, headerRow, "Original Trx Amount");

            worksheet.Column(currentTrxAmountColumn).InsertColumnsAfter(1);

            int amountOpenColumn = currentTrxAmountColumn + 1;

            worksheet.Cell(headerRow, amountOpenColumn).Value = "Amount Open";

            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            string originalLetter = worksheet.Column(originalTrxAmountColumn).ColumnLetter();
            string currentLetter = worksheet.Column(currentTrxAmountColumn).ColumnLetter();

            worksheet.Column(amountOpenColumn).Style.NumberFormat.Format = "0.00%";

            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                worksheet.Cell(row, amountOpenColumn).FormulaA1 =
                    $"=IF({originalLetter}{row}=0,0,{currentLetter}{row}/{originalLetter}{row})";
            }
        }

        private static List<string> PullForwardOldUacNotes(
    IXLWorksheet newWorksheet,
    int newHeaderRow,
    string oldUacPath,
    string[] noteColumnsToPull
)
        {
            using var oldStream = new FileStream(oldUacPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var oldWorkbook = new XLWorkbook(oldStream);
            var oldWorksheet = oldWorkbook.Worksheets.First();

            int oldHeaderRow = FindHeaderRow(oldWorksheet);

            int oldDocumentNumberColumn = FindColumn(oldWorksheet, oldHeaderRow, "Document Number");
            int newDocumentNumberColumn = FindColumn(newWorksheet, newHeaderRow, "Document Number");

            if (oldDocumentNumberColumn == -1)
                throw new Exception("Could not find Document Number column in OLD Full Cash Report.");

            if (newDocumentNumberColumn == -1)
                throw new Exception("Could not find Document Number column in NEW Full Cash Report.");

            int oldLastRow = oldWorksheet.LastRowUsed()?.RowNumber() ?? oldHeaderRow;
            int newLastRow = newWorksheet.LastRowUsed()?.RowNumber() ?? newHeaderRow;

            var oldRowsByDocumentNumber = new Dictionary<string, int>();

            for (int oldRow = oldHeaderRow + 1; oldRow <= oldLastRow; oldRow++)
            {
                string oldDocumentNumber = NormalizeDocumentNumber(
                    oldWorksheet.Cell(oldRow, oldDocumentNumberColumn).Value.ToString()
                );

                if (!string.IsNullOrWhiteSpace(oldDocumentNumber) &&
                    !oldRowsByDocumentNumber.ContainsKey(oldDocumentNumber))
                {
                    oldRowsByDocumentNumber.Add(oldDocumentNumber, oldRow);
                }
            }

            var newPayments = new List<string>();

            for (int newRow = newHeaderRow + 1; newRow <= newLastRow; newRow++)
            {
                string newDocumentNumber = NormalizeDocumentNumber(
                    newWorksheet.Cell(newRow, newDocumentNumberColumn).Value.ToString()
                );

                if (string.IsNullOrWhiteSpace(newDocumentNumber))
                    continue;

                if (!oldRowsByDocumentNumber.TryGetValue(newDocumentNumber, out int oldRow))
                {
                    newPayments.Add(newDocumentNumber);
                    continue;
                }

                foreach (string noteColumnName in noteColumnsToPull)
                {
                    int oldNoteColumn = FindColumn(oldWorksheet, oldHeaderRow, noteColumnName);
                    int newNoteColumn = FindColumn(newWorksheet, newHeaderRow, noteColumnName);

                    if (oldNoteColumn == -1 || newNoteColumn == -1)
                        continue;

                    newWorksheet.Cell(newRow, newNoteColumn).Value =
                        oldWorksheet.Cell(oldRow, oldNoteColumn).Value;
                }
            }

            return newPayments.OrderBy(x => x).ToList();
        }

        private static string NormalizeDocumentNumber(string value)
        {
            value = value.Trim();

            if (double.TryParse(value, out double number))
                return number.ToString("0");

            return value;
        }

        private static void ShowNewPaymentsLog(List<string> newPayments)
        {
            if (newPayments.Count == 0)
            {
                Console.WriteLine("No new payments found on the new UAC.");
                return;
            }

            Console.WriteLine($"New payments on new UAC: {newPayments.Count}");
            Console.WriteLine("Not found on old UAC:");
            Console.WriteLine();

            int maxToShow = Math.Min(newPayments.Count, 15);

            for (int i = 0; i < maxToShow; i++)
            {
                Console.WriteLine($"  - {newPayments[i]}");
            }

            if (newPayments.Count > maxToShow)
            {
                Console.WriteLine($"  ...and {newPayments.Count - maxToShow} more");
            }
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
    }
}