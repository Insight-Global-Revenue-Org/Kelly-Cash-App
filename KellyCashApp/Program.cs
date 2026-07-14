using ClosedXML.Excel;
using KellyCashApp.Configuration;
using KellyCashApp.Models;
using KellyCashApp.Processors;
using KellyCashApp.Processors.Kelly_Services;
using KellyCashApp.Processors.Monument;
using KellyCashApp.Processors.Allegis;
using KellyCashApp.Processors.Randstad;
using KellyCashApp.Processors.Guidant;
using KellyCashApp.Services;
using KellyCashApp.Workflows;
using System.Globalization;

Console.ForegroundColor = ConsoleColor.White;

int fixedMenuTop = ConsoleUi.DrawHeader();

var openInvoiceMatches = new Dictionary<string, OirMatch>();
var openInvoiceMatchesMultiple = new Dictionary<string, List<OirMatch>>();
var openInvoiceMatchesByClientProject = new Dictionary<string, List<OirMatch>>();
string? importedOirStatusMessage = null;
Dictionary<string, List<MicrosoftVmsMatch>>? microsoftVmsMatches = null;
bool skipMicrosoftVmsPrompt = false;

string? inputPath = null;
int defaultMenuOption = 0;

// Main application workflow loop.
// Menu navigation, OIR imports, remittance processing (Any additional process workflows will be delegated to other C# classes)
// test
while (true)
{
    int selected = ShowMenu(new[]
{
    "Import Open Invoice Report",
    "Process Remittance Payment",
    "Pull Forward OIR/UAC",
    "Settings",
    "Analytics Logs",
    "Exit"
}, defaultMenuOption, fixedMenuTop, importedOirStatusMessage);

    if (selected == 5)
        break;

    int promptTop = fixedMenuTop + 6;

    // Open Invoice Report (OIR) import workflow.
    // After importing, saves invoice/document amount mappings into system memory
    if (selected == 0)
    {
        string? oirPath = FileSelector.SelectFile(
        "Paste the full file path of the Open Invoice Report:",
        promptTop
            );

        if (string.IsNullOrWhiteSpace(oirPath) || !File.Exists(oirPath))
        {
            ClearArea(promptTop, 4);
            Console.SetCursorPosition(0, promptTop);
            Console.WriteLine("File not found. Please check the path and try again.\n");
            continue;
        }

        bool oirLoading = true;

        // Show the processing message immediately on the main thread.
        Console.SetCursorPosition(0, promptTop);
        Console.Write("Importing Open Invoice Report... /   ");

        Task oirSpinner = Task.Run(() =>
        {
            char[] frames = { '-', '\\', '|', '/' };
            int i = 0;

            while (oirLoading)
            {
                Thread.Sleep(120);

                Console.SetCursorPosition(0, promptTop);
                Console.Write(
                    $"Importing Open Invoice Report... {frames[i++ % frames.Length]}   ");
            }
        });

        try
        {
            openInvoiceMatches = OirImporter.Load(oirPath);
            openInvoiceMatchesMultiple = OirImporter.LoadMultiple(oirPath);
            openInvoiceMatchesByClientProject = OirImporter.LoadByClientProject(oirPath);
        }
        catch (Exception ex)
        {
            oirLoading = false;
            oirSpinner.Wait();

            ClearArea(promptTop, 8);
            Console.SetCursorPosition(0, promptTop);

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Something went wrong while importing the OIR:");
            Console.ResetColor();

            Console.WriteLine(ex.Message);
            Console.WriteLine();
            Console.WriteLine("Press any key to return to the menu...");
            Console.ReadKey(true);

            defaultMenuOption = 0;
            continue;
        }
        finally
        {
            oirLoading = false;
            oirSpinner.Wait();

            ClearArea(promptTop, 12);
        }

        ClearArea(promptTop, 12);
        Console.SetCursorPosition(0, promptTop);
        Console.WriteLine($"Imported {openInvoiceMatches.Count} OIR matches into memory.");
        Console.WriteLine();
        Console.WriteLine("Press 'Enter' to process your payment");

        Thread.Sleep(1000); 
        defaultMenuOption = 1;
        continue;
    }

    if (selected == 2)
    {
        ConsoleUi.ResetPage(0);
        fixedMenuTop = ConsoleUi.DrawHeader();

        OIR.ShowNotesMenu(fixedMenuTop);

        ConsoleUi.ResetPage(0);
        fixedMenuTop = ConsoleUi.DrawHeader();

        defaultMenuOption = 2;
        continue;
    }
    if (selected == 3)
    {
        ConsoleUi.ResetPage(0);
        fixedMenuTop = ConsoleUi.DrawHeader();

        Settings.ShowSettingsMenu(fixedMenuTop);

        ConsoleUi.ResetPage(0);
        fixedMenuTop = ConsoleUi.DrawHeader();

        defaultMenuOption = 3;
        continue;
    }
    if (selected == 4)
    {
        ClearArea(fixedMenuTop, 20);
        Analytics.ShowAnalyticsMenu(fixedMenuTop);

        ClearArea(fixedMenuTop, 20);
        defaultMenuOption = 4;
        continue;
    }

    // Remittance processing workflow.
    // Normalize columns, calculate aggregates,
    // auto-match invoices with documented amount, and generates the formatted output file.
    inputPath = FileSelector.SelectFile(
        "Paste the full file path of the Remittance Payment:",
        promptTop
            );

    if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
    {
        ClearArea(promptTop, 4);
        Console.SetCursorPosition(0, promptTop);
        Console.WriteLine("File not found. Please check the path and try again.\n");
        continue;
    }

    bool loading = true;

    // Show the processing message immediately.
    Console.SetCursorPosition(0, promptTop);
    Console.Write("Processing payment file... /   ");

    Task spinner = Task.Run(() =>
    {
        char[] frames = { '-', '\\', '|', '/' };
        int i = 0;

        while (loading)
        {
            Thread.Sleep(120);

            Console.SetCursorPosition(0, promptTop);
            Console.Write(
                $"Processing payment file... {frames[i++ % frames.Length]}   ");
        }
    });

    try
    {
        if (RandstadPayment.IsRandstadFormat(inputPath))
        {
            string randstadOutputPath = RandstadPayment.Process(
                inputPath,
                openInvoiceMatches,
                openInvoiceMatchesByClientProject
            );

            loading = false;
            spinner.Wait();

            ClearArea(promptTop, 8);
            Console.SetCursorPosition(0, promptTop);

            Console.WriteLine("Randstad payment processed successfully.");
            Console.WriteLine($"Updated file saved to: {randstadOutputPath}");
            Console.WriteLine();
            Console.WriteLine("Press any key to return to the menu...");
            Console.ReadKey(true);

            defaultMenuOption = 1;
            continue;
        }

        using var stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

        // Open remittance workbook using the installed ClosedXML dependency
        // FileShare.ReadWrite allows the system to process while the file is open elsewhere.
        using var workbook = new XLWorkbook(stream);

        IXLWorksheet worksheet;

        if (workbook.TryGetWorksheet("Payment Details", out var paymentDetailsSheet))
        {
            worksheet = paymentDetailsSheet;
        }
        else
        {
            worksheet = workbook.Worksheet(1);
        }
        // -------- Main Conditional loop for all Allegis Payments! --------
        // Conditional check for Microsoft payments (Re-Routing)
        if (MicrosoftPayment.IsMicrosoftFormat(worksheet))
        {
            loading = false;
            spinner.Wait();

            ClearArea(promptTop, 8);
            Console.SetCursorPosition(0, promptTop);

            if (microsoftVmsMatches == null && !skipMicrosoftVmsPrompt)
            {
                Console.WriteLine("Import VMS Timesheet Report? (Yes/No)");
                Console.Write("> ");

                string answer = Console.ReadLine()?.Trim() ?? "";

                if (answer.Equals("No", StringComparison.OrdinalIgnoreCase) ||
                    answer.Equals("N", StringComparison.OrdinalIgnoreCase))
                {
                    skipMicrosoftVmsPrompt = true;
                }

                if (answer.Equals("Yes", StringComparison.OrdinalIgnoreCase) ||
                    answer.Equals("Y", StringComparison.OrdinalIgnoreCase))
                {
                    string? vmsPath = FileSelector.SelectFile(
                        "Paste the full file path of the Microsoft VMS Timesheet Report:",
                        promptTop
                    );

                    if (string.IsNullOrWhiteSpace(vmsPath) || !File.Exists(vmsPath))
                    {
                        ClearArea(promptTop, 6);
                        Console.SetCursorPosition(0, promptTop);
                        Console.WriteLine("VMS file not found. Microsoft payment was not processed.");
                        Console.WriteLine("Press any key to return to the menu...");
                        Console.ReadKey(true);

                        defaultMenuOption = 1;
                        continue;
                    }

                    bool vmsLoading = true;

                    Task vmsSpinner = Task.Run(() =>
                    {
                        char[] frames = { '/', '-', '\\', '|' };
                        int i = 0;

                        while (vmsLoading)
                        {
                            Console.SetCursorPosition(0, promptTop);
                            Console.Write($"Importing VMS Report... {frames[i++ % frames.Length]}   ");
                            Thread.Sleep(120);
                        }
                    });

                    try
                    {
                        microsoftVmsMatches = MicrosoftVms.Import(vmsPath);
                    }
                    finally
                    {
                        vmsLoading = false;
                        vmsSpinner.Wait();

                        ClearArea(promptTop, 8);
                    }
                }
            }

            loading = true;

            spinner = Task.Run(() =>
            {
                char[] frames = { '/', '-', '\\', '|' };
                int i = 0;

                while (loading)
                {
                    Console.SetCursorPosition(0, promptTop);
                    Console.Write($"Processing Microsoft payment file... {frames[i++ % frames.Length]}   ");
                    Thread.Sleep(120);
                }
            });

            string microsoftOutputPath = MicrosoftPayment.Process(
                workbook,
                worksheet,
                inputPath,
                openInvoiceMatchesMultiple,
                microsoftVmsMatches
            );

            loading = false;
            spinner.Wait();

            ClearArea(promptTop, 8);
            Console.SetCursorPosition(0, promptTop);

            Console.WriteLine(microsoftVmsMatches == null
                ? "Microsoft payment processed successfully."
                : "Microsoft payment processed successfully with VMS report.");

            Console.WriteLine($"Updated file saved to: {microsoftOutputPath}");
            Console.WriteLine();
            Console.WriteLine("Press any key to return to the menu...");
            Console.ReadKey(true);

            defaultMenuOption = 1;
            continue;
        }

        // Conditional check for CDW payments (Re-Routing)
        if (CDWPayment.IsCDWFormat(worksheet))
        {
            string cdwOutputPath = CDWPayment.Process(
                workbook,
                worksheet,
                inputPath,
                openInvoiceMatchesMultiple
            );

            loading = false;
            spinner.Wait();

            ClearArea(promptTop, 8);
            Console.SetCursorPosition(0, promptTop);

            Console.WriteLine("CDW payment processed successfully.");
            Console.WriteLine($"Updated file saved to: {cdwOutputPath}");
            Console.WriteLine();
            Console.WriteLine("Press any key to return to the menu...");
            Console.ReadKey(true);

            defaultMenuOption = 1;
            continue;
        }

        // Conditional check for Samsung payments (Re-Routing)
        if (SamsungPayment.IsSamsungFormat(worksheet))
        {
            string samsungOutputPath = SamsungPayment.Process(
                workbook,
                worksheet,
                inputPath,
                openInvoiceMatchesMultiple
            );

            loading = false;
            spinner.Wait();

            ClearArea(promptTop, 8);
            Console.SetCursorPosition(0, promptTop);

            Console.WriteLine("Samsung payment processed successfully.");
            Console.WriteLine($"Updated file saved to: {samsungOutputPath}");
            Console.WriteLine();
            Console.WriteLine("Press any key to return to the menu...");
            Console.ReadKey(true);

            defaultMenuOption = 1;
            continue;
        }

        // Conditional check for Cushman & Wakefield payments (Re-Routing)
        if (CushmanWakefieldPayment.IsCushmanWakefieldFormat(worksheet))
        {
            string cushmanWakefieldOutputPath = CushmanWakefieldPayment.Process(
                workbook,
                worksheet,
                inputPath,
                openInvoiceMatchesMultiple
            );

            loading = false;
            spinner.Wait();

            ClearArea(promptTop, 8);
            Console.SetCursorPosition(0, promptTop);

            Console.WriteLine("Cushman Wakefield payment processed successfully.");
            Console.WriteLine($"Updated file saved to: {cushmanWakefieldOutputPath}");
            Console.WriteLine();
            Console.WriteLine("Press any key to return to the menu...");
            Console.ReadKey(true);

            defaultMenuOption = 1;
            continue;
        }

        // Conditional check for Johnson & Johnson payments (Re-Routing)
        if (JohnsonJohnsonPayment.IsJohnsonJohnsonFormat(worksheet))
        {
            string jnjOutputPath = JohnsonJohnsonPayment.Process(workbook, worksheet, inputPath);

            loading = false;
            spinner.Wait();

            ClearArea(promptTop, 8);
            Console.SetCursorPosition(0, promptTop);

            Console.WriteLine("Johnson & Johnson fixed-fee payment processed successfully.");
            Console.WriteLine($"Updated file saved to: {jnjOutputPath}");
            Console.WriteLine();
            Console.WriteLine("Press any key to return to the menu...");
            Console.ReadKey(true);

            defaultMenuOption = 1;
            continue;
        }

        // Conditional check for Guidant payments (Re-Routing)
        if (GuidantPayment.IsGuidantFormat(worksheet))
        {
            string guidantOutputPath = GuidantPayment.Process(
            workbook,
            worksheet,
            inputPath,
            openInvoiceMatches
        );
            // This better fix it
            workbook.Dispose();
            stream.Dispose();

            loading = false;
            spinner.Wait();

            ClearArea(promptTop, 8);
            Console.SetCursorPosition(0, promptTop);

            Console.WriteLine("Guidant payment processed successfully.");
            Console.WriteLine($"Updated file saved to: {guidantOutputPath}");
            Console.WriteLine();
            Console.WriteLine("Press any key to return to the menu...");
            Console.ReadKey(true);

            defaultMenuOption = 1;
            continue;
        }

        // Conditional check for Monument payments (Re-Routing)
        if (MonumentPayment.IsMonumentFormat(worksheet))
        {
            string monumentOutputPath = MonumentPayment.Process(workbook, worksheet, inputPath, openInvoiceMatches);

            loading = false;
            spinner.Wait();

            ClearArea(promptTop, 8);
            Console.SetCursorPosition(0, promptTop);

            Console.WriteLine("Monument payment processed successfully.");
            Console.WriteLine($"Updated file saved to: {monumentOutputPath}");
            Console.WriteLine();
            Console.WriteLine("Press any key to return to the menu...");
            Console.ReadKey(true);

            defaultMenuOption = 1;
            continue;
        }

        // No specialized format matched.
        // Process using standard Kelly logic.
        string kellyOutputPath = KellyPayment.Process(
            workbook,
            worksheet,
            inputPath,
            openInvoiceMatches);

        loading = false;
        spinner.Wait();

        ClearArea(promptTop, 8);
        Console.SetCursorPosition(0, promptTop);

        Console.WriteLine(
            "Kelly payment processed successfully.");

        Console.WriteLine(
            $"Updated file saved to: {kellyOutputPath}");

        Console.WriteLine();
        Console.WriteLine(
            "Press 'Enter' to process another payment.");

        Thread.Sleep(1000);

        defaultMenuOption = 1;
        continue;
    }
    catch (Exception ex)
    {
        loading = false;
        spinner.Wait();
        Console.WriteLine();

        Console.WriteLine("Something went wrong:");
        Console.WriteLine(ex.Message);
    }

}

Console.WriteLine();
Console.WriteLine("Press any key to exit...");
Console.ReadKey(true);

static int FindColumn(IXLWorksheet worksheet, int headerRow, string headerName)
{
    for (int col = 1; col <= 100; col++)
    {
        string headerText = worksheet.Cell(headerRow, col).Value.ToString().Trim();

        if (headerText.Equals(headerName, StringComparison.OrdinalIgnoreCase))
            return col;
    }

    return -1;
}

static void NormalizeColumns(IXLWorksheet worksheet, int headerRow, int contractorColumn)
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
        string currentHeader = worksheet.Cell(headerRow, insertAt).Value.ToString().Trim();

        if (!currentHeader.Equals(columnName, StringComparison.OrdinalIgnoreCase))
        {
            InsertBlankColumn(worksheet, insertAt, columnName, headerRow);
        }

        insertAt++;
    }
}

static void InsertBlankColumn(IXLWorksheet worksheet, int targetColumn, string headerName, int headerRow)
{
    int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;
    int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? targetColumn;

    for (int row = 1; row <= lastRow; row++)
    {
        for (int col = lastColumn; col >= targetColumn; col--)
        {
            worksheet.Cell(row, col + 1).CopyFrom(worksheet.Cell(row, col));
            worksheet.Cell(row, col).Clear();
        }
    }

    worksheet.Cell(headerRow, targetColumn).Value = headerName;
    worksheet.Cell(headerRow, targetColumn).Style =
        worksheet.Cell(headerRow, targetColumn - 1).Style;

    var designCell = worksheet.Cell(headerRow - 1, targetColumn);
    var leftCell = worksheet.Cell(headerRow - 1, targetColumn - 1);

    designCell.Style = leftCell.Style;

    leftCell.Style.Border.RightBorder = XLBorderStyleValues.None;
}

static int FindColumnAfter(IXLWorksheet worksheet, int headerRow, string headerName, int afterColumn)
{
    for (int col = afterColumn + 1; col <= 100; col++)
    {
        string headerText = worksheet.Cell(headerRow, col).Value.ToString().Trim();

        if (headerText.Equals(headerName, StringComparison.OrdinalIgnoreCase))
            return col;
    }

    return -1;
}

static decimal GetDecimalValue(IXLCell cell)
{
    string rawValue = cell.Value.ToString()
        .Replace("$", "")
        .Replace(",", "")
        .Trim();

    if (decimal.TryParse(rawValue, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
        return result;

    return 0;
}

static string MakeSafeFileName(string fileName)
{
    foreach (char invalidChar in Path.GetInvalidFileNameChars())
    {
        fileName = fileName.Replace(invalidChar, '_');
    }

    return fileName.Trim();
}

static bool TryMatchWithDateSpread(
    string formattedName,
    string formattedWeekEnding,
    Dictionary<string, OirMatch> openInvoiceMatches,
    out OirMatch match)
{
    match = null;

    if (!DateTime.TryParse(formattedWeekEnding, out DateTime baseDate))
        return false;

    // Try exact, then ±1, ±2 days
    for (int offset = -2; offset <= 2; offset++)
    {
        DateTime testDate = baseDate.AddDays(offset);
        string testDateString = testDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

        string testKey = $"{formattedName} {testDateString}".Trim();

        if (openInvoiceMatches.TryGetValue(testKey, out match))
        {
            return true;
        }
    }

    return false;
}

static string GetUniqueOutputPath(string folderPath, string fileName)
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

static string FormatName(string input)
{
    // Expected Normalized Format: "DOE, JOHN"
    if (!input.Contains(","))
        return input;

    var parts = input.Split(',');

    if (parts.Length < 2)
        return input;

    string last = parts[0].Trim().ToLower();
    string first = parts[1].Trim().ToLower();

    last = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(last);
    first = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(first);

    return $"{first} {last}";
}

static string FormatWeekEndingDate(IXLCell cell)
{
    if (cell.Value.IsDateTime)
    {
        return cell.GetDateTime().ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
    }

    string rawValue = cell.GetString().Trim();

    if (DateTime.TryParse(rawValue, out DateTime parsedDate))
    {
        return parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
    }

    return rawValue;
}

// The main console gui menu rendering function. Displays list of options, highlights the currently selected option,
// and allows the user to navigate with the up/down arrow keys and select with Enter. Any additional classes wired into the main Program.cs will also be updated here.
static int ShowMenu(string[] options, int defaultSelected, int menuTop, string? statusMessage = null)
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

    DrawStatusMessage(statusMessage);

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
            DrawStatusMessage(statusMessage);
        }
    }

    Console.ResetColor();
    Console.CursorVisible = true;

    int afterMenu = Math.Min(menuTop + neededRows, Console.BufferHeight - 1);
    Console.SetCursorPosition(0, afterMenu);
    Console.WriteLine();

    return selected;
}

static void DrawStatusMessage(string? message)
{
    if (string.IsNullOrWhiteSpace(message))
        return;

    int statusLine = Console.WindowHeight - 2;

    Console.SetCursorPosition(0, statusLine);
    Console.Write(new string(' ', Console.WindowWidth - 1));

    Console.SetCursorPosition(0, statusLine);
    Console.ForegroundColor = ConsoleColor.Green;

    Console.Write(message.Length >= Console.WindowWidth
        ? message[..(Console.WindowWidth - 1)]
        : message);

    Console.ResetColor();
}

static void DrawFullMenu(string[] options, int selected, int menuTop, int menuWidth)
{
    Console.ResetColor();
    Console.SetCursorPosition(0, menuTop);
    Console.WriteLine("Select".PadRight(Console.WindowWidth - 1));

    for (int i = 0; i < options.Length; i++)
    {
        Console.SetCursorPosition(0, menuTop + i + 1);

        // Clear the entire row first
        Console.Write(new string(' ', Console.WindowWidth - 1));

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

// Loads and parses the Open Invoice Report (OIR)
// into an in-memory lookup dictionary. Not persistent - if the user exits the application, they will need to re-import the OIR to have the invoice mappings available for remittance processing.

static string? PromptForFilePath(string prompt, int startLine)
{
    Console.CursorVisible = true;

    ClearArea(startLine, 6);

    Console.SetCursorPosition(0, startLine);
    Console.WriteLine(prompt);
    Console.Write("> ");

    int inputLine = Console.CursorTop;

    string? path = Console.ReadLine()?.Trim().Trim('"');

    ClearArea(startLine, 6);

    return path;
}

// HELPER FUNCTION: Console region cleanup, prevent some UI duplication artifacts I encountered
static void ClearArea(int startLine, int numberOfLines)
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
// May remove this helper later, not currently being used (needed for debug)
static void ClearLine(int line)
{
    if (line < 0 || line >= Console.BufferHeight) return;

    Console.SetCursorPosition(0, line);
    Console.Write(new string(' ', Console.WindowWidth - 1));
    Console.SetCursorPosition(0, line);
}
 
// Dynamically locate spreadsheet header rows
static int FindHeaderRow(IXLWorksheet worksheet, string requiredHeader)
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


