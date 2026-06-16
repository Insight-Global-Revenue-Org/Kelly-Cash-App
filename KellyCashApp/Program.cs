using ClosedXML.Excel;
using KellyCashApp.Configuration;
using KellyCashApp.Models;
using KellyCashApp.Processors;
using KellyCashApp.Processors.Monument;
using KellyCashApp.Processors.Allegis;
using KellyCashApp.Processors.Randstad;
using KellyCashApp.Services;
using KellyCashApp.Workflows;
using System.Globalization;

Console.ForegroundColor = ConsoleColor.White;



Console.ForegroundColor = ConsoleColor.White;

Console.WriteLine(@"
  _  __    _ _         ____                  _               
 | |/ /___| | |_   _  / ___|  ___ _ ____   _(_) ___ ___  ___ 
 | ' // _ \ | | | | | \___ \ / _ \ '__\ \ / / |/ __/ _ \/ __|
 | . \  __/ | | |_| |  ___) |  __/ |   \ V /| | (_|  __/\__ \
 |_|\_\___|_|_|\__, | |____/ \___|_|    \_/ |_|\___\___||___/
               |___/                                         
");

Console.ResetColor();

Thread.Sleep(100);

Console.WriteLine("A Week-Ending Line Total Aggregate Script");
Console.WriteLine("──────────────────────────────────────────────────");
var openInvoiceMatches = new Dictionary<string, OirMatch>();
var openInvoiceMatchesMultiple = new Dictionary<string, List<OirMatch>>();
Dictionary<string, List<MicrosoftVmsMatch>>? microsoftVmsMatches = null;
bool skipMicrosoftVmsPrompt = false;

string? inputPath = null;
int defaultMenuOption = 0;
int fixedMenuTop = Console.CursorTop;

// Main application workflow loop.
// Menu navigation, OIR imports, remittance processing (Any additional process workflows will be delegated to other C# classes)
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
}, defaultMenuOption, fixedMenuTop);

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

        Task oirSpinner = Task.Run(() =>
        {
            char[] frames = { '/', '-', '\\', '|' };
            int i = 0;

            while (oirLoading)
            {
                Console.SetCursorPosition(0, promptTop);
                Console.Write($"Importing Open Invoice Report... {frames[i++ % frames.Length]}   ");
                Thread.Sleep(120);
            }
        });

        try
        {
            openInvoiceMatches = OirImporter.Load(oirPath);
            openInvoiceMatchesMultiple = OirImporter.LoadMultiple(oirPath);
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
        ClearArea(fixedMenuTop, 10);
        OIR.ShowNotesMenu(fixedMenuTop);

        ClearArea(fixedMenuTop, 10);
        defaultMenuOption = 2;
        continue;
    }
    if (selected == 3)
    {
        ClearArea(fixedMenuTop, 10);
        Settings.ShowSettingsMenu(fixedMenuTop);

        ClearArea(fixedMenuTop, 10);
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

    Task spinner = Task.Run(() =>
    {
        char[] frames = { '/', '-', '\\', '|' };
        int i = 0;

        while (loading)
        {
            Console.SetCursorPosition(0, promptTop);
            Console.Write($"Processing payment file... {frames[i++ % frames.Length]}   ");
            Thread.Sleep(120);
        }
    });

    try
{
        if (RandstadPayment.IsRandstadFormat(inputPath))
        {
            string randstadOutputPath = RandstadPayment.Process(
                inputPath,
                openInvoiceMatches
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


        int headerRow = 13;

        // This is where we will dynamically locate the header columns
        int weekEndingColumn = FindColumn(worksheet, headerRow, "Week Ending Date");
    int contractorColumn = FindColumn(worksheet, headerRow, "Name");
    int lineTotalColumn = FindColumn(worksheet, headerRow, "Line Total");
    int locationDescriptionColumn = FindColumn(worksheet, headerRow, "Location Description");

    if (weekEndingColumn == -1 || contractorColumn == -1 || lineTotalColumn == -1 || locationDescriptionColumn == -1)
    {
            loading = false;
            spinner.Wait();

            Console.WriteLine("Could not find Week Ending Date, Name, Line Total, or Location Description.");
            Console.WriteLine("Press any key to return to the menu...");
            Console.ReadKey(true);

            defaultMenuOption = 1;
            continue;
        }
    // Normalize worksheet layout by ensuring required automation columns exist
    // will remain in a consistent position for the downstream processing.
    NormalizeColumns(worksheet, headerRow, contractorColumn);
    InsertBlankColumn(worksheet, contractorColumn + 1, "Name_", headerRow);

    // Re-find columns after normalization
    int invoiceColumn = FindColumnAfter(worksheet, headerRow, "Invoice", contractorColumn);
    int amountColumn = FindColumnAfter(worksheet, headerRow, "Amount", contractorColumn);
    int aggregateColumn = FindColumn(worksheet, headerRow, "Aggregate Line Total");
    int notesColumn = FindColumn(worksheet, headerRow, "Notes");

    weekEndingColumn = FindColumn(worksheet, headerRow, "Week Ending Date");
    contractorColumn = FindColumn(worksheet, headerRow, "Name");
    int nameFormattedColumn = FindColumnAfter(worksheet, headerRow, "Name_", contractorColumn);
    lineTotalColumn = FindColumn(worksheet, headerRow, "Line Total");

    // Add Concat column after Line Total
    InsertBlankColumn(worksheet, lineTotalColumn + 1, "Concat", headerRow);

    // Re-finding Line Total after inserting Concat
    lineTotalColumn = FindColumn(worksheet, headerRow, "Line Total");

    // And finding the new Concat column again for the downstream matching (previously was not added)
    int concatColumn = FindColumnAfter(worksheet, headerRow, "Concat", lineTotalColumn);

    if (amountColumn == -1 || aggregateColumn == -1 || weekEndingColumn == -1 || contractorColumn == -1 || lineTotalColumn == -1)
    {
        Console.WriteLine("Could not find one or more required columns.");
        Console.WriteLine($"Amount: {amountColumn}");
        Console.WriteLine($"Week Ending: {weekEndingColumn}");
        Console.WriteLine($"Contractor Name: {contractorColumn}");
            loading = false;
            spinner.Wait();

            Console.WriteLine($"Line Total: {lineTotalColumn}");
            Console.WriteLine("Press any key to return to the menu...");
            Console.ReadKey(true);

            defaultMenuOption = 1;
            continue;
        }

    int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

        var nameChanges = Rename.LoadNameChanges();

        for (int row = headerRow + 1; row <= lastRow; row++)
    {
        string rawName = worksheet.Cell(row, contractorColumn).GetString().Trim();

        if (string.IsNullOrWhiteSpace(rawName))
            continue;

            string formattedName = FormatName(rawName);
            formattedName = Rename.ApplyNameChange(formattedName, nameChanges);

            string formattedWeekEnding = FormatWeekEndingDate(worksheet.Cell(row, weekEndingColumn));

            // Generates composite lookup key used for our automatic OIR invoice matching.
            string concatValue = $"{formattedName} {formattedWeekEnding}".Trim();

        worksheet.Cell(row, nameFormattedColumn).Value = formattedName;
        worksheet.Cell(row, concatColumn).Value = concatValue;
     
    }

    loading = false;
    spinner.Wait();

    Console.Write("\r" + new string(' ', 60)); 
    Console.WriteLine(); 

    Console.WriteLine($"Found Amount column: {amountColumn}");
    Console.WriteLine($"Found Week Ending column: {weekEndingColumn}");
    Console.WriteLine($"Found Contractor Name column: {contractorColumn}");
    Console.WriteLine($"Found Line Total column: {lineTotalColumn}");
    Console.WriteLine($"Found Aggregate column: {aggregateColumn}");

    // First pass: calculate totals up to last row for each Contractor + Week Ending
    var totalsByContractorAndWeek = new Dictionary<string, decimal>();
    var lastRowByContractorAndWeek = new Dictionary<string, int>();

    for (int row = headerRow + 1; row < lastRow; row++)
    {
        string contractor = worksheet.Cell(row, contractorColumn).GetString().Trim();
        string weekEnding = worksheet.Cell(row, weekEndingColumn).GetString().Trim();

        if (string.IsNullOrWhiteSpace(contractor) || string.IsNullOrWhiteSpace(weekEnding))
            continue;

        decimal lineTotal = GetDecimalValue(worksheet.Cell(row, lineTotalColumn));
        string key = $"{contractor}|{weekEnding}";

        if (!totalsByContractorAndWeek.ContainsKey(key))
            totalsByContractorAndWeek[key] = 0;

        totalsByContractorAndWeek[key] += lineTotal;
        lastRowByContractorAndWeek[key] = row;
    }

    // Second pass: write aggregate total to the last row for each contractor
    foreach (var item in totalsByContractorAndWeek)
    {
        string key = item.Key;
        decimal aggregateTotal = item.Value;
        int targetRow = lastRowByContractorAndWeek[key];

        var cell = worksheet.Cell(targetRow, aggregateColumn);

        cell.Value = aggregateTotal;

        cell.Style = worksheet.Cell(targetRow, lineTotalColumn).Style;

        cell.Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";

        // Are credits applied? If so, make negative totals red so MSP can identify
        if (aggregateTotal < 0)
        {
            cell.Style.Font.FontColor = XLColor.Red;
        }

        string concatValue = worksheet.Cell(targetRow, concatColumn).GetString().Trim();

        // Auto-match our remittance entries against the users imported OIR invoice data.
        if (openInvoiceMatches.TryGetValue(concatValue, out OirMatch match))
        {
            worksheet.Cell(targetRow, invoiceColumn).Value = match.DocumentNumber;
            worksheet.Cell(targetRow, amountColumn).Value = match.RemainingAmount;

            worksheet.Cell(targetRow, amountColumn).Style =
                worksheet.Cell(targetRow, lineTotalColumn).Style;

            worksheet.Cell(targetRow, amountColumn).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
        }
    }

        string downloadsPath = Settings.GetRemittanceSavePath();

        string companyName = worksheet.Cell(headerRow + 1, locationDescriptionColumn).GetString().Trim();

    if (string.IsNullOrWhiteSpace(companyName))
    {
        companyName = "Remittance Payment";
    }

    companyName = MakeSafeFileName(companyName);

    decimal totalAmount = GetDecimalValue(worksheet.Cell(lastRow, lineTotalColumn));
    string formattedTotal = totalAmount.ToString("N2", CultureInfo.InvariantCulture);

    // Generate output filename using company name and payment total
    string outputPath = GetUniqueOutputPath(downloadsPath, $"{companyName} - {formattedTotal}.xlsx");

    int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? lineTotalColumn;

    if (worksheet.AutoFilter.IsEnabled)
    {
        worksheet.AutoFilter.Clear();
    }

    worksheet.Range(headerRow, 1, lastRow, lastColumn).SetAutoFilter();

    worksheet.Column(invoiceColumn).Width = 16;
    worksheet.Column(amountColumn).Width = 14;
    worksheet.Column(aggregateColumn).Width = 20;
    worksheet.Column(notesColumn).Width = 24;
    worksheet.Column(nameFormattedColumn).Width = 18;
    worksheet.Column(concatColumn).Width = 28;

    workbook.SaveAs(outputPath);

        Analytics.LogRemittanceRun($"{companyName} - {formattedTotal}");



        Console.WriteLine();
    Console.WriteLine("Week-Ending Line Totals Calculated Per Invoice Line Item");
    Console.WriteLine("All available open invoices have been automatically matched!");
    Console.WriteLine("\nPress 'Enter' to process another payment.");
    Console.WriteLine($"Updated file saved to: {outputPath}");

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
static int ShowMenu(string[] options, int defaultSelected, int menuTop)
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


