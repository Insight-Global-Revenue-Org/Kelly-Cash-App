namespace KellyCashApp
{
    internal class Analytics
    {
        private static string AnalyticsFolder =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "KellyCashApp",
                "analytics"
            );

        private static string OirStatsFile => Path.Combine(AnalyticsFolder, "oir-stats.txt");
        private static string UacStatsFile => Path.Combine(AnalyticsFolder, "uac-stats.txt");
        private static string RemittanceStatsFile => Path.Combine(AnalyticsFolder, "remittance-stats.txt");

        public static void LogOirRun(int droppedInvoiceCount)
        {
            Directory.CreateDirectory(AnalyticsFolder);

            string line = $"{DateTime.Now:MM/dd/yyyy hh:mm tt}|Dropped Invoices: {droppedInvoiceCount}";
            File.AppendAllLines(OirStatsFile, new[] { line });
        }

        public static void LogUacRun(int newPaymentCount)
        {
            Directory.CreateDirectory(AnalyticsFolder);

            string line = $"{DateTime.Now:MM/dd/yyyy hh:mm tt}|New Payments: {newPaymentCount}";
            File.AppendAllLines(UacStatsFile, new[] { line });
        }

        public static void LogRemittanceRun(string paymentName)
        {
            Directory.CreateDirectory(AnalyticsFolder);

            string line = $"{DateTime.Now:MM/dd/yyyy hh:mm tt}|Processed Payment: {paymentName}";
            File.AppendAllLines(RemittanceStatsFile, new[] { line });
        }

        public static void ShowAnalyticsMenu(int menuTop)
        {
            while (true)
            {
                int selected = ShowMenu(new[]
                {
                    "Open Invoice Report Stats",
                    "Full Cash Stats",
                    "Remittance Payment Stats",
                    "Back"
                }, 0, menuTop);

                if (selected == 3)
                    return;

                int promptTop = menuTop + 6;
                ClearArea(promptTop, 20);

                if (selected == 0)
                    ShowStats("Open Invoice Report Stats", OirStatsFile, "Dropped Invoices", promptTop);

                if (selected == 1)
                    ShowStats("Full Cash Stats", UacStatsFile, "New Payments", promptTop);

                if (selected == 2)
                    ShowStats("Remittance Payment Stats", RemittanceStatsFile, "Processed Payments", promptTop);

                ClearArea(promptTop, 22);
            }
        }

        private static void ShowStats(string title, string filePath, string totalLabel, int promptTop)
        {
            if (!File.Exists(filePath))
            {
                ClearArea(promptTop, 20);
                Console.SetCursorPosition(0, promptTop);

                Console.WriteLine(title);
                Console.WriteLine("────────────────────────────────────────────");
                Console.WriteLine();
                Console.WriteLine("No analytics data found yet.");
                Console.WriteLine();
                Console.WriteLine("Press any key to return...");
                Console.ReadKey(true);
                return;
            }

            string[] lines = File.ReadAllLines(filePath)
                .Reverse()
                .ToArray();

            int pageSize = 15;
            int page = 0;
            int totalPages = Math.Max(1, (int)Math.Ceiling(lines.Length / (double)pageSize));

            while (true)
            {
                ClearArea(promptTop, 28);
                Console.SetCursorPosition(0, promptTop);

                int totalCount;

                if (totalLabel.Equals("Processed Payments", StringComparison.OrdinalIgnoreCase))
                {
                    totalCount = lines.Length;
                }
                else
                {
                    totalCount = lines
                        .Select(line =>
                        {
                            int colonIndex = line.LastIndexOf(':');

                            if (colonIndex == -1)
                                return 0;

                            string numberText = line.Substring(colonIndex + 1).Trim();

                            return int.TryParse(numberText, out int value) ? value : 0;
                        })
                        .Sum();
                }

                Console.WriteLine(title);
                Console.WriteLine("────────────────────────────────────────────");
                Console.WriteLine();

                Console.WriteLine($"{totalLabel}: {totalCount}");

                if (title.Contains("Open Invoice", StringComparison.OrdinalIgnoreCase))
                    Console.WriteLine($"Total OIR Runs: {lines.Length}");

                if (title.Contains("Full Cash", StringComparison.OrdinalIgnoreCase))
                    Console.WriteLine($"Total Full Cash Runs: {lines.Length}");

                if (title.Contains("Remittance", StringComparison.OrdinalIgnoreCase))
                    Console.WriteLine($"Total Remittance Runs: {lines.Length}");

                Console.WriteLine();

                foreach (string line in lines.Skip(page * pageSize).Take(pageSize))
                {
                    string[] parts = line.Split('|', 2);

                    if (parts.Length == 2)
                        Console.WriteLine($"{parts[0]}  -  {parts[1]}");
                    else
                        Console.WriteLine(line);
                }

                int footerLine = Console.WindowHeight - 1;

                Console.SetCursorPosition(0, footerLine);
                Console.Write(new string(' ', Console.WindowWidth - 1));

                Console.SetCursorPosition(0, footerLine);
                Console.Write($"[PgUp] Newer  [PgDn] Older  [Esc] Back     Page {page + 1}/{totalPages}");

                ConsoleKey key = Console.ReadKey(true).Key;

                if (key == ConsoleKey.Escape)
                    return;

                if (key == ConsoleKey.PageDown && page < totalPages - 1)
                    page++;

                if (key == ConsoleKey.PageUp && page > 0)
                    page--;
            }
        }

        private static int ShowMenu(string[] options, int defaultSelected, int menuTop)
        {
            int selected = defaultSelected;
            int menuWidth = options.Max(o => o.Length) + 4;
            int neededRows = options.Length + 2;

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
                    DrawFullMenu(options, selected, menuTop, menuWidth);
            }

            Console.ResetColor();
            Console.CursorVisible = true;
            Console.SetCursorPosition(0, menuTop + neededRows);
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

                string line = i == selected ? $"> {options[i]}" : $"  {options[i]}";

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