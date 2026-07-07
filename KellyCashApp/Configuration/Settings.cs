namespace KellyCashApp.Configuration
{
    internal class Settings
    {
        private static string SettingsFolder =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "KellyCashApp"
            );

        private static string OirSavePathFile =>
            Path.Combine(SettingsFolder, "oir-save-path.txt");

        private static string UacSavePathFile =>
            Path.Combine(SettingsFolder, "uac-save-path.txt");

        private static string RemittanceSavePathFile =>
            Path.Combine(SettingsFolder, "remittance-save-path.txt");

        private static string NameChangeFilePathFile =>
            Path.Combine(SettingsFolder, "name-change-file-path.txt");

        private static string MicrosoftVmsReportFilePathFile =>
            Path.Combine(SettingsFolder, "microsoft-vms-report-file-path.txt");

        private static string NikeTrackerBoardFilePathFile =>
            Path.Combine(SettingsFolder, "nike-tracker-board-file-path.txt");

        private static string FileSelectionTypeFile =>
             Path.Combine(SettingsFolder, "file-selection-type.txt");

        public static void ShowSettingsMenu(int menuTop)
        {
            while (true)
            {
                ConsoleUi.ResetPage(menuTop);

                int selected = ShowMenu(new[]
            {
                "View Current Settings",
                "Open Invoice Report Save File Path",
                "Full Cash Save File Path",
                "Remittance Payment Save File Path",
                "Name Change File Path",
                "Additional Reporting Support",
                "File Selection Type",
                "Back"
            }, 0, menuTop);



                if (selected == 7)
                    return;

                int promptTop = menuTop;

                ConsoleUi.ResetPage(promptTop);
                Console.SetCursorPosition(0, promptTop);

                if (selected == 0)
                {
                    ConsoleUi.ResetPage(promptTop);
                    ShowCurrentSettings(menuTop);
                }

                if (selected == 1)
                {
                    SavePathSetting(
                        "Paste the folder path where OIR files should save:",
                        OirSavePathFile,
                        promptTop
                    );
                }

                if (selected == 2)
                {
                    SavePathSetting(
                        "Paste the folder path where Full Cash files should save:",
                        UacSavePathFile,
                        promptTop
                    );
                }

                if (selected == 3)
                {
                    SavePathSetting(
                        "Paste the folder path where Remittance Payment files should save:",
                        RemittanceSavePathFile,
                        promptTop
                    );
                }

                if (selected == 4)
                {
                    SaveFilePathSetting(
                        "Paste the full .txt or Excel file path for contractor name changes:",
                        NameChangeFilePathFile,
                        promptTop
                    );
                }

                if (selected == 5)
                {
                    ShowAdditionalReportingSupportMenu(promptTop);
                }

                if (selected == 6)
                {
                    SaveFileSelectionType(promptTop);
                }

                ConsoleUi.ResetPage(promptTop);
            }
        }

        public static string GetOpenInvoiceReportSavePath()
        {
            return GetSavedPath(OirSavePathFile);
        }

        public static string GetRemittanceSavePath()
        {
            return GetSavedPath(RemittanceSavePathFile);
        }

        public static string GetFullCashSavePath()
        {
            return GetSavedPath(UacSavePathFile);
        }

        public static string GetNameChangeFilePath()
        {
            if (File.Exists(NameChangeFilePathFile))
            {
                string savedPath = File.ReadAllText(NameChangeFilePathFile).Trim();

                if (File.Exists(savedPath))
                    return savedPath;
            }

            return "";
        }

        public static string GetMicrosoftVmsReportFilePath()
        {
            if (File.Exists(MicrosoftVmsReportFilePathFile))
            {
                string savedPath = File.ReadAllText(MicrosoftVmsReportFilePathFile).Trim();

                if (File.Exists(savedPath))
                    return savedPath;
            }

            return "";
        }

        public static string GetNikeTrackerBoardFilePath()
        {
            if (File.Exists(NikeTrackerBoardFilePathFile))
            {
                string savedPath = File.ReadAllText(NikeTrackerBoardFilePathFile).Trim();

                if (File.Exists(savedPath))
                    return savedPath;
            }

            return "";
        }

        private static void ShowAdditionalReportingSupportMenu(int menuTop)
        {
            while (true)
            {
                ConsoleUi.ResetPage(menuTop);

                int selected = ShowMenu(new[]
                {
            "Microsoft VMS Report File Path",
            "Nike Tracker Board File Path",
            "Back"
        }, 0, menuTop);

                if (selected == 2)
                    return;

                ConsoleUi.ResetPage(menuTop);

                if (selected == 0)
                {
                    SaveFilePathSetting(
                        "Paste the full Excel file path for the Microsoft VMS Report:",
                        MicrosoftVmsReportFilePathFile,
                        menuTop
                    );
                }

                if (selected == 1)
                {
                    SaveFilePathSetting(
                        "Paste the full Excel file path for the Nike Tracker Board:",
                        NikeTrackerBoardFilePathFile,
                        menuTop
                    );
                }

                ClearArea(menuTop, 12);
            }
        }

        public static string GetFileSelectionType()
        {
            if (File.Exists(FileSelectionTypeFile))
                return File.ReadAllText(FileSelectionTypeFile).Trim();

            return "Input File Path";
        }

        private static void SaveFileSelectionType(int promptTop)
        {
            int selected = ShowMenu(new[]
            {
        "Input File Path",
        "Input File Explorer",
        "Back"
    }, 0, promptTop);

            if (selected == 2)
                return;

            string value = selected == 0
                ? "Input File Path"
                : "Input File Explorer";

            Directory.CreateDirectory(SettingsFolder);
            File.WriteAllText(FileSelectionTypeFile, value);

            ClearArea(promptTop, 8);
            Console.SetCursorPosition(0, promptTop);
            Console.WriteLine($"File Selection Type saved as: {value}");
            Console.WriteLine();
            Console.WriteLine("Press any key to return...");
            Console.ReadKey(true);

            // Reset the entire screen before returning
            ConsoleUi.ResetPage(0);
            ConsoleUi.DrawHeader();
        }

        private static void ShowCurrentSettings(int promptTop)
        {
            ConsoleUi.ResetPage(0);

            int contentTop = ConsoleUi.DrawHeader() + 1;

            Console.SetCursorPosition(0, contentTop);

            Console.WriteLine("Current Settings");
            Console.WriteLine("────────────────────────────────────────────");
            Console.WriteLine();

            Console.WriteLine($"Open Invoice Report Save Path:");
            Console.WriteLine(GetOpenInvoiceReportSavePath());
            Console.WriteLine();

            Console.WriteLine($"Full Cash Save Path:");
            Console.WriteLine(GetFullCashSavePath());
            Console.WriteLine();

            Console.WriteLine("Remittance Payment Save Path:");
            Console.WriteLine(GetRemittanceSavePath());
            Console.WriteLine();

            Console.WriteLine("Name Change File Path:");
            Console.WriteLine(string.IsNullOrWhiteSpace(GetNameChangeFilePath()) ? "Not set yet" : GetNameChangeFilePath());
            Console.WriteLine();

            Console.WriteLine("Microsoft VMS Report File Path:");
            Console.WriteLine(string.IsNullOrWhiteSpace(GetMicrosoftVmsReportFilePath()) ? "Not set yet" : GetMicrosoftVmsReportFilePath());
            Console.WriteLine();

            Console.WriteLine("Nike Tracker Board File Path:");
            Console.WriteLine(string.IsNullOrWhiteSpace(GetNikeTrackerBoardFilePath()) ? "Not set yet" : GetNikeTrackerBoardFilePath());
            Console.WriteLine();

            Console.WriteLine("Press any key to return...");
            Console.ReadKey(true);

            // Reset the entire screen before returning
            ConsoleUi.ResetPage(0);
            ConsoleUi.DrawHeader();
        }

        private static void SavePathSetting(string prompt, string settingsFile, int promptTop)
        {
            Console.CursorVisible = true;

            ClearArea(promptTop, 8);
            Console.SetCursorPosition(0, promptTop);

            Console.WriteLine(prompt);
            Console.Write("> ");

            string? path = Console.ReadLine()?.Trim().Trim('"');

            if (string.IsNullOrWhiteSpace(path))
            {
                Console.WriteLine("No path entered.");
                Console.WriteLine("Press any key to return...");
                Console.ReadKey(true);

                // Reset the entire screen before returning
                ConsoleUi.ResetPage(0);
                ConsoleUi.DrawHeader();
                return;
            }

            if (!Directory.Exists(path))
            {
                Console.WriteLine("Folder not found. Please create it or check the path.");
                Console.WriteLine("Press any key to return...");
                Console.ReadKey(true);

                // Reset the entire screen before returning
                ConsoleUi.ResetPage(0);
                ConsoleUi.DrawHeader();
                return;
            }

            Directory.CreateDirectory(SettingsFolder);
            File.WriteAllText(settingsFile, path);

            Console.WriteLine();
            Console.WriteLine($"Saved path: {path}");
            Console.WriteLine("Press any key to return...");
            Console.ReadKey(true);

            // Reset the entire screen before returning
            ConsoleUi.ResetPage(0);
            ConsoleUi.DrawHeader();
        }

        private static string GetSavedPath(string settingsFile)
        {
            if (File.Exists(settingsFile))
            {
                string savedPath = File.ReadAllText(settingsFile).Trim();

                if (Directory.Exists(savedPath))
                    return savedPath;
            }

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads"
            );
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
                    DrawFullMenu(options, selected, menuTop, menuWidth);
            }

            Console.ResetColor();
            Console.CursorVisible = true;

            return selected;
        }

        private static void SaveFilePathSetting(string prompt, string settingsFile, int promptTop)
        {
            Console.CursorVisible = true;

            ClearArea(promptTop, 8);
            Console.SetCursorPosition(0, promptTop);

            Console.WriteLine(prompt);
            Console.Write("> ");

            string? path = Console.ReadLine()?.Trim().Trim('"');

            if (string.IsNullOrWhiteSpace(path))
            {
                Console.WriteLine("No file path entered.");
                Console.WriteLine("Press any key to return...");
                Console.ReadKey(true);

                // Reset the entire screen before returning
                ConsoleUi.ResetPage(0);
                ConsoleUi.DrawHeader();
                return;
            }

            if (!File.Exists(path))
            {
                Console.WriteLine("File not found. Please check the path.");
                Console.WriteLine("Press any key to return...");
                Console.ReadKey(true);

                // Reset the entire screen before returning
                ConsoleUi.ResetPage(0);
                ConsoleUi.DrawHeader();
                return;
            }

            Directory.CreateDirectory(SettingsFolder);
            File.WriteAllText(settingsFile, path);

            Console.WriteLine();
            Console.WriteLine($"Saved file path: {path}");
            Console.WriteLine("Press any key to return...");
            Console.ReadKey(true);

            // Reset the entire screen before returning
            ConsoleUi.ResetPage(0);
            ConsoleUi.DrawHeader();
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