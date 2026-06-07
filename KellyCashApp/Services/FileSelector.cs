using System.Windows.Forms;
using KellyCashApp.Configuration;

namespace KellyCashApp.Services
{
    internal class FileSelector
    {
        public static string? SelectFile(string prompt, int startLine)
        {
            if (Settings.GetFileSelectionType() == "Input File Explorer")
            {
                Console.CursorVisible = false;
                ClearArea(startLine, 6);
                Console.SetCursorPosition(0, startLine);

                Console.WriteLine(prompt);
                Console.WriteLine("Select in the File Explorer window...");

                string? selectedPath = OpenFileExplorer();

                ClearArea(startLine, 6);
                return selectedPath;
            }

            return PromptForFilePath(prompt, startLine);
        }

        private static string? OpenFileExplorer()
        {
            string? selectedPath = null;

            Thread thread = new Thread(() =>
            {
                using OpenFileDialog dialog = new OpenFileDialog();

                dialog.Title = "Select File";
                dialog.Filter = "Excel/Text Files|*.xlsx;*.xls;*.txt|All Files|*.*";
                dialog.InitialDirectory = _lastDirectory;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    selectedPath = dialog.FileName;

                    string? directory = Path.GetDirectoryName(dialog.FileName);

                    if (!string.IsNullOrWhiteSpace(directory))
                        _lastDirectory = directory;
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            return selectedPath;
        }

        private static void ClearArea(int startLine, int numberOfLines)
        {
            for (int i = 0; i < numberOfLines; i++)
            {
                int line = startLine + i;

                if (line < 0 || line >= Console.BufferHeight)
                    continue;

                Console.SetCursorPosition(0, line);
                Console.Write(new string(' ', Console.WindowWidth - 1));
            }

            Console.SetCursorPosition(0, startLine);
        }

        private static string _lastDirectory =
            Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        "Downloads"
        );

        private static string? PromptForFilePath(string prompt, int startLine)
        {
            Console.CursorVisible = true;

            Console.SetCursorPosition(0, startLine);
            Console.WriteLine(prompt);
            Console.Write("> ");

            return Console.ReadLine()?.Trim().Trim('"');
        }
    }
}