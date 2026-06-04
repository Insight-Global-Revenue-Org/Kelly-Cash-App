using System.Windows.Forms;

namespace KellyCashApp
{
    internal class FileSelector
    {
        public static string? SelectFile(string prompt, int startLine)
        {
            if (Settings.GetFileSelectionType() == "Input File Explorer")
                return OpenFileExplorer();

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