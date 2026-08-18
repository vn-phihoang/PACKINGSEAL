using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PACKINGSEAL.Utils
{
    public class PdfPrintManager
    {
        public void showNotification(string type, string message)
        {
            ToastForm frm = new ToastForm(type, message);
            frm.Show();
        }
        private static string GetDefaultPdfViewerPath()
        {
            try
            {
                using (RegistryKey userChoiceKey = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\UserChoice"))
                {
                    string progId = userChoiceKey?.GetValue("Progid") as string;

                    // Nếu không có Progid từ UserChoice, fallback sang ClassesRoot
                    if (string.IsNullOrEmpty(progId))
                    {
                        using (RegistryKey pdfKey = Registry.ClassesRoot.OpenSubKey(".pdf"))
                        {
                            progId = pdfKey?.GetValue("") as string;
                            if (string.IsNullOrEmpty(progId)) return null;
                        }
                    }

                    // Truy cập đến shell\open\command để lấy đường dẫn thực thi
                    using (RegistryKey commandKey = Registry.ClassesRoot.OpenSubKey($"{progId}\\shell\\open\\command"))
                    {
                        if (commandKey == null) return null;

                        string command = commandKey.GetValue("") as string;
                        if (string.IsNullOrEmpty(command)) return null;

                        // Trích xuất đường dẫn thực thi từ chuỗi command
                        int firstQuote = command.IndexOf('"');
                        int secondQuote = command.IndexOf('"', firstQuote + 1);
                        if (firstQuote >= 0 && secondQuote > firstQuote)
                        {
                            return command.Substring(firstQuote + 1, secondQuote - firstQuote - 1);
                        }

                        // Nếu không có dấu ngoặc kép, lấy đến khoảng trắng đầu tiên
                        int spaceIndex = command.IndexOf(' ');
                        if (spaceIndex > 0)
                        {
                            return command.Substring(0, spaceIndex);
                        }

                        return command.Trim();
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        public static void PrintPdfToSpecificPrinter(string pdfPath, string printerName)
        {
            string exePath = GetDefaultPdfViewerPath();
            if (string.IsNullOrEmpty(exePath) || !File.Exists(exePath))
            {
                MessageBox.Show("Default PDF reader not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string arguments = $"/t \"{pdfPath}\" \"{printerName}\"";
            try
            {
                /*ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = pdfPath,
                    Verb = "printto",
                    Arguments = $"\"{printerName}\"",
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    UseShellExecute = true
                };*/

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = arguments,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    UseShellExecute = true
                };
                var process = Process.Start(psi);
                new PdfPrintManager().showNotification("INFO", "The PDF File printing process is running.");
                if (process != null && !process.HasExited)
                {
                    process.WaitForExit(10000);
                    /*if (!exited)
                    {
                        new PdfPrintManager().showNotification("WARNING", "Printing process is taking too long or may be stuck.");
                        process.Kill();
                    }*/
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error printing PDF: " + ex.Message, "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
