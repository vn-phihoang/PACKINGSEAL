using iText.IO.Font;
using iText.Layout.Font;
using Microsoft.Win32;
using PdfiumViewer;
using System;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PACKINGSEAL.Utils
{
    public class PDFSetting
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
                new PDFSetting().showNotification("INFO", "The PDF File printing process is running.");
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

        //public static void PrintPdfFileWithPDFViewer(string pdfPath, string printerName)
        //{
        //    try
        //    {
        //        using (var document = PdfDocument.Load(pdfPath))
        //        {
        //            using (var printDocument = document.CreatePrintDocument())
        //            {
        //                // Gán máy in cụ thể
        //                printDocument.PrinterSettings.PrinterName = printerName;

        //                // Thiết lập thêm nếu cần
        //                printDocument.PrinterSettings.Copies = 1;
        //                printDocument.DefaultPageSettings.Landscape = false;

        //                // In
        //                printDocument.Print();
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error printing PDF: " + ex.Message, "Print Error",
        //            MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}

        public static void PrintPdfFileWithPDFViewer(string pdfPath, string printerName)
        {
            try
            {
                if (string.IsNullOrEmpty(pdfPath) || !File.Exists(pdfPath))
                {
                    MessageBox.Show(
                        "PDF file does not exist: " + pdfPath,
                        "Print Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    return;
                }

                if (string.IsNullOrEmpty(printerName))
                {
                    MessageBox.Show(
                        "Printer name is empty.",
                        "Print Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    return;
                }

                // Đảm bảo file PDF đã sẵn sàng trước khi load
                WaitUntilPdfReady(pdfPath, 30000);

                byte[] pdfBytes = File.ReadAllBytes(pdfPath);

                using (var ms = new MemoryStream(pdfBytes))
                using (var document = PdfiumViewer.PdfDocument.Load(ms))
                using (var printDocument = document.CreatePrintDocument(PdfiumViewer.PdfPrintMode.CutMargin))
                {
                    printDocument.PrinterSettings.PrinterName = printerName;

                    if (!printDocument.PrinterSettings.IsValid)
                    {
                        MessageBox.Show(
                            "Printer is not valid: " + printerName,
                            "Print Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);

                        return;
                    }

                    printDocument.PrinterSettings.Copies = 1;
                    printDocument.PrintController = new StandardPrintController();

                    PaperSize a4 = new PaperSize("A4", 827, 1169);

                    printDocument.DefaultPageSettings.PaperSize = a4;
                    printDocument.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
                    printDocument.DefaultPageSettings.Landscape = false;
                    printDocument.OriginAtMargins = false;

                    printDocument.PrinterSettings.DefaultPageSettings.PaperSize = a4;
                    printDocument.PrinterSettings.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
                    printDocument.PrinterSettings.DefaultPageSettings.Landscape = false;

                    printDocument.QueryPageSettings += (s, e) =>
                    {
                        e.PageSettings.PaperSize = a4;
                        e.PageSettings.Margins = new Margins(0, 0, 0, 0);
                        e.PageSettings.Landscape = false;
                    };

                    printDocument.Print();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error printing PDF: " + ex.Message,
                    "Print Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public static void WaitUntilPdfReady(string filePath, int timeoutMs)
        {
            var start = DateTime.Now;
            long lastSize = -1;
            int stableCount = 0;

            while ((DateTime.Now - start).TotalMilliseconds < timeoutMs)
            {
                try
                {
                    if (File.Exists(filePath))
                    {
                        FileInfo fi = new FileInfo(filePath);
                        long currentSize = fi.Length;

                        if (currentSize > 0 && currentSize == lastSize)
                        {
                            stableCount++;
                        }
                        else
                        {
                            stableCount = 0;
                            lastSize = currentSize;
                        }

                        // File size stable trong vài vòng kiểm tra
                        if (stableCount >= 3)
                        {
                            // Thử mở exclusive để chắc chắn không process nào còn giữ file
                            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                            {
                                if (fs.Length > 0)
                                {
                                    return;
                                }
                            }
                        }
                    }
                }
                catch (IOException)
                {
                    // File vẫn đang bị process khác giữ
                }
                catch (UnauthorizedAccessException)
                {
                    // File chưa sẵn sàng hoặc không đủ quyền
                }

                Thread.Sleep(300);
            }

            throw new TimeoutException("PDF file is still locked or not ready: " + filePath);
        }

        public static FontProvider CreateFontProviderWithCustomFonts()
        {
            FontProvider fontProvider = new FontProvider();
            Assembly assembly = Assembly.GetEntryAssembly();
            string regularFontPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\MS-UIGothic.ttf";
            if (File.Exists(regularFontPath))
            {
                FontProgram customFontProgram = FontProgramFactory.CreateFont(regularFontPath);
                fontProvider.AddFont(customFontProgram);
            }
            else
            {
                throw new FileNotFoundException($"Font file not found: {regularFontPath}");
            }

            string boldFontPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\NotoSansJP-ExtraBold.ttf";
            if (File.Exists(boldFontPath))
            {
                FontProgram customFontProgram2 = FontProgramFactory.CreateFont(boldFontPath);
                fontProvider.AddFont(customFontProgram2);
            }
            else
            {
                throw new FileNotFoundException($"Font file not found: {boldFontPath}");
            }

            return fontProvider;
        }

    }
}
