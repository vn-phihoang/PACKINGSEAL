using System;
using System.IO;

public static class TempPdfCleaner
{
    public static string GetPrintTempFolder()
    {
        string tempFolder = Path.Combine(Path.GetTempPath());

        if (!Directory.Exists(tempFolder))
        {
            Directory.CreateDirectory(tempFolder);
        }

        return tempFolder;
    }

    public static void CleanupOldTempPdfs()
    {
        try
        {
            string tempFolder = GetPrintTempFolder();

            if (!Directory.Exists(tempFolder))
            {
                return;
            }

            string[] files = Directory.GetFiles(tempFolder, "ONEGAI_PRINT_*.pdf");

            foreach (string file in files)
            {
                try
                {
                    FileInfo fileInfo = new FileInfo(file);

                    // Chỉ xoá file cũ để tránh xoá file vừa in hoặc đang spool
                    if (fileInfo.CreationTime < DateTime.Now.AddHours(-1))
                    {
                        DeleteFileSilent(file, 3000);
                    }
                }
                catch(Exception ex)
                {
                    // Ignore individual file error
                    Console.WriteLine($"Error deleting file {file}: {ex.Message}");
                }
            }
        }
        catch(Exception ex)
        {
            // Ignore cleanup error
            Console.WriteLine($"Error during cleanup: {ex.Message}");
        }
    }

    private static void DeleteFileSilent(string filePath, int timeoutMilliseconds)
    {
        if (string.IsNullOrEmpty(filePath))
        {
            return;
        }

        DateTime startTime = DateTime.Now;

        while (true)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }

                return;
            }
            catch
            {
                if ((DateTime.Now - startTime).TotalMilliseconds > timeoutMilliseconds)
                {
                    return;
                }

                System.Threading.Thread.Sleep(300);
            }
        }
    }
}