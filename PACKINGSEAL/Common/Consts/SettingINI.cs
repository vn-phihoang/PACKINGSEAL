using System;
using System.IO;
namespace PACKINGSEAL.Common.Consts
{
    public class SettingINI
    {
        public static readonly string SETTING_INI_FILE_PATH =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), @"AppData\Local\PACKINGSEAL\setting.ini");

        public static string PrinterName { get; set; }
        public static string OutputAndPrint { get; set; }
        public static string OnlyPrint { get; set; }
        public static string OnegaiPackingSealOutputFolderPath { get; set; }
        public static string PDFOutputFolderPath { get; set; }
        public static string ImportFolderPath { get; set; }
        public static string Location { get; set; }
        public static string Brand { get; set; }

        /// <summary>
        /// INI → 変数
        /// </summary>
        public static void Load()
        {
            INIFileUtil iFile = new INIFileUtil(SETTING_INI_FILE_PATH);

            PrinterName = iFile.GetValue("settings", "printer_name", "");
            OnegaiPackingSealOutputFolderPath = iFile.GetValue("settings", "onegai_packing_seal_output_folder_path", Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            //Checkbox Print PDF
            OutputAndPrint = iFile.GetValue("settings", "chk_output_print", "False");
            OnlyPrint = iFile.GetValue("settings", "chk_only_print", "False");
            PDFOutputFolderPath = iFile.GetValue("settings", "pdf_output_folder_path", Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            ImportFolderPath = iFile.GetValue("settings", "import_folder_path", Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            Location = iFile.GetValue("settings", "location", "");
            Brand = iFile.GetValue("settings", "brand", "");
        }

        /// <summary>
        /// 変数 → INI
        /// </summary>
        public static void Save()
        {
            INIFileUtil iFile = new INIFileUtil(SETTING_INI_FILE_PATH);

            iFile["settings", "printer_name"] = PrinterName;
            iFile["settings", "onegai_packing_seal_output_folder_path"] = OnegaiPackingSealOutputFolderPath;
            iFile["settings", "chk_output_print"] = OutputAndPrint;
            iFile["settings", "chk_only_print"] = OnlyPrint;
            iFile["settings", "pdf_output_folder_path"] = PDFOutputFolderPath;
            iFile["settings", "import_folder_path"] = ImportFolderPath;
        }
        public static void SetLocation(string location)
        {
            Location = location;
            INIFileUtil iFile = new INIFileUtil(SETTING_INI_FILE_PATH);
            iFile["settings", "location"] = Location;
        }

        public static void SetBrand(string brand)
        {
            Brand = brand;
            INIFileUtil iFile = new INIFileUtil(SETTING_INI_FILE_PATH);
            iFile["settings", "brand"] = Brand;
        }
    }
}
