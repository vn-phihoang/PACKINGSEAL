using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace PACKINGSEAL.Common.Consts
{
    /// <summary>
    /// ***************************************************************<br />
    /// INIファイル操作ユーティリティクラス<br />
    /// ***************************************************************<br />
    /// </summary>
    public class INIFileUtil
    {
        [DllImport("kernel32.dll")]
        private static extern int GetPrivateProfileString(
            string lpApplicationName,
            string lpKeyName,
            string lpDefault,
            StringBuilder lpReturnedstring,
            int nSize,
            string lpFileName);

        [DllImport("kernel32.dll")]
        private static extern int WritePrivateProfileString(
            string lpApplicationName,
            string lpKeyName,
            string lpstring,
            string lpFileName);

        string filePath;

        /// <summary>
        /// ***************************************************************<br />
        /// ファイル名を指定して初期化する<br />
        /// ファイルが存在しない場合は初回書き込み時に作成される<br />
        /// ***************************************************************<br />
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        public INIFileUtil(string filePath)
        {
            this.filePath = filePath;

            string fol = Path.GetDirectoryName(filePath);
            if (Directory.Exists(fol) == false)
            {
                Directory.CreateDirectory(fol);
            }

        }

        /// <summary>
        /// ***************************************************************<br />
        /// sectionとkeyからiniファイルの設定値を取得／設定する<br />
        /// ***************************************************************<br />
        /// </summary>
        /// <param name="section">セクション名</param>
        /// <param name="key">キー名</param>
        /// <param name="capacity">lpReturnedStringのバイト数 ※デフォルト:256</param>
        /// <returns>引数のセクション名・キー名に紐づく設定値
        ///          存在しない場合は""</returns>
        public string this[string section, string key, int capacity = 256]
        {
            set
            {
                WritePrivateProfileString(section, key, value, filePath);
            }
            get
            {
                StringBuilder sb = new StringBuilder(capacity);
                GetPrivateProfileString(section, key, string.Empty, sb, sb.Capacity, filePath);
                return sb.ToString();
            }
        }

        /// <summary>
        /// ***************************************************************<br />
        /// sectionとkeyからiniファイルの設定値を取得する<br />
        /// ***************************************************************<br />
        /// </summary>
        /// <param name="section">セクション名</param>
        /// <param name="key">キー名</param>
        /// <param name="defaultvalue">引数のセクション名・キー名に紐づく設定値が存在しない場合の返却値</param>
        /// <param name="capacity">lpReturnedStringのバイト数 ※デフォルト:256</param>
        /// <returns>引数のセクション名・キー名に紐づく設定値
        ///          存在しない場合は引数のdefaultvalue</returns>
        public string GetValue(string section, string key, string defaultvalue, int capacity = 256)
        {
            StringBuilder sb = new StringBuilder(capacity);
            GetPrivateProfileString(section, key, defaultvalue, sb, sb.Capacity, filePath);
            return sb.ToString();
        }
        public string GetDefaultPdfViewer()
        {
            try
            {
                // Đường dẫn registry chứa thông tin ứng dụng mở file .pdf
                string keyPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\UserChoice";
                using (RegistryKey userChoiceKey = Registry.CurrentUser.OpenSubKey(keyPath))
                {
                    if (userChoiceKey != null)
                    {
                        object progId = userChoiceKey.GetValue("ProgId");
                        if (progId != null)
                        {
                            return progId.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return "Unable to check: " + ex.Message;
            }

            return "Default PDF software not determined.";
        }
    }
}
