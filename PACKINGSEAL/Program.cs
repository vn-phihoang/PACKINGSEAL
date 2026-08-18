using PACKINGSEAL;
using PACKINGSEAL.Common.Consts;
using System;
using System.Threading.Tasks;
using System.Windows.Forms;

static class Program
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        SettingINI.Load();

        FrmMain frmMain = new FrmMain();

        frmMain.Shown += (s, e) =>
        {
            Task.Run(() =>
            {
                TempPdfCleaner.CleanupOldTempPdfs();
            });
        };

        Application.Run(frmMain);
    }
}