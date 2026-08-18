using PACKINGSEAL.Common.Consts;
using System;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace PACKINGSEAL
{
    public partial class FrmSettings : Form
    {
        public FrmSettings()
        {
            InitializeComponent();
        }
        private void FrmOptions_Load(object sender, EventArgs e)
        {
            GetPrinterList();
            string output_path = SettingINI.PDFOutputFolderPath;
            txtPDFOutputPath.Text = output_path;

            string outputAndPrint = SettingINI.OutputAndPrint?.ToString();

            if (!string.IsNullOrEmpty(outputAndPrint))
            {
                chk_output_print.Checked =  Convert.ToBoolean(outputAndPrint);
            }

            string onlyPrint = SettingINI.OnlyPrint.ToString();
            if(!string.IsNullOrEmpty(onlyPrint))
            {
                chk_only_print.Checked = Convert.ToBoolean(onlyPrint);
            }
            textBoxImportPath.Text = SettingINI.ImportFolderPath;
            string selected_printer = SettingINI.PrinterName;
            cbo_printer_list.SelectedItem = selected_printer;
        }
        private void GetPrinterList()
        {
            PrintDocument prtdoc = new PrintDocument();
            string strDefaultPrinter = prtdoc.PrinterSettings.PrinterName;

            cbo_printer_list.Items.Add("");
            foreach (String strPrinter in PrinterSettings.InstalledPrinters)
            {
                cbo_printer_list.Items.Add(strPrinter);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnPDFOutputPath_Click(object sender, EventArgs e)
        {
            try
            {
                using (var folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "Please select a folder for PDF output.";
                    folderDialog.ShowNewFolderButton = true;
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        txtPDFOutputPath.Text = folderDialog.SelectedPath;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void btnImportPath_Click(object sender, EventArgs e)
        {
            try
            {
                using (var folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "Please select a folder for PDF output.";
                    folderDialog.ShowNewFolderButton = true;
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        textBoxImportPath.Text = folderDialog.SelectedPath;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CheckedChanged(object sender, EventArgs e)
        {
            CheckBox changedCheckBox = sender as CheckBox;
            if (changedCheckBox.Checked)
            {
                foreach (Control control in groupBox1.Controls)
                {
                    if (control is CheckBox && control != changedCheckBox)
                    {
                        ((CheckBox)control).Checked = false;
                    }
                }
            }
        }

        private void btnRegister_Click(object sender, EventArgs e)
        {
            try
            {
                SettingINI.PDFOutputFolderPath = txtPDFOutputPath.Text;
                SettingINI.OnegaiPackingSealOutputFolderPath = txtPDFOutputPath.Text;
                SettingINI.OutputAndPrint = chk_output_print.Checked.ToString();
                SettingINI.OnlyPrint = chk_only_print.Checked.ToString();
                SettingINI.PrinterName = cbo_printer_list.SelectedItem.ToString();
                SettingINI.ImportFolderPath = textBoxImportPath.Text;
                SettingINI.Save();
                SettingINI.Load();

                MessageBox.Show("Option setting has been saved.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}