using iText.Html2pdf;
using iText.IO.Font;
using iText.Kernel.Events;
using iText.Kernel.Pdf;
using iText.Layout.Font;
using NPOI.SS.UserModel;
using PACKINGSEAL.Models;
using PACKINGSEAL.Utils;
using System;
using System.Data;
using System.Deployment.Application;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PACKINGSEAL
{
    public partial class FrmPackingSealSakurai : Form
    {
        public string imagePath = string.Empty;
        private string htmlPath = string.Empty;
        private string pdfOutputPath = string.Empty;
        private Assembly assembly = Assembly.GetExecutingAssembly();

        public FrmPackingSealSakurai(FilePathModel filePathModel)
        {
            InitializeComponent();
            string UserName = Environment.MachineName;
            lb_info_1.Text = "HC | Sakurai Packing Onegai Seal- " + UserName;
            string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
            var version = "";
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                version = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            this.Text = "SAKURAI PACKING ONEGAI SEAL (梱包シール) Ver: " + version;
            st_count_1.Text = "0";
            textBoxFile.Clear();
            cleanTempPDFFiles();
            imagePath = filePathModel.ImagePath;
            htmlPath = filePathModel.HtmlPath;
            if (webView21.CoreWebView2 != null)
            {
                webView21.CoreWebView2.Navigate($@"{Path.GetDirectoryName(assembly.Location)}\Template\Default.html");
            }
        }

        private void cleanTempPDFFiles()
        {
            try
            {
                string tempPath = Path.GetTempPath();
                var tempFiles = Directory.GetFiles(tempPath, "ONEGAI_*.pdf");
                foreach (var file in tempFiles)
                {
                    File.Delete(file);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot clean temporary PDF files: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        public void showNotification(string type, string message)
        {
            ToastForm frm = new ToastForm(type, message);
            frm.Show();
        }

        private void RenderHTMLToWebView(string tableContent)
        {
            string htmlContent = File.ReadAllText(htmlPath);
            htmlContent = htmlContent.Replace("[htmlContent]", tableContent);
            webView21.NavigateToString(htmlContent);
        }

        private void LoadDataGirdview(string filePath)
        {
            try
            {
                ExcelUtil excelUtil = new ExcelUtil();
                DataTable dt = new DataTable();
                dt = new DataTable();
                dt.Columns.Add("Date");
                dt.Columns.Add("InvoiNo");
                dt.Columns.Add("CartonNo");
                dt.Columns.Add("PCS");
                var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                var workbook = WorkbookFactory.Create(stream);
                string prefix = "SAKURAI";
                ISheet sheetAssortmentList = null;
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    string sheetName = workbook.GetSheetName(i);
                    if (sheetName.StartsWith(prefix))
                    {
                        sheetAssortmentList = workbook.GetSheetAt(i);
                        break;
                    }
                }

                if (sheetAssortmentList == null)
                {
                    MessageBox.Show($"Sheet with name not found: {prefix}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string date = excelUtil.GetCellValueWithMerge(sheetAssortmentList, 13, 13);
                string invoiNo = excelUtil.GetCellValueWithMerge(sheetAssortmentList, 12, 13);
                string cartonNo = excelUtil.GetCellValueWithMerge(sheetAssortmentList, 7, 2);

                for (int i = 21; i <= sheetAssortmentList.LastRowNum; i++)
                {
                    if (string.IsNullOrEmpty(excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 4)))
                    {
                        continue;
                    }
                    bool isNumber = int.TryParse(excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 4), out int number);
                    if (!isNumber)
                    {
                        continue;
                    }

                    DataRow dataRow = dt.NewRow();
                    dataRow["Date"] = date;
                    dataRow["InvoiNo"] = invoiNo;
                    dataRow["CartonNo"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 4);
                    dataRow["PCS"] = "1";

                    dt.Rows.Add(dataRow.ItemArray);
                }

                dgvData.DataSource = dt;
                st_count_1.Text = dgvData.Rows.Count.ToString();
                webView21.CoreWebView2.Navigate($@"{Path.GetDirectoryName(assembly.Location)}\Template\Default.html");
                //showNotification("SUCCESS", "Read excel file completed.");

                var resModel = new ImportResponseModel
                {
                    IsSuccess = true,
                    Message = "Read excel file completed.",
                    RecordsImported = dgvData.Rows.Count,
                    RecordsFailed = 0,
                    SourceName = Path.GetFileName(filePath),
                    ErrorDetail = string.Empty
                };

                ShowImportNotification(resModel);
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cursor = Cursors.Default;
            }
        }

        private void dgvData_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }
            var row = dgvData.Rows[e.RowIndex];
            string htmlContent = string.Empty;
            HtmlResponse templateHTML = new HtmlResponse();
            SakuraiModel sakuraiModel = new SakuraiModel();
            sakuraiModel.InvoiNo = row.Cells["InvoiNo"].Value?.ToString() ?? string.Empty;
            sakuraiModel.Date = row.Cells["Date"].Value?.ToString() ?? string.Empty;
            sakuraiModel.CartonNo = row.Cells["CartonNo"].Value?.ToString() ?? string.Empty;
            htmlContent = templateHTML.SakuraiToWebview(sakuraiModel);

            RenderHTMLToWebView(htmlContent);
        }

        private async void FrmPackingSealSakurai_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.frmSakuraiSize.Width != 0 && Properties.Settings.Default.frmSakuraiSize.Height != 0)
            {
                this.Size = Properties.Settings.Default.frmSakuraiSize;
            }
            if (Properties.Settings.Default.frmSakuraiLocation.X != 0 && Properties.Settings.Default.frmSakuraiLocation.Y != 0)
            {
                this.StartPosition = FormStartPosition.CenterScreen;
                this.Location = Properties.Settings.Default.frmSakuraiLocation;
            }
            this.WindowState = Properties.Settings.Default.frmSakuraiState;

            await webView21.EnsureCoreWebView2Async(null);
            webView21.CoreWebView2.Navigate($@"{Path.GetDirectoryName(assembly.Location)}\Template\Default.html");

            webView21.CoreWebView2.Settings.IsZoomControlEnabled = true;
            using (Graphics g = Graphics.FromHwnd(this.Handle))
            {
                float dpiX = g.DpiX;
                float dpiY = g.DpiY;
                float scaleFactor = (dpiX + dpiY) / 2f / 96f;

                webView21.ZoomFactor = scaleFactor;
            }
            if (dgvData.Rows.Count > 0)
            {
                var row = dgvData.Rows[0];
                string htmlContent = string.Empty;
                HtmlResponse templateHTML = new HtmlResponse();
                SakuraiModel sakuraiModel = new SakuraiModel();
                sakuraiModel.InvoiNo = row.Cells["InvoiNo"].Value?.ToString() ?? string.Empty;
                sakuraiModel.Date = row.Cells["Date"].Value?.ToString() ?? string.Empty;
                sakuraiModel.CartonNo = row.Cells["CartonNo"].Value?.ToString() ?? string.Empty;
                htmlContent = templateHTML.SakuraiToWebview(sakuraiModel);
                RenderHTMLToWebView(htmlContent);
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            PrintPDF();
        }

        private void PrintPDF()
        {
            Cursor = Cursors.WaitCursor;

            Utils.PDFSetting pdfPrintManager = new Utils.PDFSetting();
            if (dgvData.Rows.Count == 0)
            {
                MessageBox.Show("Please enter data before printing.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Cursor = Cursors.Default;
                return;
            }
            try
            {
                bool isSuccess = GeneratePDF();
                if (isSuccess)
                {
                    if (Common.Consts.SettingINI.OutputAndPrint != "True" && Common.Consts.SettingINI.OnlyPrint != "True")
                    {
                        MessageBox.Show($@"PDF file created successfully on {Common.Consts.SettingINI.PDFOutputFolderPath}!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        base.Cursor = Cursors.Default;
                        return;
                    }
                    string PDFOutputFolderPath = Common.Consts.SettingINI.PDFOutputFolderPath;
                    if (string.IsNullOrEmpty(PDFOutputFolderPath))
                    {
                        PDFOutputFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    }
                    string printerName = Common.Consts.SettingINI.PrinterName;
                    string invoiNo = dgvData.Rows[0].Cells["InvoiNo"].Value?.ToString() ?? string.Empty;
                    if (string.IsNullOrEmpty(printerName))
                    {
                        using (PrintDialog printDialog = new PrintDialog())
                        {
                            printDialog.AllowSomePages = false;
                            printDialog.ShowHelp = false;
                            printDialog.UseEXDialog = true;

                            if (printDialog.ShowDialog() == DialogResult.OK)
                            {
                                printerName = printDialog.PrinterSettings.PrinterName;
                            }
                            else
                            {
                                MessageBox.Show("No printer selected.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                        }
                    }

                    Utils.PDFSetting.PrintPdfToSpecificPrinter(pdfOutputPath, printerName);

                    System.Threading.Thread.Sleep(3000);
                    //new Utils.PDFSetting().showNotification("SUCCESS", "PDF printed successfully.");

                    var resModel = new PrintResponseModel
                    {
                        Brand = "SAKURAI",
                        PoNumber = invoiNo,
                        PrinterName = printerName,
                        PrintTime = DateTime.Now,
                        MsgVal = $"I/V No. {invoiNo} has been successfully printed!",
                    };

                    ShowNotificationPrint(resModel);
                }
                else
                {
                    MessageBox.Show("Failed to create PDF file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cursor = Cursors.Default;
                return;
            }
            Cursor = Cursors.Default;
        }

        private bool GeneratePDF()
        {
            StringBuilder tableBuilder = new StringBuilder();
            int index = 0;
            HtmlResponse templateHTML = new HtmlResponse();

            if (dgvData.Rows.Count == 0)
            {
                MessageBox.Show("Please enter data before printing.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            try
            {
                foreach (DataGridViewRow row in dgvData.Rows)
                {
                    if (row.IsNewRow || int.Parse(row.Cells["PCS"].Value?.ToString()) == 0)
                    {
                        continue;
                    }

                    for (int i = 0; i < int.Parse(row.Cells["PCS"].Value.ToString()); i++)
                    {
                        SakuraiModel sakuraiModel = new SakuraiModel();
                        sakuraiModel.InvoiNo = row.Cells["InvoiNo"].Value?.ToString() ?? string.Empty;
                        sakuraiModel.CartonNo = row.Cells["CartonNo"].Value?.ToString() ?? string.Empty;
                        sakuraiModel.Date = row.Cells["Date"].Value?.ToString() ?? string.Empty;
                        sakuraiModel.Index = index;

                        string tableContent = templateHTML.SakuraiToPDF(sakuraiModel);
                        tableBuilder.AppendLine(tableContent);
                        index++;
                    }
                }

                if (index == 0)
                {
                    MessageBox.Show("Invalid quantity in PCS column.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else if (index % 2 != 0)
                {
                    tableBuilder.AppendLine("</page>");
                }

                if (!File.Exists(htmlPath))
                {
                    MessageBox.Show("HTML file not found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // Tính số trang và ngày in
                int totalPages = (int)Math.Ceiling(index / 4.0); // 4 item = 1 page
                string today = DateTime.Now.ToString("yyyy/MM/dd");

                string htmlTemplate = File.ReadAllText(htmlPath);
                string fullHtml = htmlTemplate.Replace("[htmlContent]", tableBuilder.ToString());

                // Xác định đường dẫn PDF
                if (Common.Consts.SettingINI.OnlyPrint != "True")
                {
                    if (string.IsNullOrEmpty(Common.Consts.SettingINI.PDFOutputFolderPath))
                    {
                        Common.Consts.SettingINI.PDFOutputFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    }
                    pdfOutputPath = Path.Combine(Common.Consts.SettingINI.PDFOutputFolderPath, $"ONEGAI_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");
                }
                else
                {
                    pdfOutputPath = Path.Combine(Path.GetTempPath(), $"ONEGAI_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");
                }

                using (FileStream pdfDest = new FileStream(pdfOutputPath, FileMode.Create))
                {
                    var fontProvider = SakuraiCreateFontProviderWithCustomFonts();
                    var props = new ConverterProperties();
                    props.SetFontProvider(fontProvider);
                    props.SetCharset("utf-8");

                    var pdfWriter = new PdfWriter(pdfDest);
                    var pdfDoc = new PdfDocument(pdfWriter);

                    // Gắn event handler để vẽ header cho mỗi trang
                    pdfDoc.AddEventHandler(PdfDocumentEvent.END_PAGE, new HeaderFooterHandler(totalPages, today));

                    HtmlConverter.ConvertToPdf(fullHtml, pdfDoc, props);
                    pdfDoc.Close();
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating PDF: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cursor = Cursors.Default;
                return false;
            }
        }

        private void LoadImportFile()
        {
            try
            {
                string fileExt = string.Empty;
                string fileName = string.Empty;
                textBoxFile.Text = string.Empty;
                HtmlResponse templateHTML = new HtmlResponse();
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files|*.xls;*.xlsx",
                    Title = "Select file Excel",
                    InitialDirectory = Common.Consts.SettingINI.ImportFolderPath
                };

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    Cursor = Cursors.Default;
                    return;
                }

                string filePath = openFileDialog.FileName;
                fileExt = Path.GetExtension(filePath).ToLower();
                fileName = Path.GetFileName(filePath);

                if (fileExt != ".xls" && fileExt != ".xlsx")
                {
                    MessageBox.Show("Please select Excel file (.xls or .xlsx)", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Cursor = Cursors.Default;
                    return;
                }

                if (!fileName.StartsWith("(ホーチミン雛形)PACKING LIST SAKURAI"))
                {
                    MessageBox.Show("File name must start with '(ホーチミン雛形)PACKING LIST SAKURAI'", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Cursor = Cursors.Default;
                    return;
                }

                string preparedFile = templateHTML.PrepareFile(filePath);
                textBoxFile.Text = filePath;
                bool isTempFile = preparedFile != filePath;

                Cursor = Cursors.WaitCursor;
                using (FileStream fs = new FileStream(preparedFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    LoadDataGirdview(preparedFile);
                }
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cursor = Cursors.Default;
                return;
            }
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            LoadImportFile();
        }

        private void textBoxFile_DragDrop(object sender, DragEventArgs e)
        {
            textBoxFile.Text = string.Empty;
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            HtmlResponse templateHTML = new HtmlResponse();
            foreach (string originalPath in files)
            {
                string preparedFile = templateHTML.PrepareFile(originalPath);
                textBoxFile.Text += originalPath + Environment.NewLine;
                bool isTempFile = preparedFile != originalPath;
                string extension = Path.GetExtension(originalPath).ToLower();

                if (extension != ".xls" && extension != ".xlsx")
                {
                    MessageBox.Show("Please select Excel file (.xls or .xlsx)", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (!Path.GetFileName(originalPath).StartsWith("(ホーチミン雛形)PACKING LIST SAKURAI"))
                {
                    MessageBox.Show("File name must start with '(ホーチミン雛形)PACKING LIST SAKURAI'", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                try
                {
                    using (FileStream fs = new FileStream(preparedFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        LoadDataGirdview(preparedFile);
                    }

                    if (isTempFile && File.Exists(preparedFile))
                    {
                        try
                        {
                            File.Delete(preparedFile);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Cannot delete temporary file: " + ex.Message);
                        }
                    }
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Cannot read file: " + ex.Message);
                }
            }
        }

        private void textBoxFile_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void dgvData_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            string stt = (e.RowIndex + 1).ToString();
            using (SolidBrush brush = new SolidBrush(dgvData.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString(stt,
                    dgvData.RowHeadersDefaultCellStyle.Font,
                    brush,
                    e.RowBounds.Location.X + 10,
                    e.RowBounds.Location.Y + 4);
            }
        }

        private void FrmPackingSealSakurai_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.frmSakuraiState = this.WindowState;
            if (this.WindowState == FormWindowState.Normal)
            {
                Properties.Settings.Default.frmSakuraiLocation = this.Location;
                Properties.Settings.Default.frmSakuraiSize = this.Size;
            }
            else
            {
                Properties.Settings.Default.frmSakuraiLocation = this.RestoreBounds.Location;
                Properties.Settings.Default.frmSakuraiSize = this.RestoreBounds.Size;
            }

            Properties.Settings.Default.Save();
            webView21.Dispose();
            Application.Exit();
        }

        private void dgvData_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvData.Columns[e.ColumnIndex].Name == "PCS")
            {
                var pcs = dgvData.Rows[e.RowIndex].Cells[e.ColumnIndex];

                if (string.IsNullOrEmpty(pcs.Value.ToString()))
                {
                    pcs.Value = "0";
                }
            }

            if (dgvData.Columns[e.ColumnIndex].Name == "Quantity")
            {
                var quantity = dgvData.Rows[e.RowIndex].Cells[e.ColumnIndex];
                if (string.IsNullOrEmpty(quantity.Value.ToString()))
                {
                    quantity.Value = "0";
                }
            }
        }

        private void dgvData_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgvData.CurrentCell.ColumnIndex == dgvData.Columns["PCS"].Index)
            {
                System.Windows.Forms.TextBox tb = e.Control as System.Windows.Forms.TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= Tb_KeyPress_Strict;
                    tb.KeyPress += Tb_KeyPress_Strict;
                }
            }
        }

        private void Tb_KeyPress_Strict(object sender, KeyPressEventArgs e)
        {
            // Chỉ cho phép số (0–9), không cho phép khoảng trắng hay ký tự khác
            if (!Regex.IsMatch(e.KeyChar.ToString(), @"^\d$") && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void toolStripButtonZoomOut_Click(object sender, EventArgs e)
        {
            webView21.ZoomFactor -= 0.1;
            if (webView21.ZoomFactor < 1.0)
            {
                webView21.ZoomFactor = 1.0;
            }
            ResizeRuler(webView21.ZoomFactor);
        }

        private void toolStripButtonZoomIn_Click(object sender, EventArgs e)
        {
            webView21.ZoomFactor += 0.1;
            ResizeRuler(webView21.ZoomFactor);
        }

        private void toolStripButtonReset_Click(object sender, EventArgs e)
        {
            webView21.ZoomFactor = 1.0;
            ResizeRuler(webView21.ZoomFactor);
        }

        private void ResizeRuler(double ZoomFactor)
        {
            verticalRuler1.ZoomFactor = ZoomFactor;
            verticalRuler1.Invalidate();
            horizontalRuler1.ZoomFactor = (float)webView21.ZoomFactor;
            horizontalRuler1.Invalidate();
        }

        private void webView_ZoomFactorChanged(object sender, EventArgs e)
        {
            float zoom = (float)webView21.ZoomFactor;
            if (zoom < 1.0f)
            {
                webView21.ZoomFactor = 1.0f;
                return;
            }
            ResizeRuler(zoom);
        }

        private void btnOpition_Click(object sender, EventArgs e)
        {
            var frm = new FrmSettings();
            if (!Application.OpenForms.OfType<FrmSettings>().Any())
            {
                if (frm == null)
                    frm = new FrmSettings();
                frm.Show();
            }
            else
            {
                frm.Focus();
            }
        }

        private void buttonBack_Click(object sender, EventArgs e)
        {
            FrmMain frm = new FrmMain();
            frm.Show();
            this.Hide();
        }

        private void FrmPackingSealSakurai_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.P)
            {
                PrintPDF();
            }
            else if (e.Control && e.KeyCode == Keys.O)
            {
                LoadImportFile();
            }
        }

        private static FontProvider SakuraiCreateFontProviderWithCustomFonts()
        {
            FontProvider fontProvider = new FontProvider();
            Assembly assembly = Assembly.GetEntryAssembly();
            string regularFontPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\MS-Gothic.ttf";
            if (File.Exists(regularFontPath))
            {
                FontProgram customFontProgram = FontProgramFactory.CreateFont(regularFontPath);
                fontProvider.AddFont(customFontProgram);
            }
            else
            {
                throw new FileNotFoundException($"Font file not found: {regularFontPath}");
            }

            string boldFontPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\BIZ-UDGothic-Bold.ttf";
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

        private void ShowNotificationPrint(PrintResponseModel resModel)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(resModel.MsgVal))
                {
                    notifyIcon1.BalloonTipTitle = $"Announcement {resModel.Brand}";
                    notifyIcon1.BalloonTipText =
                        $"{resModel.MsgVal}\n" +
                        $"I/V No.: {resModel.PoNumber}\n" +
                        $"Printer: {resModel.PrinterName}\n" +
                        $"Time: {resModel.PrintTime:yyyy/MM/dd HH:mm:ss}";

                    notifyIcon1.ShowBalloonTip(2000);
                    notifyIcon1.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowImportNotification(ImportResponseModel resModel)
        {
            try
            {
                string status = resModel.IsSuccess ? "Success" : "Failed";

                if (resModel.Message != null)
                {
                    notifyIcon1.Icon = resModel.IsSuccess ? SystemIcons.Information : SystemIcons.Error;
                    notifyIcon1.BalloonTipTitle = "Data Import Notification";
                    notifyIcon1.BalloonTipText =
                        $"Status: {status}\n" +
                        $"File: {resModel.SourceName}\n" +
                        $"Success: {resModel.RecordsImported}, Error: {resModel.RecordsFailed}\n" +
                        $"{resModel.Message}";

                    notifyIcon1.ShowBalloonTip(4000);
                    notifyIcon1.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
