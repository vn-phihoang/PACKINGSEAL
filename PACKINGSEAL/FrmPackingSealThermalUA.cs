using iText.Html2pdf;
using NPOI.SS.UserModel;
using PACKINGSEAL.Models;
using PACKINGSEAL.Utils;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
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
    public partial class FrmPackingSealThermalUA : Form
    {
        public string imagePath = string.Empty;
        private string htmlPath = string.Empty;
        private string pdfOutputPath = string.Empty;
        private Assembly assembly = Assembly.GetExecutingAssembly();

        public FrmPackingSealThermalUA(FilePathModel filePathModel)
        {
            InitializeComponent();
            string UserName = Environment.MachineName;
            lb_info_1.Text = "HC | Thermal Underamour Packing Onegai Seal- " + UserName;
            string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
            var version = "";
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                version = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            this.Text = "THERMAL UNDERAMOUR PACKING ONEGAI SEAL (梱包シール) Ver: " + version;
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
                dt.Columns.Add("PONo");
                dt.Columns.Add("Customer");
                dt.Columns.Add("Product");
                dt.Columns.Add("Description");
                dt.Columns.Add("StyleNo");
                dt.Columns.Add("ColorCode");
                dt.Columns.Add("SizeName");
                dt.Columns.Add("Quantity");
                dt.Columns.Add("OrderNo");
                dt.Columns.Add("JAN");
                dt.Columns.Add("PCS");
                var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                var workbook = WorkbookFactory.Create(stream);
                var sheetName = "PackingList";
                var sheetAssortmentList = workbook.GetSheet(sheetName);
                if (sheetAssortmentList == null)
                {
                    MessageBox.Show($"Sheet with name not found: {sheetName}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string poNo = GetFirstPart(Path.GetFileName(filePath));
                string customer = string.Empty;
                string customerCD = string.Empty;
                string mySQL = $@"SELECT [X_CONSIGNEE_CD], [X_CONSIGNEE_NAME]
                     FROM [ENVIETNAMPO].[dbo].[_TBL_PO_H]
                     WHERE X_PO_NO = '{poNo}'";
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["SQLConnectionString"].ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(mySQL, conn))
                    {
                        using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                        {
                            DataTable dt_po = new DataTable();
                            adapter.Fill(dt_po);

                            if (dt_po.Rows.Count > 0)
                            {
                                customerCD = dt_po.Rows[0]["X_CONSIGNEE_CD"].ToString();
                                customer = dt_po.Rows[0]["X_CONSIGNEE_NAME"].ToString();
                            }
                        }
                    }
                }

                if (string.IsNullOrEmpty(customer))
                {
                    MessageBox.Show($"This PONo customer name not found: {poNo}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                for (int i = 1; i <= sheetAssortmentList.LastRowNum; i++)
                {
                    int quantity = (int)Math.Round(double.Parse(excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 6)));
                    if (customerCD == "15050" || customerCD == "9802")
                    {
                        customer = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 7);
                    }
                    DataRow dataRow = dt.NewRow();
                    dataRow["PONo"] = poNo;
                    dataRow["Customer"] = customer;
                    dataRow["Product"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 0);
                    dataRow["Description"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 1);
                    dataRow["StyleNo"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 3);
                    dataRow["ColorCode"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 4);
                    dataRow["SizeName"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 5);
                    dataRow["Quantity"] = quantity.ToString("N0");
                    dataRow["OrderNo"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 9);
                    dataRow["JAN"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 10);
                    dataRow["PCS"] = "1";
                    dt.Rows.Add(dataRow.ItemArray);
                }

                dgvData.DataSource = dt;
                st_count_1.Text = dgvData.Rows.Count.ToString();
                webView21.CoreWebView2.Navigate($@"{Path.GetDirectoryName(assembly.Location)}\Template\Default.html");

                // showNotification("SUCCESS", "Read excel file completed.");
                var resModel = new ImportResponseModel
                {
                    IsSuccess = true,
                    Message = "Read excel file completed.",
                    RecordsImported = dgvData.Rows.Count,
                    RecordsFailed = 0,
                    SourceName = Path.GetFileName(filePath),
                    ErrorDetail = string.Empty,
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
            ThermalUnderamour thermalUnderamour = new ThermalUnderamour();
            thermalUnderamour.ImagePath = imagePath;
            thermalUnderamour.PONo = row.Cells["PONo"].Value?.ToString() ?? string.Empty;
            thermalUnderamour.Customer = row.Cells["Customer"].Value?.ToString() ?? string.Empty;
            thermalUnderamour.Product = row.Cells["Product"].Value?.ToString() ?? string.Empty;
            thermalUnderamour.Description = row.Cells["Description"].Value?.ToString() ?? string.Empty;
            thermalUnderamour.StyleNo = row.Cells["StyleNo"].Value?.ToString() ?? string.Empty;
            thermalUnderamour.ColorCode = row.Cells["ColorCode"].Value?.ToString() ?? string.Empty;
            thermalUnderamour.SizeName = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
            thermalUnderamour.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
            thermalUnderamour.JAN = row.Cells["JAN"].Value?.ToString() ?? string.Empty;
            thermalUnderamour.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString());
            htmlContent = templateHTML.ThermalUAToWebview(thermalUnderamour);
            RenderHTMLToWebView(htmlContent);
        }

        private async void FrmPackingSealThermalUA_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.frmThermalUASize.Width != 0 && Properties.Settings.Default.frmThermalUASize.Height != 0)
            {
                this.Size = Properties.Settings.Default.frmThermalUASize;
            }
            if (Properties.Settings.Default.frmThermalUALocation.X != 0 && Properties.Settings.Default.frmThermalUALocation.Y != 0)
            {
                this.StartPosition = FormStartPosition.CenterScreen;
                this.Location = Properties.Settings.Default.frmThermalUALocation;
            }
            this.WindowState = Properties.Settings.Default.frmThermalUAState;

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
                ThermalUnderamour thermalUnderamour = new ThermalUnderamour();
                thermalUnderamour.ImagePath = imagePath;
                thermalUnderamour.PONo = row.Cells["PONo"].Value?.ToString() ?? string.Empty;
                thermalUnderamour.Customer = row.Cells["Customer"].Value?.ToString() ?? string.Empty;
                thermalUnderamour.Product = row.Cells["Product"].Value?.ToString() ?? string.Empty;
                thermalUnderamour.Description = row.Cells["Description"].Value?.ToString() ?? string.Empty;
                thermalUnderamour.StyleNo = row.Cells["StyleNo"].Value?.ToString() ?? string.Empty;
                thermalUnderamour.ColorCode = row.Cells["ColorCode"].Value?.ToString() ?? string.Empty;
                thermalUnderamour.SizeName = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
                thermalUnderamour.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
                thermalUnderamour.JAN = row.Cells["JAN"].Value?.ToString() ?? string.Empty;
                thermalUnderamour.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString());
                htmlContent = templateHTML.ThermalUAToWebview(thermalUnderamour);
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
                    string orderNo = dgvData.Rows[0].Cells["PONo"].Value?.ToString() ?? "";
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
                        Brand = "UNDERAMOUR",
                        PoNumber = orderNo,
                        PrinterName = printerName,
                        PrintTime = DateTime.Now,
                        MsgVal = "PDF printed successfully."
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
            string orderNo = string.Empty;
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
                    if (row.IsNewRow || int.Parse(row.Cells["pcs"].Value?.ToString()) == 0)
                    {
                        continue;
                    }

                    for (int i = 0; i < int.Parse(row.Cells["pcs"].Value.ToString()); i++)
                    {
                        ThermalUnderamour thermalUnderamour = new ThermalUnderamour();
                        thermalUnderamour.ImagePath = imagePath;
                        thermalUnderamour.PONo = row.Cells["PONo"].Value?.ToString() ?? string.Empty;
                        thermalUnderamour.Customer = row.Cells["Customer"].Value?.ToString() ?? string.Empty;
                        thermalUnderamour.Product = row.Cells["Product"].Value?.ToString() ?? string.Empty;
                        thermalUnderamour.Description = row.Cells["Description"].Value?.ToString() ?? string.Empty;
                        thermalUnderamour.StyleNo = row.Cells["StyleNo"].Value?.ToString() ?? string.Empty;
                        thermalUnderamour.ColorCode = row.Cells["ColorCode"].Value?.ToString() ?? string.Empty;
                        thermalUnderamour.SizeName = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
                        thermalUnderamour.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
                        thermalUnderamour.JAN = row.Cells["JAN"].Value?.ToString() ?? string.Empty;
                        thermalUnderamour.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString());
                        string tableContent = templateHTML.ThermalUAToPDF(thermalUnderamour);
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

                string htmlTemplate = File.ReadAllText(htmlPath);
                string fullHtml = htmlTemplate.Replace("[htmlContent]", tableBuilder.ToString());

                if (Common.Consts.SettingINI.OnlyPrint != "True")
                {
                    if (string.IsNullOrEmpty(Common.Consts.SettingINI.PDFOutputFolderPath))
                    {
                        Common.Consts.SettingINI.PDFOutputFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    }
                    pdfOutputPath = Path.Combine(Common.Consts.SettingINI.PDFOutputFolderPath, $"ONEGAI_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.pdf");
                }
                else
                {
                    pdfOutputPath = Path.Combine(Path.GetTempPath(), $"ONEGAI_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.pdf");
                }

                using (FileStream pdfDest = new FileStream(pdfOutputPath, FileMode.Create))
                {
                    PDFSetting pDFSetting = new PDFSetting();
                    var fontProvider = PDFSetting.CreateFontProviderWithCustomFonts();
                    var props = new ConverterProperties();
                    props.SetFontProvider(fontProvider);
                    props.SetCharset("utf-8");

                    HtmlConverter.ConvertToPdf(fullHtml, pdfDest, props);
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

                if (!Path.GetFileNameWithoutExtension(filePath).EndsWith("PackingList"))
                {
                    MessageBox.Show("File name must Ends with 'PackingList'", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cursor = Cursors.Default;
                return;
            }
            Cursor = Cursors.Default;
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
                if (!Path.GetFileNameWithoutExtension(originalPath).EndsWith("PackingList"))
                {
                    MessageBox.Show("File name must Ends with 'PackingList'", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

        private void FrmPackingSealThermalUA_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.frmThermalUAState = this.WindowState;
            if (this.WindowState == FormWindowState.Normal)
            {
                Properties.Settings.Default.frmThermalUALocation = this.Location;
                Properties.Settings.Default.frmThermalUASize = this.Size;
            }
            else
            {
                Properties.Settings.Default.frmThermalUALocation = this.RestoreBounds.Location;
                Properties.Settings.Default.frmThermalUASize = this.RestoreBounds.Size;
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

            if (dgvData.CurrentCell.ColumnIndex == dgvData.Columns["Quantity"].Index)
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

        private void FrmPackingSealtThermalUA_KeyDown(object sender, KeyEventArgs e)
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

        private static string GetFirstPart(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            string[] parts = input.Split('_');
            return parts.Length > 0 ? parts[0] : string.Empty;
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
                        $"PONo: {resModel.PoNumber}\n" +
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
