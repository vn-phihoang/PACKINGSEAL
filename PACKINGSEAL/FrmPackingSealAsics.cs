using iText.Html2pdf;
using Microsoft.VisualBasic.FileIO;
using PACKINGSEAL.Models;
using PACKINGSEAL.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Deployment.Application;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PACKINGSEAL
{
    public partial class FrmPackingSealAsics : Form
    {
        public string imagePath = string.Empty;
        private string htmlPath = string.Empty;
        private string pdfOutputPath = string.Empty;
        private Assembly assembly = Assembly.GetExecutingAssembly();

        public FrmPackingSealAsics(FilePathModel filePathModel)
        {
            InitializeComponent();
            string UserName = Environment.MachineName;
            lb_info_1.Text = "HC | Asics Packing Onegai Seal- " + UserName;
            string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
            var version = "";
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                version = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            this.Text = "ASICS PACKING ONEGAI SEAL (梱包シール) Ver: " + version;
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
                DataTable dt = new DataTable();
                dt = new DataTable();
                dt.Columns.Add("Customer");
                dt.Columns.Add("OrderNo");
                dt.Columns.Add("POLine");
                dt.Columns.Add("SKUNo");
                dt.Columns.Add("ItemCode");
                dt.Columns.Add("ProductName");
                dt.Columns.Add("SampleNo");
                dt.Columns.Add("ColorCode");
                dt.Columns.Add("SizeName");
                dt.Columns.Add("Quantity");
                dt.Columns.Add("KeyNo");
                dt.Columns.Add("ContractNo");
                dt.Columns.Add("Vendors");
                dt.Columns.Add("PCS");

                if (Path.GetExtension(filePath).ToLower().Equals(".usv"))
                {
                    using (TextFieldParser parser = new TextFieldParser(filePath))
                    {
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(",");
                        parser.HasFieldsEnclosedInQuotes = true;

                        while (!parser.EndOfData)
                        {
                            string[] fields = parser.ReadFields();

                            // Đảm bảo mảng có ít nhất 30 phần tử
                            if (fields.Length < 30)
                            {
                                Array.Resize(ref fields, 30);
                                for (int j = 0; j < fields.Length; j++)
                                {
                                    if (fields[j] == null)
                                    {
                                        fields[j] = string.Empty;
                                    }
                                }
                            }

                            if (fields.Length >= 10)
                            {
                                DataRow dataRow = dt.NewRow();
                                dataRow["Customer"] = fields[12].Trim().Replace("\"", "");
                                dataRow["OrderNo"] = fields[0].Trim().Replace("\"", "");
                                dataRow["PoLine"] = fields[22].Trim().Replace("\"", "");
                                dataRow["SKUNo"] = fields[21].Trim().Replace("\"", "");

                                string ItemCode = fields[1].Trim().Replace("\"", "");
                                if (ItemCode.Length > 4)
                                {
                                    ItemCode = ItemCode.Insert(4, " - ");
                                }
                                dataRow["ItemCode"] = ItemCode;

                                dataRow["ProductName"] = fields[2].Trim().Replace("\"", "");
                                dataRow["SampleNo"] = fields[6].Trim().Replace("\"", "");
                                dataRow["ColorCode"] = fields[17].Trim().Replace("\"", "");
                                dataRow["SizeName"] = fields[18].Trim().Replace("\"", "");
                                dataRow["Quantity"] = Convert.ToInt32(fields[5].Trim().Replace("\"", ""));
                                dataRow["KeyNo"] = fields[20].Trim().Replace("\"", "");
                                dataRow["ContractNo"] = fields[7].Trim().Replace("\"", "");
                                dataRow["Vendors"] = fields[10].Trim().Replace("\"", "");
                                dataRow["PCS"] = "1";

                                dt.Rows.Add(dataRow.ItemArray);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show($"File USV null.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Cursor = Cursors.Default;
                    return;
                }


                dgvData.DataSource = dt;
                st_count_1.Text = dgvData.Rows.Count.ToString();
                webView21.CoreWebView2.Navigate($@"{Path.GetDirectoryName(assembly.Location)}\Template\Default.html");

                //MessageBox.Show($"Read excel file completed.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //showNotification("SUCCESS", "Read excel file completed.");

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
            AsicsModel asicsModel = new AsicsModel();
            asicsModel.ImagePath = imagePath;
            asicsModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
            asicsModel.Customer = row.Cells["Customer"].Value?.ToString() ?? string.Empty;
            asicsModel.POLine = row.Cells["POLine"].Value?.ToString() ?? string.Empty;
            string SkuNo = row.Cells["SKUNo"].Value.ToString() ?? string.Empty;
            if (!SkuNo.StartsWith("-") && !string.IsNullOrEmpty(SkuNo))
            {
                SkuNo = "-" + SkuNo;
            }
            asicsModel.SKUNo = SkuNo;
            asicsModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
            asicsModel.SampleNo = row.Cells["SampleNo"].Value?.ToString() ?? string.Empty;
            asicsModel.ItemCode = row.Cells["ItemCode"].Value?.ToString() ?? string.Empty;
            string sizeName = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
            if (!sizeName.StartsWith(".") && !string.IsNullOrEmpty(sizeName))
            {
                sizeName = "." + sizeName;
            }
            asicsModel.SizeName = sizeName;
            string colorCode = row.Cells["ColorCode"].Value?.ToString() ?? string.Empty;
            if (!colorCode.StartsWith(".") && !string.IsNullOrEmpty(colorCode))
            {
                colorCode = "." + colorCode;
            }
            asicsModel.ColorCode = colorCode;
            asicsModel.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString());
            asicsModel.KeyNo = row.Cells["KeyNo"].Value?.ToString() ?? string.Empty;
            asicsModel.ContractNo = row.Cells["ContractNo"].Value?.ToString() ?? string.Empty;
            asicsModel.Vendors = row.Cells["Vendors"].Value?.ToString() ?? string.Empty;
            asicsModel.ProductName = row.Cells["ProductName"].Value?.ToString() ?? string.Empty;
            htmlContent = templateHTML.AsicsToWebview(asicsModel);

            RenderHTMLToWebView(htmlContent);
        }

        private async void FrmPackingSealAsics_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.frmAsicsSize.Width != 0 && Properties.Settings.Default.frmAsicsSize.Height != 0)
            {
                this.Size = Properties.Settings.Default.frmAsicsSize;
            }
            if (Properties.Settings.Default.frmAsicsLocation.X != 0 && Properties.Settings.Default.frmAsicsLocation.Y != 0)
            {
                this.StartPosition = FormStartPosition.CenterScreen;
                this.Location = Properties.Settings.Default.frmAsicsLocation;
            }
            this.WindowState = Properties.Settings.Default.frmAsicsState;

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
                AsicsModel asicsModel = new AsicsModel();
                asicsModel.ImagePath = imagePath;
                asicsModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
                asicsModel.Customer = row.Cells["Customer"].Value?.ToString() ?? string.Empty;
                asicsModel.POLine = row.Cells["POLine"].Value?.ToString() ?? string.Empty;
                string SkuNo = row.Cells["SKUNo"].Value.ToString() ?? string.Empty;
                if (!SkuNo.StartsWith("-") && !string.IsNullOrEmpty(SkuNo))
                {
                    SkuNo = "-" + SkuNo;
                }
                asicsModel.SKUNo = SkuNo;
                asicsModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
                asicsModel.SampleNo = row.Cells["SampleNo"].Value?.ToString() ?? string.Empty;
                asicsModel.ItemCode = row.Cells["ItemCode"].Value?.ToString() ?? string.Empty;
                string sizeName = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
                if (!sizeName.StartsWith(".") && !string.IsNullOrEmpty(sizeName))
                {
                    sizeName = "." + sizeName;
                }
                asicsModel.SizeName = sizeName;
                string colorCode = row.Cells["ColorCode"].Value?.ToString() ?? string.Empty;
                if (!colorCode.StartsWith(".") && !string.IsNullOrEmpty(colorCode))
                {
                    colorCode = "." + colorCode;
                }
                asicsModel.ColorCode = colorCode;
                asicsModel.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString());
                asicsModel.KeyNo = row.Cells["KeyNo"].Value?.ToString() ?? string.Empty;
                asicsModel.ContractNo = row.Cells["ContractNo"].Value?.ToString() ?? string.Empty;
                asicsModel.Vendors = row.Cells["Vendors"].Value?.ToString() ?? string.Empty;
                asicsModel.ProductName = row.Cells["ProductName"].Value?.ToString() ?? string.Empty;
                htmlContent = templateHTML.AsicsToWebview(asicsModel);
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
                    string orderNo = dgvData.Rows[0].Cells["OrderNo"].Value?.ToString() ?? "";
                    string printerName = Common.Consts.SettingINI.PrinterName;
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

                    //MessageBox.Show("PDF printed successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    /*System.Threading.Thread.Sleep(3000);
                    new Utils.PDFSetting().showNotification("SUCCESS", "PDF printed successfully.");*/

                    var resModel = new PrintResponseModel
                    {
                        Brand = "ASICS",
                        PoNumber = orderNo,
                        PrinterName = printerName,
                        PrintTime = DateTime.Now,
                        MsgVal = $"PONo {orderNo} has been successfully printed!",
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
                        AsicsModel asicsModel = new AsicsModel();
                        asicsModel.ImagePath = imagePath;
                        asicsModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
                        asicsModel.Customer = row.Cells["Customer"].Value?.ToString() ?? string.Empty;
                        asicsModel.POLine = row.Cells["POLine"].Value?.ToString() ?? string.Empty;
                        string SkuNo = row.Cells["SKUNo"].Value.ToString() ?? string.Empty;
                        if (!SkuNo.StartsWith("-") && !string.IsNullOrEmpty(SkuNo))
                        {
                            SkuNo = "-" + SkuNo;
                        }
                        asicsModel.SKUNo = SkuNo;
                        asicsModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
                        asicsModel.SampleNo = row.Cells["SampleNo"].Value?.ToString() ?? string.Empty;
                        asicsModel.ItemCode = row.Cells["ItemCode"].Value?.ToString() ?? string.Empty;
                        string sizeName = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
                        if (!sizeName.StartsWith(".") && !string.IsNullOrEmpty(sizeName))
                        {
                            sizeName = "." + sizeName;
                        }
                        asicsModel.SizeName = sizeName;
                        string colorCode = row.Cells["ColorCode"].Value?.ToString() ?? string.Empty;
                        if (!colorCode.StartsWith(".") && !string.IsNullOrEmpty(colorCode))
                        {
                            colorCode = "." + colorCode;
                        }
                        asicsModel.ColorCode = colorCode;
                        asicsModel.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString());
                        asicsModel.KeyNo = row.Cells["KeyNo"].Value?.ToString() ?? string.Empty;
                        asicsModel.ContractNo = row.Cells["ContractNo"].Value?.ToString() ?? string.Empty;
                        asicsModel.Vendors = row.Cells["Vendors"].Value?.ToString() ?? string.Empty;
                        asicsModel.ProductName = row.Cells["ProductName"].Value?.ToString() ?? string.Empty;
                        string tableContent = templateHTML.AsicsToPDF(asicsModel);
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
                    Filter = "USV Files|*.usv",
                    Title = "Select USV File (SKU ONEGAI)",
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

                if (fileExt != ".usv")
                {
                    MessageBox.Show("Please select USV file (SKU ONEGAI)", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Cursor = Cursors.Default;
                    return;
                }

                if (!fileName.StartsWith("SKU_ONEGAI"))
                {
                    MessageBox.Show("File name must start with 'SKU_ONEGAI'", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Cursor = Cursors.Default;
                    return;
                }

                string preparedFile = templateHTML.PrepareFile(filePath);
                textBoxFile.Text = filePath;

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

            if (files == null || files.Length == 0)
                return;

            string originalPath = files[0];
            textBoxFile.Text = originalPath + Environment.NewLine;

            string extension = Path.GetExtension(originalPath).ToLower();

            if (extension != ".usv")
            {
                MessageBox.Show("Please select USV file (.usv)", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!Path.GetFileName(originalPath).StartsWith("SKU_ONEGAI", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("File name must start with 'SKU_ONEGAI'", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (FileStream fs = new FileStream(originalPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    LoadDataGirdview(originalPath);
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show("Cannot read file: " + ex.Message);
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

        private bool isUpdating = false;

        private void dgvData_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (isUpdating)
            {
                return;
            }

            if (dgvData.Columns.Contains("Quantity") && e.ColumnIndex == dgvData.Columns["Quantity"].Index)
            {
                int rowIndex = e.RowIndex;
                if (rowIndex < 0 || rowIndex >= dgvData.Rows.Count)
                {
                    return;
                }
                string quantityStr = dgvData.Rows[rowIndex].Cells["Quantity"].Value?.ToString() ?? "1";
                quantityStr = quantityStr.Replace(",", "").Trim();

                if (!int.TryParse(quantityStr.Split('/')[0], out int quantity) || quantity <= 0)
                {
                    MessageBox.Show("Quantity must be greater than 0.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    dgvData.Rows[rowIndex].Cells["Quantity"].Value = "1";
                    return;
                }
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

        private void FrmPackingSealAsics_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.frmAsicsState = this.WindowState;
            if (this.WindowState == FormWindowState.Normal)
            {
                Properties.Settings.Default.frmAsicsLocation = this.Location;
                Properties.Settings.Default.frmAsicsSize = this.Size;
            }
            else
            {
                Properties.Settings.Default.frmAsicsLocation = this.RestoreBounds.Location;
                Properties.Settings.Default.frmAsicsSize = this.RestoreBounds.Size;
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

        private void FrmPackingSealAsics_KeyDown(object sender, KeyEventArgs e)
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
                    notifyIcon1.Icon = this.Icon;
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

        private void buttonHelp_Click(object sender, EventArgs e)
        {
            try
            {
                string assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string sourceFile = Path.Combine(assemblyPath, @"Resources\HelpAsics.pdf");
                string tempFile = Path.Combine(Path.GetTempPath(), "HelpAsics.pdf");
                File.Copy(sourceFile, tempFile, true);
                Process.Start(new ProcessStartInfo(tempFile)
                {
                    UseShellExecute = true
                });

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.ToString());
            }
        }
    }
}
