using iText.Html2pdf;
using NPOI.SS.UserModel;
using PACKINGSEAL.Models;
using PACKINGSEAL.Utils;
using System;
using System.Collections.Generic;
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
    public partial class FrmPackingSealUniqlo : Form
    {
        public string imagePath = string.Empty;
        private string htmlPath = string.Empty;
        private string pdfOutputPath = string.Empty;
        private Assembly assembly = Assembly.GetExecutingAssembly();

        public FrmPackingSealUniqlo(FilePathModel filePathModel)
        {
            InitializeComponent();
            string UserName = Environment.MachineName;
            lb_info_1.Text = "HC | Uniqlo Packing Onegai Seal- " + UserName;
            string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
            var version = "";
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                version = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            this.Text = "UNIQLO PACKING ONEGAI SEAL (梱包シール) Ver: " + version;
            st_count_1.Text = "0";
            textBoxFile.Clear();
            cleanTempPDFFiles();
            imagePath = filePathModel.ImagePath;
            htmlPath = filePathModel.HtmlPath;
            if (webView21.CoreWebView2 != null)
            {
                //webView21.CoreWebView2.NavigateToString("<html><body></body></html>");
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
                dt.Columns.Add("MaterialCode");
                dt.Columns.Add("DeliveryDestination");
                dt.Columns.Add("PONo");
                dt.Columns.Add("OrderNo");
                dt.Columns.Add("SampleNo");
                dt.Columns.Add("ColorName");
                dt.Columns.Add("ColorCode");
                dt.Columns.Add("SizeName");
                dt.Columns.Add("QuantityInBox");
                dt.Columns.Add("Quantity");
                dt.Columns.Add("KeyNo");
                dt.Columns.Add("ContractNo");
                dt.Columns.Add("PCS");
                var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                var workbook = WorkbookFactory.Create(stream);
                var sheetName = "アソート表";
                var sheetAssortmentList = workbook.GetSheet(sheetName);
                if (sheetAssortmentList == null)
                {
                    MessageBox.Show($"Sheet with name not found: {sheetName}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                sheetName = "パッキングリスト";
                var sheetPackingList = workbook.GetSheet(sheetName);
                if (sheetPackingList == null)
                {
                    MessageBox.Show($"Sheet with name not found: {sheetName}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                //int actualRowCount = GetActualRowCount(sheetAssortmentList);
                string orderNo = excelUtil.GetCellValueWithMerge(sheetAssortmentList, 2, 2);
                string deliveryDestination = Regex.Replace(excelUtil.GetCellValueWithMerge(sheetAssortmentList, 3, 2), @"^.*?： ", "");
                string poNo = excelUtil.GetCellValueWithMerge(sheetAssortmentList, 7, 2);
                string sampleNo = excelUtil.GetCellValueWithMerge(sheetAssortmentList, 7, 11).Substring(0, 10);
                string contractNo = excelUtil.GetCellValueWithMerge(sheetPackingList, 6, 0);

                for (int i = 0; i <= sheetAssortmentList.LastRowNum; i++)
                {
                    if (!string.IsNullOrEmpty(excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 1)) && excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 1) == "ユニクロ　アソート表")
                    {
                        i += 10; // Skip the next 10 rows
                    }
                    if (string.IsNullOrEmpty(excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 2)))
                    {
                        continue; // Skip rows without MaterialCode
                    }

                    int quantity = (int)Math.Round(double.Parse(excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 13)));
                    string cellValue = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 13);

                    int totalBoxes = 1;

                    if (excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 2).StartsWith("HT"))
                    {
                        if (excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 2).EndsWith("140") || excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 2).EndsWith("160"))
                        {
                            totalBoxes = quantity / 1000;
                            if (quantity % 1000 != 0)
                            {
                                totalBoxes += 1;
                            }
                        }
                        else
                        {
                            totalBoxes = quantity / 2000;
                            if (quantity % 2000 != 0)
                            {
                                totalBoxes += 1;
                            }
                        }
                    }
                    else if (excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 2).StartsWith("WT"))
                    {
                        totalBoxes = quantity / 1000;
                        if (quantity % 1000 != 0)
                        {
                            totalBoxes += 1;
                        }
                    }
                    for (int boxIndex = 0; boxIndex < totalBoxes; boxIndex++)
                    {
                        DataRow dataRow = dt.NewRow();
                        dataRow["OrderNo"] = orderNo;
                        dataRow["DeliveryDestination"] = deliveryDestination;
                        dataRow["PONo"] = poNo;
                        dataRow["SampleNo"] = sampleNo;
                        dataRow["MaterialCode"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 2);
                        dataRow["SizeName"] = Regex.Replace(excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 9), @"^.*?[：:]\s*", "");
                        dataRow["ColorName"] = Regex.Replace(excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 11), @"^.*?[：:]\s*", "");
                        Match match = Regex.Match(excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 11), @"^(.*?)\s*:");
                        dataRow["ColorCode"] = match.Groups[1].Value;
                        int boxQuantity = 0;
                        if (excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 2).StartsWith("HT"))
                        {
                            if (excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 2).EndsWith("140") || excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 2).EndsWith("160"))
                            {
                                boxQuantity = (boxIndex == totalBoxes - 1) ? quantity - 1000 * (totalBoxes - 1) : 1000;
                            }
                            else
                            {
                                boxQuantity = (boxIndex == totalBoxes - 1) ? quantity - 2000 * (totalBoxes - 1) : 2000;
                            }
                            dataRow["QuantityInBox"] = boxQuantity.ToString("N0");
                        }
                        else if (excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 2).StartsWith("WT"))
                        {
                            boxQuantity = (boxIndex == totalBoxes - 1) ? quantity - 1000 * (totalBoxes - 1) : 1000;
                            dataRow["QuantityInBox"] = boxQuantity.ToString("N0");
                        }
                        else
                        {
                            dataRow["QuantityInBox"] = quantity.ToString("N0");
                        }
                        dataRow["Quantity"] = quantity.ToString("N0");
                        dataRow["KeyNo"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 15);
                        dataRow["ContractNo"] = contractNo;
                        dataRow["PCS"] = "1";
                        dt.Rows.Add(dataRow.ItemArray);
                    }
                }

                dgvData.DataSource = dt;
                st_count_1.Text = dgvData.Rows.Count.ToString();
                //webView21.CoreWebView2.NavigateToString("<html><body></body></html>");
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
            UniqloModel uniqloModel = new UniqloModel();
            uniqloModel.ImagePath = imagePath;
            uniqloModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
            uniqloModel.DeliveryDestination = row.Cells["DeliveryDestination"].Value?.ToString() ?? string.Empty;
            uniqloModel.PoNo = row.Cells["PONo"].Value?.ToString() ?? string.Empty;
            uniqloModel.SampleNo = row.Cells["SampleNo"].Value?.ToString() ?? string.Empty;
            uniqloModel.MaterialCode = row.Cells["MaterialCode"].Value?.ToString() ?? string.Empty;
            uniqloModel.SizeName = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
            uniqloModel.ColorName = row.Cells["ColorName"].Value?.ToString() ?? string.Empty;
            uniqloModel.ColorCode = row.Cells["ColorCode"].Value?.ToString() ?? string.Empty;
            uniqloModel.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString());
            uniqloModel.KeyNo = row.Cells["KeyNo"].Value?.ToString() ?? string.Empty;
            uniqloModel.ContractNo = row.Cells["ContractNo"].Value?.ToString() ?? string.Empty;
            uniqloModel.QuantityInBox = string.IsNullOrEmpty(row.Cells["QuantityInBox"].Value.ToString()) ? 0 : double.Parse(row.Cells["QuantityInBox"].Value.ToString());
            htmlContent = templateHTML.UQToWebview(uniqloModel);

            RenderHTMLToWebView(htmlContent);
        }

        private async void FrmPackingSealUniqlo_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.frmUniqloSize.Width != 0 && Properties.Settings.Default.frmUniqloSize.Height != 0)
            {
                this.Size = Properties.Settings.Default.frmUniqloSize;
            }
            if (Properties.Settings.Default.frmUniqloLocation.X != 0 && Properties.Settings.Default.frmUniqloLocation.Y != 0)
            {
                this.StartPosition = FormStartPosition.CenterScreen;
                this.Location = Properties.Settings.Default.frmUniqloLocation;
            }
            this.WindowState = Properties.Settings.Default.frmUniqloState;

            //await webView21.EnsureCoreWebView2Async();
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
                UniqloModel uniqloModel = new UniqloModel();
                uniqloModel.ImagePath = imagePath;
                uniqloModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
                uniqloModel.DeliveryDestination = row.Cells["DeliveryDestination"].Value?.ToString() ?? string.Empty;
                uniqloModel.PoNo = row.Cells["PONo"].Value?.ToString() ?? string.Empty;
                uniqloModel.SampleNo = row.Cells["SampleNo"].Value?.ToString() ?? string.Empty;
                uniqloModel.MaterialCode = row.Cells["MaterialCode"].Value?.ToString() ?? string.Empty;
                uniqloModel.SizeName = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
                uniqloModel.ColorName = row.Cells["ColorName"].Value?.ToString() ?? string.Empty;
                uniqloModel.ColorCode = row.Cells["ColorCode"].Value?.ToString() ?? string.Empty;
                uniqloModel.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : int.Parse(row.Cells["Quantity"].Value.ToString());
                uniqloModel.KeyNo = row.Cells["KeyNo"].Value?.ToString() ?? string.Empty;
                uniqloModel.ContractNo = row.Cells["ContractNo"].Value?.ToString() ?? string.Empty;
                uniqloModel.QuantityInBox = string.IsNullOrEmpty(row.Cells["QuantityInBox"].Value.ToString()) ? 0 : double.Parse(row.Cells["QuantityInBox"].Value.ToString());
                htmlContent = templateHTML.UQToWebview(uniqloModel);
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
                    string materialCode = dgvData.Rows[0].Cells["MaterialCode"].Value?.ToString() ?? "";
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
                        Brand = "UNIQLO",
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
                        UniqloModel uniqloModel = new UniqloModel();
                        uniqloModel.ImagePath = imagePath;
                        uniqloModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
                        uniqloModel.DeliveryDestination = row.Cells["DeliveryDestination"].Value?.ToString() ?? string.Empty;
                        uniqloModel.PoNo = row.Cells["PONo"].Value?.ToString() ?? string.Empty;
                        uniqloModel.SampleNo = row.Cells["SampleNo"].Value?.ToString() ?? string.Empty;
                        uniqloModel.MaterialCode = row.Cells["MaterialCode"].Value?.ToString() ?? string.Empty;
                        uniqloModel.SizeName = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
                        uniqloModel.ColorName = row.Cells["ColorName"].Value?.ToString() ?? string.Empty;
                        uniqloModel.ColorCode = row.Cells["ColorCode"].Value?.ToString() ?? string.Empty;
                        uniqloModel.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString());
                        uniqloModel.KeyNo = row.Cells["KeyNo"].Value?.ToString() ?? string.Empty;
                        uniqloModel.ContractNo = row.Cells["ContractNo"].Value?.ToString() ?? string.Empty;
                        uniqloModel.QuantityInBox = string.IsNullOrEmpty(row.Cells["QuantityInBox"].Value.ToString()) ? 0 : double.Parse(row.Cells["QuantityInBox"].Value.ToString());
                        string tableContent = templateHTML.UQToPDF(uniqloModel);
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

                if (!fileName.StartsWith("生産指図書") && !fileName.StartsWith("ProductionInstructions"))
                {
                    MessageBox.Show("File name must start with '生産指図書' Or 'ProductionInstructions'", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                if (!Path.GetFileName(originalPath).StartsWith("生産指図書") && !Path.GetFileName(originalPath).StartsWith("ProductionInstructions"))
                {
                    MessageBox.Show("File name must start with '生産指図書' OR 'ProductionInstructions'", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                string materialCode = dgvData.Rows[rowIndex].Cells["MaterialCode"].Value?.ToString() ?? "";
                string quantityStr = dgvData.Rows[rowIndex].Cells["Quantity"].Value?.ToString() ?? "1";
                quantityStr = quantityStr.Replace(",", "").Trim();

                if (!int.TryParse(quantityStr.Split('/')[0], out int quantity) || quantity <= 0)
                {
                    MessageBox.Show("Quantity must be greater than 0.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    dgvData.Rows[rowIndex].Cells["Quantity"].Value = "1";
                    dgvData.Rows[rowIndex].Cells["QuantityInBox"].Value = "1";
                    return;
                }

                int boxSize = 0;

                if (materialCode.StartsWith("HT"))
                {
                    boxSize = 2000;
                    if (materialCode.EndsWith("140") || materialCode.EndsWith("160"))
                    {
                        boxSize = 1000;
                    }
                }
                else if (materialCode.StartsWith("WT"))
                {
                    boxSize = 1000;
                }

                if (boxSize == 0)
                {
                    dgvData.Rows[rowIndex].Cells["Quantity"].Value = quantity.ToString("N0");
                    dgvData.Rows[rowIndex].Cells["QuantityInBox"].Value = quantity.ToString("N0");
                    return;
                }

                // Tính số lượng box
                List<int> boxQuantities = new List<int>();
                int fullBoxes = quantity / boxSize;
                int lastBoxQty = quantity % boxSize;

                for (int i = 0; i < fullBoxes; i++)
                {
                    boxQuantities.Add(boxSize);
                }
                if (lastBoxQty > 0)
                {
                    boxQuantities.Add(lastBoxQty);
                }

                string sizeName = dgvData.Rows[rowIndex].Cells["SizeName"].Value?.ToString() ?? "";
                string colorCode = dgvData.Rows[rowIndex].Cells["ColorCode"].Value?.ToString() ?? "";
                string orderNo = dgvData.Rows[rowIndex].Cells["OrderNo"].Value?.ToString() ?? "";

                DataTable dt = dgvData.DataSource as DataTable;
                if (dt == null)
                {
                    return;
                }

                isUpdating = true;

                // Xác định dòng gốc trong DataTable
                DataRow currentRow = ((DataRowView)dgvData.Rows[rowIndex].DataBoundItem).Row;
                int baseIndex = dt.Rows.IndexOf(currentRow);

                // Xóa các dòng cùng loại (trừ dòng gốc)
                var rowsToDelete = dt.AsEnumerable()
                    .Where((r, idx) =>
                        r != currentRow &&
                        r.Field<string>("MaterialCode") == materialCode &&
                        r.Field<string>("SizeName") == sizeName &&
                        r.Field<string>("ColorCode") == colorCode &&
                        r.Field<string>("OrderNo") == orderNo
                    ).ToList();

                foreach (var r in rowsToDelete)
                {
                    dt.Rows.Remove(r);
                }

                dgvData.Rows[rowIndex].Cells["PCS"].Value = "1";
                dgvData.Rows[rowIndex].Cells["QuantityInBox"].Value = boxQuantities[0].ToString("N0");
                dgvData.Rows[rowIndex].Cells["Quantity"].Value = boxQuantities[0].ToString("N0");

                for (int i = 1; i < boxQuantities.Count; i++)
                {
                    DataRow newRow = dt.NewRow();
                    newRow["OrderNo"] = orderNo;
                    newRow["DeliveryDestination"] = dgvData.Rows[rowIndex].Cells["DeliveryDestination"].Value?.ToString() ?? "";
                    newRow["PONo"] = dgvData.Rows[rowIndex].Cells["PONo"].Value?.ToString() ?? "";
                    newRow["SampleNo"] = dgvData.Rows[rowIndex].Cells["SampleNo"].Value?.ToString() ?? "";
                    newRow["MaterialCode"] = materialCode;
                    newRow["SizeName"] = sizeName;
                    newRow["ColorName"] = dgvData.Rows[rowIndex].Cells["ColorName"].Value?.ToString() ?? "";
                    newRow["ColorCode"] = colorCode;
                    newRow["KeyNo"] = dgvData.Rows[rowIndex].Cells["KeyNo"].Value?.ToString() ?? "";
                    newRow["ContractNo"] = dgvData.Rows[rowIndex].Cells["ContractNo"].Value?.ToString() ?? "";
                    newRow["PCS"] = "1";
                    newRow["QuantityInBox"] = boxQuantities[i].ToString("N0");
                    newRow["Quantity"] = boxQuantities[i].ToString("N0");

                    dt.Rows.InsertAt(newRow, baseIndex + i);
                }
                st_count_1.Text = dgvData.Rows.Count.ToString();
                isUpdating = false;
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

        private void FrmPackingSealUniqlo_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.frmUniqloState = this.WindowState;
            if (this.WindowState == FormWindowState.Normal)
            {
                Properties.Settings.Default.frmUniqloLocation = this.Location;
                Properties.Settings.Default.frmUniqloSize = this.Size;
            }
            else
            {
                Properties.Settings.Default.frmUniqloLocation = this.RestoreBounds.Location;
                Properties.Settings.Default.frmUniqloSize = this.RestoreBounds.Size;
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

        private void FrmPackingSealUniqlo_KeyDown(object sender, KeyEventArgs e)
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
    }
}
