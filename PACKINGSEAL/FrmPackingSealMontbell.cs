using iText.Html2pdf;
using Microsoft.VisualBasic.FileIO;
using NPOI.SS.UserModel;
using PACKINGSEAL.Models;
using PACKINGSEAL.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Deployment.Application;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PACKINGSEAL
{
    public partial class FrmPackingSealMontbell : Form
    {
        public string imagePath = string.Empty;
        private string htmlPath = string.Empty;
        private string pdfOutputPath = string.Empty;
        private string folderPath = string.Empty;
        private DataTable dataTable = new DataTable();
        private string OrderNoValue = string.Empty;
        private Assembly assembly = Assembly.GetExecutingAssembly();
        private int typeProduct = 1; // 1: Tag Label, 2: Care Label
        private DataTable sizeTable = new DataTable();
        private DataTable originalData;

        public FrmPackingSealMontbell(FilePathModel filePathModel)
        {
            InitializeComponent();
            string UserName = Environment.MachineName;
            lb_info_1.Text = "HC | Montbell Packing Onegai Seal- " + UserName;
            string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
            var version = "";
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                version = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            this.Text = "MONTBELL PACKING ONEGAI SEAL (梱包シール) Ver: " + version;
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
            LoadComboboxSize();
        }

        private void LoadComboboxSize()
        {
            sizeTable = new DataTable("SizeDefinition");
            sizeTable.Columns.Add("SizeCode", typeof(string));
            sizeTable.Columns.Add("Quantity", typeof(int));
            sizeTable.Rows.Add("", 0);
            sizeTable.Rows.Add("80x62.5", 800);
            sizeTable.Rows.Add("58x88.5", 3200);
            sizeTable.Rows.Add("45x68", 5000);

            comboBoxSize.DataSource = sizeTable;
            comboBoxSize.DisplayMember = "SizeCode";
            comboBoxSize.ValueMember = "Quantity";

            comboBoxSize.DrawMode = DrawMode.OwnerDrawFixed;
            comboBoxSize.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBoxSize.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        private void comboBoxSize_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
            {
                return;
            }

            ComboBox combo = sender as ComboBox;
            DataRowView row = (DataRowView)combo.Items[e.Index];
            string subkey = row["Quantity"].ToString();
            string value = row["SizeCode"].ToString();

            e.DrawBackground();
            e.Graphics.DrawString(value, e.Font, Brushes.Black, e.Bounds.Left, e.Bounds.Top);
            e.Graphics.DrawString(subkey, e.Font, Brushes.Black, e.Bounds.Right - e.Graphics.MeasureString(subkey, e.Font).Width, e.Bounds.Top);
            e.DrawFocusRectangle();
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

        private void showNotification(string type, string message)
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

        private void LoadDataGirdview(string filePath, bool isParentFolder)
        {
            try
            {
                ExcelUtil excelUtil = new ExcelUtil();
                string extension = Path.GetExtension(filePath).ToLower();
                if (Path.GetExtension(filePath).ToLower().Equals(".usv"))
                {
                    if (isParentFolder == false)
                    {
                        CreateDataTable();
                        OrderNoValue = string.Empty;
                    }
                    using (TextFieldParser parser = new TextFieldParser(filePath))
                    {
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(",");

                        if (!parser.EndOfData)
                        {
                            string[] fields = parser.ReadFields();

                            if (fields.Length > 50)
                            {
                                OrderNoValue = string.IsNullOrWhiteSpace(fields[50])
                                    ? string.Empty
                                    : fields[50].Trim().Replace("\"", "");
                            }
                            else
                            {
                                MessageBox.Show($"File USV not enough columns in first row.",
                                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                Cursor = Cursors.Default;
                                return;
                            }
                        }
                        else
                        {
                            MessageBox.Show($"File USV null.",
                                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Cursor = Cursors.Default;
                            return;
                        }
                    }

                    foreach (DataRow row in dataTable.Rows)
                    {
                        row["OrderNo"] = OrderNoValue;
                    }
                }
                else
                {
                    CreateDataTable();
                    var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                    var workbook = WorkbookFactory.Create(stream);
                    var sheetName = "PackingList";
                    var sheet = workbook.GetSheet(sheetName);
                    if (sheet == null)
                    {
                        MessageBox.Show($"Sheet with name not found: {sheetName}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    Dictionary<string, int> quantityTotal = new Dictionary<string, int>();
                    for (int i = 1; i <= sheet.LastRowNum; i++)
                    {
                        string KeyProductItemCode = excelUtil.GetCellValueWithMerge(sheet, i, 0) + excelUtil.GetCellValueWithMerge(sheet, i, 4);
                        string styleNo = excelUtil.GetCellValueWithMerge(sheet, i, 4);
                        int quantity = int.Parse(excelUtil.GetCellValueWithMerge(sheet, i, 7));
                        if (string.IsNullOrEmpty(OrderNoValue))
                        {
                            AppCommonModule appCommonModule = new AppCommonModule();
                            OrderNoValue = appCommonModule.GetXNote1(excelUtil.GetCellValueWithMerge(sheet, i, 2));
                        }
                        if (quantityTotal.ContainsKey(KeyProductItemCode))
                        {
                            quantityTotal[KeyProductItemCode] += quantity;
                        }
                        else
                        {
                            quantityTotal[KeyProductItemCode] = quantity;
                        }
                        if (typeProduct == 1)
                        {
                            /*int totalBoxes = quantity / 3200;
                            if (quantity % 3200 != 0)
                            {
                                totalBoxes += 1;
                            }

                            for (int boxIndex = 0; boxIndex < totalBoxes; boxIndex++)
                            {
                                DataRow dataRow = dataTable.NewRow();

                                string line = excelUtil.GetCellValueWithMerge(sheet, i, 1);
                                string[] parts = line.Split(' ');
                                dataRow["MaterialCode"] = parts.Length > 1 ? parts[0] : "";

                                dataRow["KeyProductItemCode"] = KeyProductItemCode;
                                dataRow["PONo"] = excelUtil.GetCellValueWithMerge(sheet, i, 2);
                                dataRow["StyleNo"] = styleNo;
                                dataRow["ItemCode"] = excelUtil.GetCellValueWithMerge(sheet, i, 0);
                                dataRow["OrderNo"] = OrderNoValue;
                                int total = Convert.ToInt32(quantityTotal[KeyProductItemCode]);
                                dataRow["Quantity"] = total.ToString("N0");
                                dataRow["ColorName"] = excelUtil.GetCellValueWithMerge(sheet, i, 5);
                                dataRow["SizeName"] = excelUtil.GetCellValueWithMerge(sheet, i, 6);
                                int sizeQuantityStr = Convert.ToInt32(excelUtil.GetCellValueWithMerge(sheet, i, 7));
                                dataRow["SizeQuantity"] = sizeQuantityStr.ToString("N0");

                                // Tính số lượng trong box hiện tại
                                int boxQuantity;

                                if (boxIndex == totalBoxes - 1)
                                {
                                    boxQuantity = quantity - 3200 * (totalBoxes - 1);
                                }
                                else
                                {
                                    boxQuantity = 3200;
                                }

                                dataRow["QuantityInBox"] = boxQuantity.ToString("N0");
                                dataRow["PCS"] = "1";

                                dataTable.Rows.Add(dataRow);
                            }*/
                            DataRow dataRow = dataTable.NewRow();

                            string line = excelUtil.GetCellValueWithMerge(sheet, i, 1);
                            string[] parts = line.Split(' ');
                            dataRow["MaterialCode"] = parts.Length > 1 ? parts[0] : "";

                            dataRow["KeyProductItemCode"] = KeyProductItemCode;
                            dataRow["PONo"] = excelUtil.GetCellValueWithMerge(sheet, i, 2);
                            dataRow["StyleNo"] = styleNo;
                            dataRow["ItemCode"] = excelUtil.GetCellValueWithMerge(sheet, i, 0);
                            dataRow["OrderNo"] = OrderNoValue;
                            int total = Convert.ToInt32(quantityTotal[KeyProductItemCode]);
                            dataRow["Quantity"] = total.ToString("N0");
                            dataRow["ColorName"] = excelUtil.GetCellValueWithMerge(sheet, i, 5);
                            dataRow["SizeName"] = excelUtil.GetCellValueWithMerge(sheet, i, 6);
                            int sizeQuantityStr = Convert.ToInt32(excelUtil.GetCellValueWithMerge(sheet, i, 7));
                            dataRow["SizeQuantity"] = sizeQuantityStr.ToString("N0");
                            dataRow["Key"] = KeyProductItemCode + "_" + excelUtil.GetCellValueWithMerge(sheet, i, 5) + "_" + excelUtil.GetCellValueWithMerge(sheet, i, 6);
                            dataRow["QuantityInBox"] = sizeQuantityStr.ToString("N0");
                            dataRow["PCS"] = "1";

                            dataTable.Rows.Add(dataRow);
                        }
                        else
                        {
                            DataRow dataRow = dataTable.NewRow();

                            string line = excelUtil.GetCellValueWithMerge(sheet, i, 1);
                            string[] parts = line.Split(' ');
                            dataRow["MaterialCode"] = parts.Length > 1 ? parts[0] : "";

                            dataRow["KeyProductItemCode"] = KeyProductItemCode;
                            dataRow["PONo"] = excelUtil.GetCellValueWithMerge(sheet, i, 2);
                            dataRow["StyleNo"] = styleNo;
                            dataRow["ItemCode"] = excelUtil.GetCellValueWithMerge(sheet, i, 0);
                            dataRow["OrderNo"] = OrderNoValue;
                            int total = Convert.ToInt32(quantityTotal[KeyProductItemCode]);
                            dataRow["Quantity"] = total.ToString("N0");
                            dataRow["ColorName"] = excelUtil.GetCellValueWithMerge(sheet, i, 5);
                            dataRow["SizeName"] = excelUtil.GetCellValueWithMerge(sheet, i, 6);
                            int sizeQuantityStr = Convert.ToInt32(excelUtil.GetCellValueWithMerge(sheet, i, 7));
                            dataRow["SizeQuantity"] = sizeQuantityStr.ToString("N0");
                            dataRow["QuantityInBox"] = "0";
                            dataRow["PCS"] = "1";

                            dataTable.Rows.Add(dataRow);
                        }
                    }

                    foreach (DataRow row in dataTable.Rows)
                    {
                        string keyProductItemCode = row["KeyProductItemCode"].ToString();
                        if (quantityTotal.ContainsKey(keyProductItemCode))
                        {
                            row["Quantity"] = ((int)quantityTotal[keyProductItemCode]).ToString("N0");
                        }
                    }
                }

                dgvData.DataSource = dataTable;
                originalData = dataTable.Copy();
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
            MontbellModel montbellModel = new MontbellModel();
            montbellModel.ImagePath = imagePath;
            montbellModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
            montbellModel.PoNo = row.Cells["PONo"].Value?.ToString() ?? string.Empty;
            montbellModel.ItemCode = row.Cells["ItemCode"].Value?.ToString() ?? string.Empty;
            montbellModel.MaterialCode = row.Cells["MaterialCode"].Value?.ToString() ?? string.Empty;
            montbellModel.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString());
            montbellModel.StyleNo = row.Cells["StyleNo"].Value?.ToString() ?? string.Empty;
            montbellModel.Color = row.Cells["ColorName"].Value?.ToString() ?? string.Empty;
            montbellModel.Size = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
            montbellModel.SizeQuantity = string.IsNullOrEmpty(row.Cells["SizeQuantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["SizeQuantity"].Value.ToString());
            montbellModel.QuantityInBox = string.IsNullOrEmpty(row.Cells["QuantityInBox"].Value.ToString()) ? 0 : double.Parse(row.Cells["QuantityInBox"].Value.ToString());
            montbellModel.TypeProduct = typeProduct;
            htmlContent = templateHTML.MonbellToWebview(montbellModel);

            RenderHTMLToWebView(htmlContent);
        }

        private async void FrmPackingSealMontbell_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.frmMontbellSize.Width != 0 && Properties.Settings.Default.frmMontbellSize.Height != 0)
            {
                this.Size = Properties.Settings.Default.frmMontbellSize;
            }
            if (Properties.Settings.Default.frmMontbellLocation.X != 0 && Properties.Settings.Default.frmMontbellLocation.Y != 0)
            {
                this.StartPosition = FormStartPosition.CenterScreen;
                this.Location = Properties.Settings.Default.frmMontbellLocation;
            }
            this.WindowState = Properties.Settings.Default.frmMontbellState;

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
                MontbellModel montbellModel = new MontbellModel();
                montbellModel.ImagePath = imagePath;
                montbellModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
                montbellModel.PoNo = row.Cells["PONo"].Value?.ToString() ?? string.Empty;
                montbellModel.ItemCode = row.Cells["ItemCode"].Value?.ToString() ?? string.Empty;
                montbellModel.MaterialCode = row.Cells["MaterialCode"].Value?.ToString() ?? string.Empty;
                montbellModel.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : int.Parse(row.Cells["Quantity"].Value.ToString());
                montbellModel.StyleNo = row.Cells["StyleNo"].Value?.ToString() ?? string.Empty;
                montbellModel.Color = row.Cells["ColorName"].Value?.ToString() ?? string.Empty;
                montbellModel.Size = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
                montbellModel.SizeQuantity = string.IsNullOrEmpty(row.Cells["SizeQuantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["SizeQuantity"].Value.ToString());
                montbellModel.QuantityInBox = string.IsNullOrEmpty(row.Cells["QuantityInBox"].Value.ToString()) ? 0 : double.Parse(row.Cells["QuantityInBox"].Value.ToString());
                montbellModel.TypeProduct = typeProduct;
                htmlContent = templateHTML.MonbellToWebview(montbellModel);
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
                    string orderNo = dgvData.Rows[0].Cells["PONo"].Value?.ToString() ?? "";
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
                    System.Threading.Thread.Sleep(3000);
                    //new Utils.PDFSetting().showNotification("SUCCESS", "PDF printed successfully.");

                    var resModel = new PrintResponseModel
                    {
                        Brand = "MONTBELL",
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
                        MontbellModel montbellModel = new MontbellModel();
                        montbellModel.ImagePath = imagePath;
                        montbellModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
                        montbellModel.PoNo = row.Cells["PONo"].Value?.ToString() ?? string.Empty;
                        montbellModel.ItemCode = row.Cells["ItemCode"].Value?.ToString() ?? string.Empty;
                        montbellModel.MaterialCode = row.Cells["MaterialCode"].Value?.ToString() ?? string.Empty;
                        montbellModel.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString());
                        montbellModel.StyleNo = row.Cells["StyleNo"].Value?.ToString() ?? string.Empty;
                        montbellModel.Color = row.Cells["ColorName"].Value?.ToString() ?? string.Empty;
                        montbellModel.Size = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
                        montbellModel.SizeQuantity = string.IsNullOrEmpty(row.Cells["SizeQuantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["SizeQuantity"].Value.ToString());
                        montbellModel.QuantityInBox = string.IsNullOrEmpty(row.Cells["QuantityInBox"].Value.ToString()) ? 0 : double.Parse(row.Cells["QuantityInBox"].Value.ToString());
                        montbellModel.TypeProduct = typeProduct;
                        string tableContent = templateHTML.MonbellToPDF(montbellModel);
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
                string nameWithoutExt = string.Empty;
                bool isParentFolder = false;
                HtmlResponse templateHTML = new HtmlResponse();
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "USV and Excel Files|*.usv;*.xls;*.xlsx",
                    Title = "Select File",
                    InitialDirectory = Common.Consts.SettingINI.ImportFolderPath,
                    Multiselect = true
                };
                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    Cursor = Cursors.Default;
                    return;
                }
                textBoxFile.Text = string.Empty;
                string[] selectedFiles = openFileDialog.FileNames;
                foreach (string filePath in selectedFiles)
                {
                    fileExt = Path.GetExtension(filePath).ToLower();
                    fileName = Path.GetFileName(filePath);
                    nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                    if (folderPath.Equals(Path.GetDirectoryName(filePath)))
                    {
                        isParentFolder = true;
                    }
                    folderPath = Path.GetDirectoryName(filePath);

                    if (fileExt == ".usv")
                    {
                        if (!nameWithoutExt.EndsWith("ONEGAI", StringComparison.OrdinalIgnoreCase))
                        {
                            MessageBox.Show($"USV file '{fileName}' must end with 'ONEGAI'", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            Cursor = Cursors.Default;
                            return;
                        }
                    }
                    else if (fileExt == ".xls" || fileExt == ".xlsx")
                    {
                        if (!nameWithoutExt.EndsWith("PackingList", StringComparison.OrdinalIgnoreCase))
                        {
                            MessageBox.Show($"Excel file '{fileName}' must end with 'PackingList'", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            Cursor = Cursors.Default;
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show($"UnSupported file type: {fileName}", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Cursor = Cursors.Default;
                        return;
                    }
                    string preparedFile = templateHTML.PrepareFile(filePath);
                    textBoxFile.Text += filePath + Environment.NewLine;
                    bool isTempFile = preparedFile != filePath;

                    using (FileStream fs = new FileStream(preparedFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        LoadDataGirdview(preparedFile, isParentFolder);
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
                            Cursor = Cursors.Default;
                            return;
                        }
                    }
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
            bool isParentFolder = false;
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            HtmlResponse templateHTML = new HtmlResponse();
            foreach (string originalPath in files)
            {
                string preparedFile = templateHTML.PrepareFile(originalPath);
                textBoxFile.Text += originalPath + Environment.NewLine;
                bool isTempFile = preparedFile != originalPath;
                string extension = Path.GetExtension(originalPath).ToLower();

                string fileName = Path.GetFileName(originalPath);
                if (folderPath.Equals(Path.GetDirectoryName(originalPath)))
                {
                    isParentFolder = true;
                }
                folderPath = Path.GetDirectoryName(originalPath);
                string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                if (extension == ".usv")
                {
                    if (!nameWithoutExt.EndsWith("ONEGAI", StringComparison.OrdinalIgnoreCase))
                    {
                        MessageBox.Show($"USV file '{fileName}' must end with 'ONEGAI'", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }
                else if (extension == ".xls" || extension == ".xlsx")
                {
                    if (!nameWithoutExt.EndsWith("PackingList", StringComparison.OrdinalIgnoreCase))
                    {
                        MessageBox.Show($"Excel file '{fileName}' must end with 'PackingList'", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show($"UnSupported file type: {fileName}", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                try
                {
                    using (FileStream fs = new FileStream(preparedFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        LoadDataGirdview(preparedFile, isParentFolder);
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

        private void FrmPackingSealMontbell_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.frmMontbellState = this.WindowState;
            if (this.WindowState == FormWindowState.Normal)
            {
                Properties.Settings.Default.frmMontbellLocation = this.Location;
                Properties.Settings.Default.frmMontbellSize = this.Size;
            }
            else
            {
                Properties.Settings.Default.frmMontbellLocation = this.RestoreBounds.Location;
                Properties.Settings.Default.frmMontbellSize = this.RestoreBounds.Size;
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

        private void FrmPackingSealMontbell_KeyDown(object sender, KeyEventArgs e)
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

        private void CreateDataTable()
        {
            dataTable = new DataTable();
            dataTable.Columns.Add("MaterialCode");
            dataTable.Columns.Add("KeyProductItemCode");
            dataTable.Columns.Add("ItemCode");
            dataTable.Columns.Add("PONo");
            dataTable.Columns.Add("OrderNo");
            dataTable.Columns.Add("StyleNo");
            dataTable.Columns.Add("ColorName");
            dataTable.Columns.Add("SizeName");
            dataTable.Columns.Add("SizeQuantity");
            dataTable.Columns.Add("Quantity");
            dataTable.Columns.Add("QuantityInBox");
            dataTable.Columns.Add("PCS");
            dataTable.Columns.Add("Key");
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
                    //notifyIcon1.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool isUpdating = false;
        private void dgvData_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (isUpdating)
            {
                return;
            }
            if (dgvData.Columns.Contains("SizeQuantity") && e.ColumnIndex == dgvData.Columns["SizeQuantity"].Index)
            {
                int rowIndex = e.RowIndex;
                if (rowIndex < 0 || rowIndex >= dgvData.Rows.Count)
                {
                    return;
                }
                string keyProductItemCode = dgvData.Rows[rowIndex].Cells["KeyProductItemCode"].Value?.ToString() ?? "";
                string sizeQuantityStr = dgvData.Rows[rowIndex].Cells["SizeQuantity"].Value?.ToString() ?? "1";
                sizeQuantityStr = sizeQuantityStr.Replace(",", "").Trim();

                if (!int.TryParse(sizeQuantityStr.Split('/')[0], out int quantity) || quantity <= 0)
                {
                    MessageBox.Show("Quantity must be greater than 0.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    dgvData.Rows[rowIndex].Cells["SizeQuantity"].Value = "1";
                    dgvData.Rows[rowIndex].Cells["QuantityInBox"].Value = "1";
                    return;
                }
                if (typeProduct == 1)
                {
                    int boxSize = 3200;

                    if (boxSize == 0)
                    {
                        dgvData.Rows[rowIndex].Cells["SizeQuantity"].Value = quantity.ToString("N0");
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
                    string colorName = dgvData.Rows[rowIndex].Cells["ColorName"].Value?.ToString() ?? "";
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
                            r.Field<string>("KeyProductItemCode") == keyProductItemCode &&
                            r.Field<string>("SizeName") == sizeName &&
                            r.Field<string>("ColorName") == colorName &&
                            r.Field<string>("OrderNo") == orderNo
                        ).ToList();

                    foreach (var r in rowsToDelete)
                    {
                        dt.Rows.Remove(r);
                    }

                    dgvData.Rows[rowIndex].Cells["PCS"].Value = "1";
                    dgvData.Rows[rowIndex].Cells["QuantityInBox"].Value = boxQuantities[0].ToString("N0");
                    dgvData.Rows[rowIndex].Cells["SizeQuantity"].Value = quantity.ToString("N0");

                    for (int i = 1; i < boxQuantities.Count; i++)
                    {
                        DataRow newRow = dt.NewRow();
                        newRow["MaterialCode"] = dgvData.Rows[rowIndex].Cells["MaterialCode"].Value?.ToString() ?? "";
                        newRow["KeyProductItemCode"] = keyProductItemCode;
                        newRow["ItemCode"] = dgvData.Rows[rowIndex].Cells["ItemCode"].Value?.ToString() ?? "";
                        newRow["PONo"] = dgvData.Rows[rowIndex].Cells["PONo"].Value?.ToString() ?? "";
                        newRow["OrderNo"] = orderNo;
                        newRow["StyleNo"] = dgvData.Rows[rowIndex].Cells["StyleNo"].Value?.ToString() ?? "";
                        newRow["ColorName"] = dgvData.Rows[rowIndex].Cells["ColorName"].Value?.ToString() ?? "";
                        newRow["SizeName"] = sizeName;
                        newRow["Quantity"] = dgvData.Rows[rowIndex].Cells["Quantity"].Value?.ToString() ?? "";
                        newRow["SizeQuantity"] = quantity.ToString("N0"); ;
                        newRow["QuantityInBox"] = boxQuantities[i].ToString("N0");
                        newRow["PCS"] = "1";

                        dt.Rows.InsertAt(newRow, baseIndex + i);
                    }

                    //Tính Quantity tổng theo keyProductItemCode
                    var groups = dt.AsEnumerable().GroupBy(r => r.Field<string>("KeyProductItemCode") + "|" +
                                                            r.Field<string>("OrderNo"));

                    foreach (var g in groups)
                    {
                        int totalQty = g.Sum(r =>
                        {
                            string qtyStr = r.Field<string>("QuantityInBox") ?? "0";
                            int.TryParse(qtyStr.Replace(",", ""), out int q);
                            return q;
                        });

                        foreach (var r in g)
                        {
                            r["Quantity"] = totalQty.ToString("N0");
                        }
                    }
                }
                else
                {
                    dgvData.Rows[rowIndex].Cells["SizeQuantity"].Value = quantity.ToString("N0");
                    DataTable dt = dgvData.DataSource as DataTable;
                    if (dt == null)
                    {
                        return;
                    }
                    var groups = dt.AsEnumerable().GroupBy(r => r.Field<string>("KeyProductItemCode") + "|" +
                                                r.Field<string>("OrderNo"));

                    foreach (var g in groups)
                    {
                        int totalQty = g.Sum(r =>
                        {
                            string qtyStr = r.Field<string>("SizeQuantity") ?? "0";
                            int.TryParse(qtyStr.Replace(",", ""), out int q);
                            return q;
                        });

                        foreach (var r in g)
                        {
                            r["Quantity"] = totalQty.ToString("N0");
                        }
                    }
                }

                st_count_1.Text = dgvData.Rows.Count.ToString();
                isUpdating = false;
            }
        }

        private void radioButtonTag_CheckedChanged(object sender, EventArgs e)
        {
            typeProduct = 1;
            dgvData.Columns["QuantityInBox"].Visible = true;
            st_count_1.Text = "0";
            dataTable.Rows.Clear();
            dgvData.DataSource = dataTable;
            textBoxFile.Clear();
            cleanTempPDFFiles();
            webView21.CoreWebView2.Navigate($@"{Path.GetDirectoryName(assembly.Location)}\Template\Default.html");
            comboBoxSize.SelectedValue = "0";
            comboBoxSize.Enabled = true;
        }

        private void radioButtonCare_CheckedChanged(object sender, EventArgs e)
        {
            typeProduct = 2;
            dgvData.Columns["QuantityInBox"].Visible = false;
            st_count_1.Text = "0";
            dataTable.Rows.Clear();
            dgvData.DataSource = dataTable;
            textBoxFile.Clear();
            cleanTempPDFFiles();
            webView21.CoreWebView2.Navigate($@"{Path.GetDirectoryName(assembly.Location)}\Template\Default.html");
            comboBoxSize.SelectedValue = "0";
            comboBoxSize.Enabled = false;
        }

        private void comboBoxSize_SelectionChangeCommitted(object sender, EventArgs e)
        {
            // Nếu chọn Size = 0 (All in one box) thì không cần gom lại
            if (comboBoxSize.SelectedValue.ToString() == "0")
            {
                dgvData.DataSource = originalData.Copy();
                return;
            }

            // Gom nhóm theo Key
            var keys = new HashSet<string>();
            foreach (DataGridViewRow dgvRow in dgvData.Rows)
            {
                if (!dgvRow.IsNewRow && dgvRow.Cells["Key"].Value != null)
                {
                    keys.Add(dgvRow.Cells["Key"].Value.ToString());
                }
            }

            foreach (string key in keys)
            {
                string keyProductItemCode = null;
                string materialCode = null;
                object itemCode = null;
                object poNo = null;
                object orderNo = null;
                object styleNo = null;
                object colorName = null;
                object sizeName = null;
                object quantitykeyProductItemCode = null;

                int sizeQuantity = 0;

                // Lấy tất cả dòng cần xóa từ dataTable
                var rowsToRemove = dataTable.AsEnumerable()
                    .Where(r => r["Key"].ToString() == key)
                    .ToList();

                if (rowsToRemove.Count == 0)
                {
                    continue;
                }

                // Lấy dữ liệu từ dòng đầu tiên trước khi xóa
                var firstRow = rowsToRemove.First();
                keyProductItemCode = firstRow["KeyProductItemCode"].ToString();
                materialCode = firstRow["MaterialCode"].ToString();
                itemCode = firstRow["ItemCode"];
                poNo = firstRow["PoNO"];
                orderNo = firstRow["OrderNo"];
                styleNo = firstRow["StyleNo"];
                colorName = firstRow["ColorName"];
                sizeName = firstRow["SizeName"];
                quantitykeyProductItemCode = firstRow["Quantity"];
                sizeQuantity = int.Parse(firstRow["SizeQuantity"].ToString(), NumberStyles.AllowThousands, CultureInfo.CurrentCulture);

                // Xóa tất cả dòng cũ
                foreach (DataRow r in rowsToRemove)
                {
                    dataTable.Rows.Remove(r);
                }

                int QuantityInboxValue = Convert.ToInt32(comboBoxSize.SelectedValue.ToString());

                // Nếu không có QuantityInboxValue → gom thành 1 dòng duy nhất
                if (QuantityInboxValue == 0)
                {
                    DataRow newRow = dataTable.NewRow();
                    newRow["MaterialCode"] = materialCode;
                    newRow["KeyProductItemCode"] = keyProductItemCode;
                    newRow["ItemCode"] = itemCode;
                    newRow["PoNO"] = poNo;
                    newRow["OrderNo"] = orderNo;
                    newRow["StyleNo"] = styleNo;
                    newRow["ColorName"] = colorName;
                    newRow["SizeName"] = sizeName;
                    newRow["SizeQuantity"] = sizeQuantity;
                    newRow["Quantity"] = quantitykeyProductItemCode;
                    newRow["PCS"] = "1";
                    newRow["Key"] = key;
                    if (dataTable.Columns.Contains("QuantityInBox"))
                    {
                        newRow["QuantityInBox"] = sizeQuantity;
                    }
                    dataTable.Rows.Add(newRow);
                }
                else
                {
                    // Tính số lượng box
                    int fullBoxes = sizeQuantity / QuantityInboxValue;
                    int lastBoxQty = sizeQuantity % QuantityInboxValue;

                    for (int i = 0; i < fullBoxes; i++)
                    {
                        DataRow newRow = dataTable.NewRow();
                        newRow["MaterialCode"] = materialCode;
                        newRow["KeyProductItemCode"] = keyProductItemCode;
                        newRow["ItemCode"] = itemCode;
                        newRow["PoNO"] = poNo;
                        newRow["OrderNo"] = orderNo;
                        newRow["StyleNo"] = styleNo;
                        newRow["ColorName"] = colorName;
                        newRow["SizeName"] = sizeName;
                        newRow["SizeQuantity"] = sizeQuantity;
                        newRow["Quantity"] = quantitykeyProductItemCode;
                        newRow["PCS"] = "1";
                        newRow["Key"] = key;
                        if (dataTable.Columns.Contains("QuantityInBox"))
                        {
                            newRow["QuantityInBox"] = QuantityInboxValue;
                        }
                        dataTable.Rows.Add(newRow);
                    }

                    if (lastBoxQty > 0)
                    {
                        DataRow newRow = dataTable.NewRow();
                        newRow["MaterialCode"] = materialCode;
                        newRow["KeyProductItemCode"] = keyProductItemCode;
                        newRow["ItemCode"] = itemCode;
                        newRow["PoNO"] = poNo;
                        newRow["OrderNo"] = orderNo;
                        newRow["StyleNo"] = styleNo;
                        newRow["ColorName"] = colorName;
                        newRow["SizeName"] = sizeName;
                        newRow["SizeQuantity"] = sizeQuantity;
                        newRow["Quantity"] = quantitykeyProductItemCode;
                        newRow["PCS"] = "1";
                        newRow["Key"] = key;
                        if (dataTable.Columns.Contains("QuantityInBox"))
                        {
                            newRow["QuantityInBox"] = lastBoxQty;
                        }
                        dataTable.Rows.Add(newRow);
                    }
                }
            }
            dgvData.DataSource = dataTable;
        }

    }
}
