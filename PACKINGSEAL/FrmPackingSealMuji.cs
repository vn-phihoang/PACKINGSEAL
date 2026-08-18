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
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PACKINGSEAL
{
    public partial class FrmPackingSealMuji : Form
    {
        public string imagePath = string.Empty;
        private string htmlPath = string.Empty;
        private string pdfOutputPath = string.Empty;
        private string folderPath = string.Empty;
        private DataTable dataTable = new DataTable();
        private DataTable sizeTable = new DataTable();
        // Khai báo biến để giữ dữ liệu gốc
        private DataTable originalData;
        private string poNoValue = string.Empty;
        private string poCustomerValue = string.Empty;
        private string usvImportPath = string.Empty;
        private string deliveryDestinationValue = string.Empty;
        private Assembly assembly = Assembly.GetExecutingAssembly();

        public FrmPackingSealMuji(FilePathModel filePathModel)
        {
            InitializeComponent();
            string UserName = Environment.MachineName;
            lb_info_1.Text = "HC | Muji Packing Onegai Seal- " + UserName;
            string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
            var version = "";
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                version = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            this.Text = "MUJI PACKING ONEGAI SEAL (梱包シール) Ver: " + version;
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

        private void LoadComboboxSize()
        {
            sizeTable = new DataTable("SizeDefinition");
            sizeTable.Columns.Add("SizeCode", typeof(string));
            sizeTable.Columns.Add("Quantity", typeof(int));
            sizeTable.Rows.Add("", 0);
            sizeTable.Rows.Add("170x46", 1000);
            sizeTable.Rows.Add("250x54", 1200);
            sizeTable.Rows.Add("280x54", 1200);
            sizeTable.Rows.Add("280x63", 1200);
            sizeTable.Rows.Add("315x54", 1200);
            sizeTable.Rows.Add("340x54", 1200);
            sizeTable.Rows.Add("364x54", 1200);
            sizeTable.Rows.Add("365x54", 1200);
            sizeTable.Rows.Add("440x54", 800);
            sizeTable.Rows.Add("450x54", 800);
            sizeTable.Rows.Add("470x54", 900);
            sizeTable.Rows.Add("480x54", 900);
            sizeTable.Rows.Add("480x63", 900);
            sizeTable.Rows.Add("485x63", 900);
            sizeTable.Rows.Add("490x63", 900);
            sizeTable.Rows.Add("62x110", 600);
            sizeTable.Rows.Add("50x103", 2000);
            sizeTable.Rows.Add("50x115", 2000);
            sizeTable.Rows.Add("103.5x197", 600);
            
            comboBoxSize.DataSource = sizeTable;
            comboBoxSize.DisplayMember = "SizeCode";
            comboBoxSize.ValueMember = "Quantity";

            comboBoxSize.DrawMode = DrawMode.OwnerDrawFixed;
            comboBoxSize.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBoxSize.AutoCompleteSource = AutoCompleteSource.ListItems;
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
                        poNoValue = string.Empty;
                        poCustomerValue = string.Empty;
                        deliveryDestinationValue = string.Empty;
                        comboBoxSize.SelectedIndex = 0;
                    }

                    if (File.Exists(filePath))
                    {
                        if (string.IsNullOrEmpty(poNoValue))
                        {
                            DirectoryInfo dir = new DirectoryInfo(filePath);
                            // Lấy thư mục cha của file
                            poNoValue = dir.Parent?.Parent?.Name;
                            poCustomerValue = dir.Parent?.Name;

                            if (!string.IsNullOrEmpty(poNoValue))
                            {
                                AppCommonModule appCommonModule = new AppCommonModule();
                                deliveryDestinationValue = appCommonModule.GetCustomer(poNoValue);
                            }
                        }

                        using (TextFieldParser parser = new TextFieldParser(filePath))
                        {
                            parser.TextFieldType = FieldType.Delimited;
                            parser.SetDelimiters(",");

                            while (!parser.EndOfData)
                            {
                                string[] fields = parser.ReadFields();

                                if (fields.Length >= 10)
                                {
                                    if (string.IsNullOrEmpty(deliveryDestinationValue))
                                    {
                                        deliveryDestinationValue = fields[2]?.Trim().Replace("\"", "");
                                    }
                                    if (string.IsNullOrEmpty(poNoValue) && fields.Length > 42)
                                    {
                                        poNoValue = fields[42]?.Trim().Replace("\"", "");
                                    }

                                    DataRow dataRow = dataTable.NewRow();
                                    dataRow["DeliveryDestination"] = deliveryDestinationValue;
                                    dataRow["KeyNo"] = fields[5]?.Trim().Replace("\"", "");
                                    dataRow["QuantityInBox"] = "0";
                                    dataRow["Quantity"] = Convert.ToInt32(fields[6]?.Trim().Replace("\"", ""));
                                    dataRow["ItemCode"] = fields[7]?.Trim().Replace("\"", "");
                                    dataRow["PoNO"] = poNoValue;
                                    dataRow["POCustomer"] = poCustomerValue;
                                    dataRow["PCS"] = "1";

                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                        }

                        usvImportPath = filePath;
                    }
                    else
                    {
                        MessageBox.Show($"File USV null.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Cursor = Cursors.Default;
                        return;
                    }
                }
                else
                {
                    //if (isParentFolder == false)
                    //{
                    //    CreateDataTable();
                    //    comboBoxSize.SelectedIndex = 0;
                    //}
                    var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                    var workbook = WorkbookFactory.Create(stream);
                    var sheet = workbook.GetSheetAt(0);
                    if (sheet == null)
                    {
                        MessageBox.Show($"Sheet with name not found: {sheet}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Cursor = Cursors.Default;
                        return;
                    }

                    Dictionary<string, int> quantityTotal = new Dictionary<string, int>();

                    deliveryDestinationValue = excelUtil.GetCellValueWithMerge(sheet, 7, 4);
                    poNoValue = excelUtil.GetCellValueWithMerge(sheet, 2, 13);

                    foreach (DataRow row in dataTable.Rows)
                    {
                        row["DeliveryDestination"] = deliveryDestinationValue;
                        row["PoNO"] = poNoValue;
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
            MujiModel mujiModel = new MujiModel();
            mujiModel.ImagePath = imagePath;
            mujiModel.DeliveryDestination = row.Cells["DeliveryDestination"].Value?.ToString() ?? string.Empty;
            mujiModel.PoNo = row.Cells["PONo"].Value?.ToString() ?? string.Empty;
            mujiModel.POCustomer = row.Cells["POCustomer"].Value?.ToString() ?? string.Empty;
            mujiModel.ItemCode = row.Cells["ItemCode"].Value?.ToString() ?? string.Empty;
            mujiModel.KeyNo = row.Cells["KeyNo"].Value?.ToString() ?? string.Empty;
            mujiModel.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString());
            mujiModel.QuantityInBox = string.IsNullOrEmpty(row.Cells["QuantityInBox"].Value.ToString()) ? 0 : double.Parse(row.Cells["QuantityInBox"].Value.ToString());
            htmlContent = templateHTML.MujiToWebview(mujiModel);
            RenderHTMLToWebView(htmlContent);
        }

        private async void FrmPackingSealMuji_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.frmMujiSize.Width != 0 && Properties.Settings.Default.frmMujiSize.Height != 0)
            {
                this.Size = Properties.Settings.Default.frmMujiSize;
            }
            if (Properties.Settings.Default.frmMujiLocation.X != 0 && Properties.Settings.Default.frmMujiLocation.Y != 0)
            {
                this.StartPosition = FormStartPosition.CenterScreen;
                this.Location = Properties.Settings.Default.frmMujiLocation;
            }
            this.WindowState = Properties.Settings.Default.frmMujiState;

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
                MujiModel mujiModel = new MujiModel();
                mujiModel.ImagePath = imagePath;
                mujiModel.DeliveryDestination = row.Cells["DeliveryDestination"].Value?.ToString() ?? string.Empty;
                mujiModel.PoNo = row.Cells["PONo"].Value?.ToString() ?? string.Empty;
                mujiModel.POCustomer = row.Cells["POCustomer"].Value?.ToString() ?? string.Empty;
                mujiModel.ItemCode = row.Cells["ItemCode"].Value?.ToString() ?? string.Empty;
                mujiModel.KeyNo = row.Cells["KeyNo"].Value?.ToString() ?? string.Empty;
                mujiModel.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString());
                mujiModel.QuantityInBox = string.IsNullOrEmpty(row.Cells["QuantityInBox"].Value.ToString()) ? 0 : double.Parse(row.Cells["QuantityInBox"].Value.ToString());
                htmlContent = templateHTML.MujiToWebview(mujiModel);
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

                    //MessageBox.Show("PDF printed successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    System.Threading.Thread.Sleep(3000);
                    //new Utils.PDFSetting().showNotification("SUCCESS", "PDF printed successfully.");

                    var resModel = new PrintResponseModel
                    {
                        Brand = "MUJI",
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
                        MujiModel mujiModel = new MujiModel();
                        mujiModel.ImagePath = imagePath;
                        mujiModel.DeliveryDestination = row.Cells["DeliveryDestination"].Value?.ToString() ?? string.Empty;
                        mujiModel.PoNo = row.Cells["PONo"].Value?.ToString() ?? string.Empty;
                        mujiModel.POCustomer = row.Cells["POCustomer"].Value?.ToString() ?? string.Empty;
                        mujiModel.ItemCode = row.Cells["ItemCode"].Value?.ToString() ?? string.Empty;
                        mujiModel.KeyNo = row.Cells["KeyNo"].Value?.ToString() ?? string.Empty;
                        mujiModel.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString());
                        mujiModel.QuantityInBox = string.IsNullOrEmpty(row.Cells["QuantityInBox"].Value.ToString()) ? 0 : double.Parse(row.Cells["QuantityInBox"].Value.ToString());
                        string tableContent = templateHTML.MujiToPDF(mujiModel);
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
                    Filter = "USV or Excel Files|*.usv;*.xls;*.xlsx",
                    Title = "Select USV or Excel File",
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

                    if (fileExt == ".usv")
                    {
                        if (folderPath.Equals(Path.GetDirectoryName(Path.GetDirectoryName(filePath))))
                        {
                            isParentFolder = true;
                        }
                        else
                        {
                            folderPath = Path.GetDirectoryName(Path.GetDirectoryName(filePath));
                        }
                    }
                    else if (fileExt == ".xls" || fileExt == ".xlsx")
                    {
                        if (folderPath.Equals(Path.GetDirectoryName(filePath)))
                        {
                            isParentFolder = true;
                        }
                        else
                        {
                            folderPath = Path.GetDirectoryName(filePath);
                        }
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
                if (folderPath.Equals(Path.GetDirectoryName(originalPath)) || usvImportPath.Equals(fileName))
                {
                    isParentFolder = true;
                }
                folderPath = Path.GetDirectoryName(fileName);
                string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);

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

        private void FrmPackingSealMuji_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.frmMujiState = this.WindowState;
            if (this.WindowState == FormWindowState.Normal)
            {
                Properties.Settings.Default.frmMujiLocation = this.Location;
                Properties.Settings.Default.frmMujiSize = this.Size;
            }
            else
            {
                Properties.Settings.Default.frmMujiLocation = this.RestoreBounds.Location;
                Properties.Settings.Default.frmMujiSize = this.RestoreBounds.Size;
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
            dataTable.Columns.Add("DeliveryDestination");
            dataTable.Columns.Add("ItemCode");
            dataTable.Columns.Add("PONo");
            dataTable.Columns.Add("POCustomer");
            dataTable.Columns.Add("KeyNo");
            dataTable.Columns.Add("QuantityInBox", typeof(int));
            dataTable.Columns.Add("Quantity", typeof(int));
            dataTable.Columns.Add("PCS");
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

        private void comboBoxSize_SelectionChangeCommitted(object sender, EventArgs e)
        {
            //2026/02/06 check nếu chọn Size = 0 (All in one box) thì không cần gom lại
            if (comboBoxSize.SelectedValue.ToString() == "0")
            {
                dgvData.DataSource = originalData.Copy();
                return;
            }

            // Gom nhóm theo KeyNo
            var keyNos = new HashSet<string>();
            foreach (DataGridViewRow dgvRow in dgvData.Rows)
            {
                if (!dgvRow.IsNewRow && dgvRow.Cells["KeyNo"].Value != null)
                {
                    keyNos.Add(dgvRow.Cells["KeyNo"].Value.ToString());
                }
            }

            foreach (string keyNo in keyNos)
            {
                object deliveryDestination = null;
                object itemCode = null;
                object poNo = null;
                object poCustomer = null;
                int quantity = 0;

                // Lấy tất cả dòng cần xóa từ dataTable
                var rowsToRemove = dataTable.AsEnumerable()
                    .Where(r => r["KeyNo"].ToString() == keyNo)
                    .ToList();

                if (rowsToRemove.Count == 0) 
                {
                    continue; 
                }

                // Lấy dữ liệu từ dòng đầu tiên trước khi xóa
                var firstRow = rowsToRemove.First();
                deliveryDestination = firstRow["DeliveryDestination"];
                itemCode = firstRow["ItemCode"];
                poNo = firstRow["PoNO"];
                poCustomer = firstRow["POCustomer"];
                quantity = Convert.ToInt32(firstRow["Quantity"]);

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
                    newRow["KeyNo"] = keyNo;
                    newRow["DeliveryDestination"] = deliveryDestination;
                    newRow["ItemCode"] = itemCode;
                    newRow["PoNO"] = poNo;
                    newRow["POCustomer"] = poCustomer;
                    newRow["Quantity"] = quantity;
                    newRow["PCS"] = "1";
                    if (dataTable.Columns.Contains("QuantityInBox"))
                    {
                        newRow["QuantityInBox"] = quantity;
                    }
                    dataTable.Rows.Add(newRow);
                }
                else
                {
                    // Tính số lượng box
                    int fullBoxes = quantity / QuantityInboxValue;
                    int lastBoxQty = quantity % QuantityInboxValue;

                    for (int i = 0; i < fullBoxes; i++)
                    {
                        DataRow newRow = dataTable.NewRow();
                        newRow["KeyNo"] = keyNo;
                        newRow["DeliveryDestination"] = deliveryDestination;
                        newRow["ItemCode"] = itemCode;
                        newRow["PoNO"] = poNo;
                        newRow["POCustomer"] = poCustomer;
                        newRow["Quantity"] = quantity;
                        newRow["PCS"] = "1";
                        if (dataTable.Columns.Contains("QuantityInBox"))
                        {
                            newRow["QuantityInBox"] = QuantityInboxValue;
                        }
                        dataTable.Rows.Add(newRow);
                    }

                    if (lastBoxQty > 0)
                    {
                        DataRow newRow = dataTable.NewRow();
                        newRow["KeyNo"] = keyNo;
                        newRow["DeliveryDestination"] = deliveryDestination;
                        newRow["ItemCode"] = itemCode;
                        newRow["PoNO"] = poNo;
                        newRow["POCustomer"] = poCustomer;
                        newRow["Quantity"] = quantity;
                        newRow["PCS"] = "1";
                        if (dataTable.Columns.Contains("QuantityInBox"))
                        {
                            newRow["QuantityInBox"] = lastBoxQty;
                        }
                        dataTable.Rows.Add(newRow);
                    }
                }
            }

            // Sau khi xử lý xong, gán lại DataSource
            dgvData.DataSource = dataTable;
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
                    dgvData.Rows[rowIndex].Cells["QuantityInBox"].Value = "1";
                    return;
                }

                double QuantityInbox = 0;
                if (!string.IsNullOrEmpty(comboBoxSize.SelectedValue.ToString()))
                {
                    QuantityInbox = Convert.ToInt32(comboBoxSize.SelectedValue.ToString());
                }

                if (QuantityInbox == 0)
                {
                    dgvData.Rows[rowIndex].Cells["Quantity"].Value = quantity;
                    dgvData.Rows[rowIndex].Cells["QuantityInBox"].Value = quantity;
                    return;
                }

                // Tính số lượng box
                List<double> boxQuantities = new List<double>();
                double fullBoxes = quantity / QuantityInbox;
                double lastBoxQty = quantity % QuantityInbox;

                for (int i = 1; i < fullBoxes; i++)
                {
                    boxQuantities.Add(QuantityInbox);
                }
                if (lastBoxQty > 0)
                {
                    boxQuantities.Add(lastBoxQty);
                }

                string keyNo = dgvData.Rows[rowIndex].Cells["KeyNo"].Value?.ToString() ?? "";
                string itemCode = dgvData.Rows[rowIndex].Cells["ItemCode"].Value?.ToString() ?? "";
                string pONo = dgvData.Rows[rowIndex].Cells["PONo"].Value?.ToString() ?? "";

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
                        r.Field<string>("KeyNo") == keyNo &&
                        r.Field<string>("ItemCode") == itemCode &&
                        r.Field<string>("PONo") == pONo
                    ).ToList();

                foreach (var r in rowsToDelete)
                {
                    dt.Rows.Remove(r);
                }

                dgvData.Rows[rowIndex].Cells["PCS"].Value = "1";
                dgvData.Rows[rowIndex].Cells["QuantityInBox"].Value = boxQuantities[0];
                dgvData.Rows[rowIndex].Cells["Quantity"].Value = quantity;

                for (int i = 1; i < boxQuantities.Count; i++)
                {
                    DataRow newRow = dt.NewRow();
                    newRow["DeliveryDestination"] = dgvData.Rows[rowIndex].Cells["DeliveryDestination"].Value?.ToString() ?? "";
                    newRow["ItemCode"] = dgvData.Rows[rowIndex].Cells["ItemCode"].Value?.ToString() ?? "";
                    newRow["PONo"] = dgvData.Rows[rowIndex].Cells["PONo"].Value?.ToString() ?? "";
                    newRow["POCustomer"] = dgvData.Rows[rowIndex].Cells["POCustomer"].Value?.ToString() ?? "";
                    newRow["KeyNo"] = dgvData.Rows[rowIndex].Cells["KeyNo"].Value?.ToString() ?? "";
                    newRow["QuantityInBox"] = boxQuantities[i];
                    newRow["Quantity"] = quantity;
                    newRow["PCS"] = "1";

                    dt.Rows.InsertAt(newRow, baseIndex + i);
                }
                st_count_1.Text = dgvData.Rows.Count.ToString();
                isUpdating = false;
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
