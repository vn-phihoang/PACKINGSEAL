using Microsoft.VisualBasic.FileIO;
using NPOI.SS.UserModel;
using PACKINGSEAL.Models;
using PACKINGSEAL.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Deployment.Application;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace PACKINGSEAL
{
    public partial class FrmMontbell : Form
    {
        public string imagePath = string.Empty;
        private string pdfOutputPath = string.Empty;
        private string folderPath = string.Empty;
        private DataTable dataTable = new DataTable();
        private Assembly assembly = Assembly.GetExecutingAssembly();
        private string OrderNoValue = string.Empty;
        private int typeProduct = 1; // 1: Tag Label, 2: Care Label
        private PrintDocument printDocument = new PrintDocument();
        private DataTable sizeTable = new DataTable();
        private DataTable originalData;
        private MontbellModel currentModel;
        private int printedIndex = 0;
        private float zoomFactor = 1.0f;
        private List<MontbellModel> printQueue;
        private Point crosshairPoint = Point.Empty;
        private Bitmap layoutCache;

        public FrmMontbell(FilePathModel filePathModel)
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
            LoadComboboxSize();
            printDocument = new PrintDocument();
            printDocument.PrintPage += printDocument_PrintPage;
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
                            dataRow["LotNo"] = "1/1";
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

            currentModel = new MontbellModel();
            currentModel.ImagePath = imagePath;
            currentModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
            currentModel.PoNo = row.Cells["PONo"].Value?.ToString() ?? string.Empty;
            currentModel.ItemCode = row.Cells["ItemCode"].Value?.ToString() ?? string.Empty;
            currentModel.MaterialCode = row.Cells["MaterialCode"].Value?.ToString() ?? string.Empty;
            currentModel.StyleNo = row.Cells["StyleNo"].Value?.ToString() ?? string.Empty;
            currentModel.Color = row.Cells["ColorName"].Value?.ToString() ?? string.Empty;
            currentModel.Size = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
            currentModel.SizeQuantity = string.IsNullOrEmpty(row.Cells["SizeQuantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["SizeQuantity"].Value.ToString());
            currentModel.LotNo = row.Cells["LotNo"].Value?.ToString() ?? string.Empty;
            currentModel.TypeProduct = typeProduct;

            double.TryParse(row.Cells["Quantity"].Value?.ToString(), out double qty);
            currentModel.Quantity = qty;

            double.TryParse(row.Cells["QuantityInBox"].Value?.ToString(), out double qtyBox);
            currentModel.QuantityInBox = qtyBox;

            printPreviewControl1.Document = printDocument;
            printPreviewControl1.InvalidatePreview();
            panel12.Invalidate();
        }

        private void DrawSingleLayout(Graphics g, MontbellModel model)
        {
            if (model == null)
            {
                return;
            }

            float width = 345;   // 8.7 cm
            float height = 175;  // 4.4 cm
            System.Drawing.Color customColor = ColorTranslator.FromHtml("#10274B");
            Pen borderPen = new Pen(customColor, 1);
            Brush textBrush = new SolidBrush(customColor);

            // Font từ PrivateFontCollection
            PrivateFontCollection pfc = new PrivateFontCollection();
            pfc.AddFontFile($@"{Path.GetDirectoryName(assembly.Location)}\Fonts\MS-UIGothic.ttf");
            pfc.AddFontFile($@"{Path.GetDirectoryName(assembly.Location)}\Fonts\NotoSansJP-ExtraBold.ttf");

            Font fontRegular = new Font(pfc.Families[0], 9, FontStyle.Regular);
            Font fontBold = new Font(pfc.Families[1], 9, FontStyle.Bold);

            // Vẽ khung ngoài
            Rectangle outerRect = new Rectangle(0, 0, (int)width, (int)height);
            RoundedRectangleHelper.Draw(g, borderPen, outerRect, 5);

            int padding = 7;
            float contentHeight = 90; // vùng nội dung
            float rowHeight = contentHeight / 5f;
            int spacing = 2;
            int titleWidth = 115;

            // ---------------- Order No ----------------
            float yRow1 = padding;
            Rectangle labelRectRow1 = new Rectangle(padding, (int)yRow1, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow1);
            g.DrawString("Order No", fontRegular, textBrush,
                         labelRectRow1.Left + 2, labelRectRow1.Top + (labelRectRow1.Height - fontRegular.Height) / 2);

            Rectangle valueRectOrderNo = new Rectangle(padding + titleWidth, (int)yRow1,
                (int)(width - titleWidth - 2 * padding), (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectOrderNo.Left, valueRectOrderNo.Bottom, valueRectOrderNo.Right, valueRectOrderNo.Bottom);

            string textOrderNo = model.OrderNo;
            SizeF sizeOrderNo = g.MeasureString(textOrderNo, fontRegular);
            Font fontOrderNo = fontRegular;
            if (sizeOrderNo.Width > valueRectOrderNo.Width - 5)
            {
                float shrink = (valueRectOrderNo.Width - 5) / sizeOrderNo.Width;
                fontOrderNo = new Font(fontRegular.FontFamily, fontRegular.Size * shrink, fontRegular.Style);
            }
            g.DrawString(textOrderNo, fontOrderNo, textBrush,
                         valueRectOrderNo.Left + 5, valueRectOrderNo.Top + (valueRectOrderNo.Height - fontOrderNo.Height) / 2);

            //----------------TENTAC(PO NO)----------------
            float yRow2 = padding + 1 * rowHeight + spacing * 1;
            Rectangle labelRectRow2 = new Rectangle(padding, (int)yRow2, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow2);
            g.DrawString("TENTAC(PO NO)", fontRegular, textBrush,
                         labelRectRow2.Left + 2, labelRectRow2.Top + (labelRectRow2.Height - fontRegular.Height) / 2);

            int totalValueWidth = (int)(width - titleWidth - 2 * padding);
            int partWidthRow2 = totalValueWidth / 10;
            Rectangle valueRectRow2 = new Rectangle(padding + titleWidth, (int)yRow2, totalValueWidth, (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow2.Left, valueRectRow2.Bottom, valueRectRow2.Right, valueRectRow2.Bottom);

            // 4 cột: 3/3/3/1
            Rectangle valueRectLeft = new Rectangle(valueRectRow2.Left, valueRectRow2.Top, partWidthRow2 * 3, valueRectRow2.Height);
            Rectangle valueRectCenter = new Rectangle(valueRectLeft.Right, valueRectRow2.Top, partWidthRow2 * 2, valueRectRow2.Height);
            Rectangle valueRectRight = new Rectangle(valueRectCenter.Right, valueRectRow2.Top, partWidthRow2 * 4, valueRectRow2.Height);
            Rectangle valueRectPCS = new Rectangle(valueRectRight.Right, valueRectRow2.Top, partWidthRow2 * 1, valueRectRow2.Height);

            // --- PoNo ---
            SizeF textSizePoNo = g.MeasureString(model.PoNo, fontBold);
            Font drawFontPoNo = fontBold;
            if (textSizePoNo.Width > valueRectLeft.Width - 10)
            {
                float shrinkRatio = (valueRectLeft.Width - 10) / textSizePoNo.Width;
                float newSize = (fontBold.Size * shrinkRatio) + 1;
                drawFontPoNo = new Font(fontBold.FontFamily, newSize, fontBold.Style);
            }
            float valueTextYPoNo = valueRectLeft.Top + (valueRectLeft.Height - drawFontPoNo.Height) / 2;
            g.DrawString(model.PoNo, drawFontPoNo, textBrush, valueRectLeft.Left + 5, valueTextYPoNo);

            // --- TTL Style ---
            float valueTextYStyleNo = valueRectCenter.Top + (valueRectCenter.Height - fontRegular.Height) / 2;
            g.DrawString("TTL Style", fontRegular, textBrush, valueRectCenter.Left + 5, valueTextYStyleNo);

            // --- Quantity ---
            string valueTextQuantity = model.Quantity.ToString("N0");
            SizeF textSizeQuantity = g.MeasureString(valueTextQuantity, fontRegular);
            Font drawFontQuantity = fontRegular;
            if (textSizeQuantity.Width > valueRectRight.Width - 10)
            {
                float shrinkRatio = (valueRectRight.Width - 10) / textSizeQuantity.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontQuantity = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            StringFormat sfRight = new StringFormat
            {
                Alignment = StringAlignment.Far,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString(valueTextQuantity, drawFontQuantity, textBrush, valueRectRight, sfRight);

            // --- PCS ---
            Font unitFont = new Font(fontRegular.FontFamily, 6, FontStyle.Regular);
            g.DrawString("PCS", unitFont, textBrush, valueRectPCS, sfRight);

            //----------------ITEM Code/Material----------------
            /*float yRow3 = padding + 2 * rowHeight + spacing * 2;
            Rectangle labelRectItemCode = new Rectangle(padding, (int)yRow3, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectItemCode);
            float labelTextItemCode = labelRectItemCode.Top + (labelRectItemCode.Height - fontRegular.Height) / 2;
            g.DrawString("ITEM Code/Material", fontRegular, textBrush, labelRectItemCode.Left + 2, labelTextItemCode);

            Rectangle valueRectItemCode = new Rectangle(padding + titleWidth, (int)yRow3,
                (int)(width - titleWidth - 2 * padding), (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectItemCode.Left, valueRectItemCode.Bottom, valueRectItemCode.Right, valueRectItemCode.Bottom);

            string valueTextItemCode = model.ItemCode + "/" + model.MaterialCode;

            SizeF textSizeItemCode = g.MeasureString(valueTextItemCode, fontRegular);
            Font drawFontItemCode = fontRegular;
            if (textSizeItemCode.Width > valueRectItemCode.Width - 11)
            {
                float shrinkRatio = (valueRectItemCode.Width - 11) / textSizeItemCode.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontItemCode = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            float valueTextYItemCode = valueRectItemCode.Top + (valueRectItemCode.Height - drawFontItemCode.Height) / 2;
            g.DrawString(valueTextItemCode, drawFontItemCode, textBrush, valueRectItemCode.Left + 5, valueTextYItemCode);*/

            //----------------ITEM Code/Material----------------
            float yRow3 = padding + 2 * rowHeight + spacing * 2;
            Rectangle labelRectItemCode = new Rectangle(padding, (int)yRow3, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectItemCode);
            float labelTextItemCode = labelRectItemCode.Top + (labelRectItemCode.Height - fontRegular.Height) / 2;
            g.DrawString("ITEM Code/Material", fontRegular, textBrush, labelRectItemCode.Left + 2, labelTextItemCode);

            Rectangle valueRectItemCode = new Rectangle(padding + titleWidth, (int)yRow3,
                (int)(width - titleWidth - 2 * padding), (int)(rowHeight - spacing));

            g.DrawLine(borderPen, valueRectItemCode.Left, valueRectItemCode.Bottom, valueRectItemCode.Right, valueRectItemCode.Bottom);

            int partWidthItemCode = valueRectItemCode.Width / 10;
            Rectangle valueRectLeftRow3 = new Rectangle(valueRectItemCode.Left, valueRectItemCode.Top, partWidthItemCode * 8, valueRectItemCode.Height);
            Rectangle valueRectRightRow3 = new Rectangle(valueRectLeftRow3.Right, valueRectItemCode.Top, partWidthItemCode * 2, valueRectItemCode.Height);

            // ItemCode/MaterialCode (căn trái)
            string valueTextItemCode = model.ItemCode + "/" + model.MaterialCode;
            SizeF textSizeItemCode = g.MeasureString(valueTextItemCode, fontRegular);
            Font drawFontItemCode = fontRegular;
            if (textSizeItemCode.Width > valueRectLeftRow3.Width - 11)
            {
                float shrinkRatio = (valueRectLeftRow3.Width - 11) / textSizeItemCode.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontItemCode = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            float valueTextYItemCode = valueRectLeftRow3.Top + (valueRectLeftRow3.Height - drawFontItemCode.Height) / 2;
            g.DrawString(valueTextItemCode, drawFontItemCode, textBrush, valueRectLeftRow3.Left + 5, valueTextYItemCode);

            // LotNo (căn phải)
            string lotNoText = model.LotNo;
            SizeF textSizeLotNo = g.MeasureString(lotNoText, fontRegular);
            Font drawFontLotNo = fontRegular;
            if (textSizeLotNo.Width > valueRectRightRow3.Width - 10)
            {
                float shrinkRatio = (valueRectRightRow3.Width - 10) / textSizeLotNo.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontLotNo = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }

            StringFormat sfLot = new StringFormat
            {
                Alignment = StringAlignment.Far,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString(lotNoText, drawFontLotNo, textBrush, valueRectRightRow3, sfLot);

            //----------------Style No----------------
            float yRow4 = padding + 3 * rowHeight + spacing * 3;
            Rectangle labelRectStyleNo = new Rectangle(padding, (int)yRow4, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectStyleNo);
            float labelTextYStyleNo = labelRectStyleNo.Top + (labelRectStyleNo.Height - fontRegular.Height) / 2;
            g.DrawString("Style No", fontRegular, textBrush, labelRectStyleNo.Left + 2, labelTextYStyleNo);
            Rectangle valueRectRow4 = new Rectangle(padding + titleWidth, (int)yRow4,
                (int)(width - titleWidth - 2 * padding), (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow4.Left, valueRectRow4.Bottom, valueRectRow4.Right, valueRectRow4.Bottom);

            // Chia thành 10 phần (4+2+4)
            int totalWidthRow4 = valueRectRow4.Width;
            int partWidthRow4 = totalWidthRow4 / 10;

            Rectangle rectPartStyleNo = new Rectangle(valueRectRow4.Left, valueRectRow4.Top, partWidthRow4 * 4, valueRectRow4.Height);
            Rectangle rectPartSize = new Rectangle(rectPartStyleNo.Right, valueRectRow4.Top, partWidthRow4 * 2, valueRectRow4.Height);
            Rectangle rectPartSizeName = new Rectangle(rectPartSize.Right, valueRectRow4.Top, partWidthRow4 * 4, valueRectRow4.Height);

            // Hàm vẽ chữ co font nếu vượt khung
            void DrawTextInRect(Graphics gr, Rectangle rectChill, string text, Font font, Brush brush, bool alignRight = false)
            {
                if (string.IsNullOrEmpty(text)) return;

                SizeF textSize3 = gr.MeasureString(text, font);
                Font drawFont3 = font;
                if (textSize3.Width > rectChill.Width - 6)
                {
                    float shrinkRatio = (rectChill.Width - 6) / textSize3.Width;
                    float newSize = font.Size * shrinkRatio;
                    drawFont3 = new Font(font.FontFamily, newSize, font.Style);
                }

                StringFormat sf3 = new StringFormat
                {
                    LineAlignment = StringAlignment.Center,
                    Alignment = alignRight ? StringAlignment.Far : StringAlignment.Near
                };

                gr.DrawString(text, drawFont3, brush, rectChill, sf3);
            }

            DrawTextInRect(g, rectPartStyleNo, "\u00A0" + model.StyleNo, fontRegular, textBrush);
            DrawTextInRect(g, rectPartSize, "Size", fontRegular, textBrush);
            DrawTextInRect(g, rectPartSizeName, "\u00A0" + model.Size, fontBold, textBrush, true);

            //----------------Color----------------
            float yRow5 = padding + 4 * rowHeight + spacing * 4;
            Rectangle labelRectColor = new Rectangle(padding, (int)yRow5, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectColor);
            float labelTextYColor = labelRectColor.Top + (labelRectColor.Height - fontRegular.Height) / 2;
            g.DrawString("Color", fontRegular, textBrush, labelRectColor.Left + 2, labelTextYColor);
            Rectangle valueRectRow5 = new Rectangle(padding + titleWidth, (int)yRow5,
                (int)(width - titleWidth - 2 * padding), (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow5.Left, valueRectRow5.Bottom, valueRectRow5.Right, valueRectRow5.Bottom);

            // Chia thành 10 phần (3+3+3+1)
            int totalWidthRow5 = valueRectRow5.Width;
            int partWidthRow5 = totalWidthRow5 / 10;

            Rectangle rectPartColorCode = new Rectangle(valueRectRow5.Left, valueRectRow5.Top, partWidthRow5 * 3, valueRectRow5.Height);
            Rectangle rectPartQTY = new Rectangle(rectPartColorCode.Right, valueRectRow5.Top, partWidthRow5 * 2, valueRectRow5.Height);
            Rectangle rectPartQuantity = new Rectangle(rectPartQTY.Right, valueRectRow5.Top, partWidthRow5 * 4, valueRectRow5.Height);
            Rectangle rectPartUnit = new Rectangle(rectPartQuantity.Right, valueRectRow5.Top, partWidthRow5 * 1, valueRectRow5.Height);

            string quantity = string.Empty;
            if (model.TypeProduct == 1)
            {
                quantity = model.QuantityInBox.ToString("N0") + "/" + model.SizeQuantity.ToString("N0");
            }
            else
            {
                quantity = model.SizeQuantity.ToString("N0");
            }

            DrawTextInRect(g, rectPartColorCode, "\u00A0" + model.Color, fontRegular, textBrush);
            DrawTextInRect(g, rectPartQTY, "QTY", fontRegular, textBrush);
            DrawTextInRect(g, rectPartQuantity, quantity, fontBold, textBrush, true);
            // --- Unit (PCS) ---
            g.DrawString("PCS", unitFont, textBrush, rectPartUnit, sfRight);

            // --- Vùng dưới (1.9 cm) ---
            float bottomTop = padding + contentHeight + spacing * 5;
            float bottomHeight = height - contentHeight - 2 * padding;

            // Notice (chỉ chữ, không khung)
            string noticeText = "Please open and confirm this package immediately on receipt.";
            Font fontSmall = new Font(fontRegular.FontFamily, 6.5f, FontStyle.Regular);
            g.DrawString(noticeText, fontSmall, textBrush, padding, bottomTop);

            // Footer (chỉ chữ, không khung, đặt ngay dưới notice, khít lại)
            string footerText = "MADE IN VIETNAM";
            float footerTop = bottomTop + fontSmall.Height;
            g.DrawString(footerText, fontSmall, textBrush, padding, footerTop);

            // Hình ảnh + QR 
            float imageTop = footerTop + fontSmall.Height + 3;
            float imageHeight = height - imageTop - 2;
            int qrSize = 42;

            if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
            {
                using (System.Drawing.Image img = System.Drawing.Image.FromFile(imagePath))
                {
                    // Giữ tỉ lệ gốc của ảnh
                    float aspectRatio = (float)img.Width / img.Height;

                    // Chiều rộng tối đa: toàn bộ vùng còn lại trừ QR và padding
                    int maxWidth = (int)(width - padding * 2 - qrSize - 10);

                    // Tính kích thước ảnh theo tỉ lệ
                    int targetHeight = (int)imageHeight;
                    int targetWidth = (int)(targetHeight * aspectRatio);

                    // Nếu ảnh quá rộng thì thu lại theo maxWidth
                    if (targetWidth > maxWidth)
                    {
                        targetWidth = maxWidth;
                        targetHeight = (int)(targetWidth / aspectRatio);
                    }

                    Rectangle imgRect = new Rectangle(padding, (int)imageTop, targetWidth, targetHeight);
                    g.DrawImage(img, imgRect);
                }
            }

            // QR code: giữ vuông vắn, đặt cùng hàng với ảnh
            Rectangle qrRect = new Rectangle((int)(width - padding - qrSize), (int)imageTop, qrSize, qrSize);
            g.DrawRectangle(borderPen, qrRect);
            g.DrawString(string.Empty, fontRegular, textBrush,
                qrRect.Left + (qrRect.Width - g.MeasureString(string.Empty, fontRegular).Width) / 2,
                qrRect.Top + (qrRect.Height - fontRegular.Height) / 2);
        }

        private void DrawPrintLayout(Graphics g, MontbellModel model, Rectangle rect)
        {
            if (model == null)
            {
                return;
            }

            System.Drawing.Color customColor = ColorTranslator.FromHtml("#10274B");
            Pen borderPen = new Pen(customColor, 1);
            Brush textBrush = new SolidBrush(customColor);

            // Font từ PrivateFontCollection
            PrivateFontCollection pfc = new PrivateFontCollection();
            pfc.AddFontFile($@"{Path.GetDirectoryName(assembly.Location)}\Fonts\MS-UIGothic.ttf");
            pfc.AddFontFile($@"{Path.GetDirectoryName(assembly.Location)}\Fonts\NotoSansJP-ExtraBold.ttf");

            Font fontRegular = new Font(pfc.Families[0], 9, FontStyle.Regular);
            Font fontBold = new Font(pfc.Families[1], 9, FontStyle.Bold);

            // Vẽ khung ngoài theo rect
            Rectangle outerRect = new Rectangle(rect.Left, rect.Top, rect.Width, rect.Height);
            RoundedRectangleHelper.Draw(g, borderPen, outerRect, 5);

            int padding = 7;
            float contentHeight = rect.Height * 0.5f; // thay vì cố định 90
            float rowHeight = contentHeight / 5f;
            int titleWidth = 115;
            int spacing = 2;

            // Vẽ nội dung tương tự như DrawSingleLayout nhưng sử dụng rect và contentHeight mới
            // ---------------- Order No ----------------
            float yRow1 = rect.Top + padding;
            Rectangle labelRectRow1 = new Rectangle(rect.Left + padding, (int)yRow1, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow1);
            g.DrawString("Order No", fontRegular, textBrush,
                         labelRectRow1.Left + 2, labelRectRow1.Top + (labelRectRow1.Height - fontRegular.Height) / 2);
            Rectangle valueRectOrderNo = new Rectangle(rect.Left + padding + titleWidth, (int)yRow1,
                (int)(rect.Width - titleWidth - 2 * padding), (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectOrderNo.Left, valueRectOrderNo.Bottom, valueRectOrderNo.Right, valueRectOrderNo.Bottom);
            string textOrderNo = model.OrderNo;
            SizeF sizeOrderNo = g.MeasureString(textOrderNo, fontRegular);
            Font fontOrderNo = fontRegular;
            if (sizeOrderNo.Width > valueRectOrderNo.Width - 5)
            {
                float shrink = (valueRectOrderNo.Width - 5) / sizeOrderNo.Width;
                fontOrderNo = new Font(fontRegular.FontFamily, fontRegular.Size * shrink, fontRegular.Style);
            }
            g.DrawString(textOrderNo, fontOrderNo, textBrush,
                         valueRectOrderNo.Left + 5, valueRectOrderNo.Top + (valueRectOrderNo.Height - fontOrderNo.Height) / 2);

            //----------------TENTAC(PO NO)----------------
            float yRow2 = rect.Top + padding + 1 * rowHeight + spacing * 1;
            Rectangle labelRectRow2 = new Rectangle(rect.Left + padding, (int)yRow2, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow2);
            g.DrawString("TENTAC(PO NO)", fontRegular, textBrush,
                         labelRectRow2.Left + 2, labelRectRow2.Top + (labelRectRow2.Height - fontRegular.Height) / 2);
            int totalValueWidth = (int)(rect.Width - titleWidth - 2 * padding);
            int partWidthRow2 = totalValueWidth / 10;
            Rectangle valueRectRow2 = new Rectangle(rect.Left + padding + titleWidth, (int)yRow2, totalValueWidth, (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow2.Left, valueRectRow2.Bottom, valueRectRow2.Right, valueRectRow2.Bottom);
            // 4 cột: 3/3/3/1
            Rectangle valueRectLeft = new Rectangle(valueRectRow2.Left, valueRectRow2.Top, partWidthRow2 * 3, valueRectRow2.Height);
            Rectangle valueRectCenter = new Rectangle(valueRectLeft.Right, valueRectRow2.Top, partWidthRow2 * 2, valueRectRow2.Height);
            Rectangle valueRectRight = new Rectangle(valueRectCenter.Right, valueRectRow2.Top, partWidthRow2 * 4, valueRectRow2.Height);
            Rectangle valueRectPCS = new Rectangle(valueRectRight.Right, valueRectRow2.Top, partWidthRow2 * 1, valueRectRow2.Height);
            // --- PoNo ---
            SizeF textSizePoNo = g.MeasureString(model.PoNo, fontBold);
            Font drawFontPoNo = fontBold;
            if (textSizePoNo.Width > valueRectLeft.Width - 10)
            {
                float shrinkRatio = (valueRectLeft.Width - 10) / textSizePoNo.Width;
                float newSize = (fontBold.Size * shrinkRatio) + 1;
                drawFontPoNo = new Font(fontBold.FontFamily, newSize, fontBold.Style);
            }
            float valueTextYPoNo = valueRectLeft.Top + (valueRectLeft.Height - drawFontPoNo.Height) / 2;
            g.DrawString(model.PoNo, drawFontPoNo, textBrush, valueRectLeft.Left + 5, valueTextYPoNo);
            // --- TTL Style ---
            float valueTextYStyleNo = valueRectCenter.Top + (valueRectCenter.Height - fontRegular.Height) / 2;
            g.DrawString("TTL Style", fontRegular, textBrush, valueRectCenter.Left + 5, valueTextYStyleNo);
            // --- Quantity ---
            string valueTextQuantity = model.Quantity.ToString("N0");
            SizeF textSizeQuantity = g.MeasureString(valueTextQuantity, fontRegular);
            Font drawFontQuantity = fontRegular;
            if (textSizeQuantity.Width > valueRectRight.Width - 10)
            {
                float shrinkRatio = (valueRectRight.Width - 10) / textSizeQuantity.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontQuantity = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            StringFormat sfRight = new StringFormat
            {
                Alignment = StringAlignment.Far,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString(valueTextQuantity, drawFontQuantity, textBrush, valueRectRight, sfRight);
            // --- PCS ---
            Font unitFont = new Font(fontRegular.FontFamily, 6, FontStyle.Regular);
            g.DrawString("PCS", unitFont, textBrush, valueRectPCS, sfRight);

            //----------------ITEM Code/Material----------------
            /*float yRow3 = rect.Top + padding + 2 * rowHeight + spacing * 2;
            Rectangle labelRectItemCode = new Rectangle(rect.Left + padding, (int)yRow3, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectItemCode);
            float labelTextItemCode = labelRectItemCode.Top + (labelRectItemCode.Height - fontRegular.Height) / 2;
            g.DrawString("ITEM Code/Material", fontRegular, textBrush, labelRectItemCode.Left + 2, labelTextItemCode);
            Rectangle valueRectItemCode = new Rectangle(rect.Left + padding + titleWidth, (int)yRow3,
                (int)(rect.Width - titleWidth - 2 * padding), (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectItemCode.Left, valueRectItemCode.Bottom, valueRectItemCode.Right, valueRectItemCode.Bottom);
            string valueTextItemCode = model.ItemCode + "/" + model.MaterialCode;
            SizeF textSizeItemCode = g.MeasureString(valueTextItemCode, fontRegular);
            Font drawFontItemCode = fontRegular;
            if (textSizeItemCode.Width > valueRectItemCode.Width - 10)
            {
                float shrinkRatio = (valueRectItemCode.Width - 10) / textSizeItemCode.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontItemCode = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            g.DrawString(valueTextItemCode, drawFontItemCode, textBrush, valueRectItemCode.Left + 2, labelTextItemCode);*/

            float yRow3 = rect.Top + padding + 2 * rowHeight + spacing * 2;
            Rectangle labelRectItemCode = new Rectangle(rect.Left + padding, (int)yRow3, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectItemCode);
            float labelTextItemCode = labelRectItemCode.Top + (labelRectItemCode.Height - fontRegular.Height) / 2;
            g.DrawString("ITEM Code/Material", fontRegular, textBrush, labelRectItemCode.Left + 2, labelTextItemCode);

            Rectangle valueRectItemCode = new Rectangle(rect.Left + padding + titleWidth, (int)yRow3,
                (int)(rect.Width - titleWidth - 2 * padding), (int)(rowHeight - spacing));

            g.DrawLine(borderPen, valueRectItemCode.Left, valueRectItemCode.Bottom, valueRectItemCode.Right, valueRectItemCode.Bottom);

            int partWidthItemCode = valueRectItemCode.Width / 10;
            Rectangle valueRectLeftRow3 = new Rectangle(valueRectItemCode.Left, valueRectItemCode.Top, partWidthItemCode * 8, valueRectItemCode.Height);
            Rectangle valueRectRightRow3 = new Rectangle(valueRectLeftRow3.Right, valueRectItemCode.Top, partWidthItemCode * 2, valueRectItemCode.Height);

            string valueTextItemCode = model.ItemCode + "/" + model.MaterialCode;
            SizeF textSizeItemCode = g.MeasureString(valueTextItemCode, fontRegular);
            Font drawFontItemCode = fontRegular;
            if (textSizeItemCode.Width > valueRectLeftRow3.Width - 10)
            {
                float shrinkRatio = (valueRectLeftRow3.Width - 10) / textSizeItemCode.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontItemCode = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            float valueTextYItemCode = valueRectLeftRow3.Top + (valueRectLeftRow3.Height - drawFontItemCode.Height) / 2;
            g.DrawString(valueTextItemCode, drawFontItemCode, textBrush, valueRectLeftRow3.Left + 5, valueTextYItemCode);

            // LotNo (căn phải)
            string lotNoText = model.LotNo;
            SizeF textSizeLotNo = g.MeasureString(lotNoText, fontRegular);
            Font drawFontLotNo = fontRegular;
            if (textSizeLotNo.Width > valueRectRightRow3.Width - 10)
            {
                float shrinkRatio = (valueRectRightRow3.Width - 10) / textSizeLotNo.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontLotNo = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }

            StringFormat sfLot = new StringFormat
            {
                Alignment = StringAlignment.Far,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString(lotNoText, drawFontLotNo, textBrush, valueRectRightRow3, sfLot);

            //----------------Style No----------------
            float yRow4 = rect.Top + padding + 3 * rowHeight + spacing * 3;
            Rectangle labelRectStyleNo = new Rectangle(rect.Left + padding, (int)yRow4, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectStyleNo);
            float labelTextYStyleNo = labelRectStyleNo.Top + (labelRectStyleNo.Height - fontRegular.Height) / 2;
            g.DrawString("Style No", fontRegular, textBrush, labelRectStyleNo.Left + 2, labelTextYStyleNo);
            Rectangle valueRectRow4 = new Rectangle(rect.Left + padding + titleWidth, (int)yRow4,
                (int)(rect.Width - titleWidth - 2 * padding), (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow4.Left, valueRectRow4.Bottom, valueRectRow4.Right, valueRectRow4.Bottom);
            // Chia thành 10 phần (4+2+4)
            int totalWidthRow4 = valueRectRow4.Width;
            int partWidthRow4 = totalWidthRow4 / 10;
            Rectangle rectPartStyleNo = new Rectangle(valueRectRow4.Left, valueRectRow4.Top, partWidthRow4 * 4, valueRectRow4.Height);
            Rectangle rectPartSize = new Rectangle(rectPartStyleNo.Right, valueRectRow4.Top, partWidthRow4 * 2, valueRectRow4.Height);
            Rectangle rectPartSizeName = new Rectangle(rectPartSize.Right, valueRectRow4.Top, partWidthRow4 * 4, valueRectRow4.Height);
            // Hàm vẽ chữ co font nếu vượt khung
            void DrawTextInRect(Graphics gr, Rectangle rectChill, string text, Font font, Brush brush, bool alignRight = false)
            {
                if (string.IsNullOrEmpty(text)) return;
                SizeF textSize3 = gr.MeasureString(text, font);
                Font drawFont3 = font;
                if (textSize3.Width > rectChill.Width - 6)
                {
                    float shrinkRatio = (rectChill.Width - 6) / textSize3.Width;
                    float newSize = font.Size * shrinkRatio;
                    drawFont3 = new Font(font.FontFamily, newSize, font.Style);
                }
                StringFormat sf3 = new StringFormat
                {
                    LineAlignment = StringAlignment.Center,
                    Alignment = alignRight ? StringAlignment.Far : StringAlignment.Near
                };
                gr.DrawString(text, drawFont3, brush, rectChill, sf3);
            }
            DrawTextInRect(g, rectPartStyleNo, "\u00A0" + model.StyleNo, fontRegular, textBrush);
            DrawTextInRect(g, rectPartSize, "Size", fontRegular, textBrush);
            DrawTextInRect(g, rectPartSizeName, "\u00A0" + model.Size, fontBold, textBrush, true);

            //----------------Color----------------
            float yRow5 = rect.Top + padding + 4 * rowHeight + spacing * 4;
            Rectangle labelRectColor = new Rectangle(rect.Left + padding, (int)yRow5, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectColor);
            float labelTextYColor = labelRectColor.Top + (labelRectColor.Height - fontRegular.Height) / 2;
            g.DrawString("Color", fontRegular, textBrush, labelRectColor.Left + 2, labelTextYColor);
            Rectangle valueRectRow5 = new Rectangle(rect.Left + padding + titleWidth, (int)yRow5,
                (int)(rect.Width - titleWidth - 2 * padding), (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow5.Left, valueRectRow5.Bottom, valueRectRow5.Right, valueRectRow5.Bottom);
            // Chia thành 10 phần (3+3+3+1)
            int totalWidthRow5 = valueRectRow5.Width;
            int partWidthRow5 = totalWidthRow5 / 10;
            Rectangle rectPartColorCode = new Rectangle(valueRectRow5.Left, valueRectRow5.Top, partWidthRow5 * 3, valueRectRow5.Height);
            Rectangle rectPartQTY = new Rectangle(rectPartColorCode.Right, valueRectRow5.Top, partWidthRow5 * 2, valueRectRow5.Height);
            Rectangle rectPartQuantity = new Rectangle(rectPartQTY.Right, valueRectRow5.Top, partWidthRow5 * 4, valueRectRow5.Height);
            Rectangle rectPartUnit = new Rectangle(rectPartQuantity.Right, valueRectRow5.Top, partWidthRow5 * 1, valueRectRow5.Height);
            string quantity = model.TypeProduct == 1
                ? model.QuantityInBox.ToString("N0") + "/" + model.SizeQuantity.ToString("N0")
                : model.SizeQuantity.ToString("N0");
            DrawTextInRect(g, rectPartColorCode, "\u00A0" + model.Color, fontRegular, textBrush);
            DrawTextInRect(g, rectPartQTY, "QTY", fontRegular, textBrush);
            DrawTextInRect(g, rectPartQuantity, quantity, fontBold, textBrush, true);
            // --- Unit (PCS) ---
            g.DrawString("PCS", unitFont, textBrush, rectPartUnit, sfRight);

            // --- Vùng dưới (1.9 cm) ---
            float bottomTop = rect.Top + padding + contentHeight + spacing * 5;
            float bottomHeight = rect.Height - contentHeight - 2 * padding;
            // Notice (chỉ chữ, không khung)
            string noticeText = "Please open and confirm this package immediately on receipt.";
            Font fontSmall = new Font(fontRegular.FontFamily, 6.5f, FontStyle.Regular);
            g.DrawString(noticeText, fontSmall, textBrush, rect.Left + padding, bottomTop);
            // Footer (chỉ chữ, không khung, đặt ngay dưới notice, khít lại)
            string footerText = "MADE IN VIETNAM";
            float footerTop = bottomTop + fontSmall.Height;
            g.DrawString(footerText, fontSmall, textBrush, rect.Left + padding, footerTop);
            // Hình ảnh + QR 
            float imageTop = footerTop + fontSmall.Height + 3;
            float imageHeight = rect.Bottom - padding - imageTop;
            int qrSize = 42;
            if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
            {
                using (System.Drawing.Image img = System.Drawing.Image.FromFile(imagePath))
                {
                    // Giữ tỉ lệ gốc của ảnh
                    float aspectRatio = (float)img.Width / img.Height;
                    // Chiều rộng tối đa: toàn bộ vùng còn lại trừ QR và padding
                    int maxWidth = (int)(rect.Width - padding * 2 - qrSize - 10);
                    // Tính kích thước ảnh theo tỉ lệ
                    int targetHeight = (int)imageHeight;
                    int targetWidth = (int)(targetHeight * aspectRatio);
                    // Nếu ảnh quá rộng thì thu lại theo maxWidth
                    if (targetWidth > maxWidth)
                    {
                        targetWidth = maxWidth;
                        targetHeight = (int)(targetWidth / aspectRatio);
                    }
                    Rectangle imgRect = new Rectangle(rect.Left + padding, (int)imageTop, targetWidth, targetHeight);
                    g.DrawImage(img, imgRect);
                }
            }
            // QR code: giữ vuông vắn, đặt cùng hàng với ảnh
            Rectangle qrRect = new Rectangle((int)(rect.Right - padding - qrSize), (int)imageTop, qrSize, qrSize);
            g.DrawRectangle(borderPen, qrRect);
            g.DrawString(string.Empty, fontRegular, textBrush,
                qrRect.Left + (qrRect.Width - g.MeasureString(string.Empty, fontRegular).Width) / 2,
                qrRect.Top + (qrRect.Height - fontRegular.Height) / 2);
        }

        /*private void printDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (printQueue == null || printQueue.Count == 0)
            {
                return;
            }

            int cols = 2, rows = 6;
            int temWidth = 345;   // ~8.7 cm
            int temHeight = 175;  // ~4.4 cm
            int paddingX = 8;    // khoảng cách ngang
            int paddingY = 8;    // khoảng cách dọc

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    int index = printedIndex + r * cols + c;
                    if (index >= printQueue.Count)
                    {
                        continue;
                    }

                    int x = e.MarginBounds.Left + c * (temWidth + paddingX);
                    int y = e.MarginBounds.Top + r * (temHeight + paddingY);

                    Rectangle rect = new Rectangle(x, y, temWidth, temHeight);

                    // Vẽ nội dung tem
                    DrawPrintLayout(e.Graphics, printQueue[index], rect);

                    // Vẽ crop marks (boong cắt) hở ra ngoài
                    int markLen = 15;
                    int offset = 4; // khoảng hở ra ngoài

                    // Trên - trái
                    e.Graphics.DrawLine(Pens.Gray, rect.Left - offset, rect.Top - offset,
                                        rect.Left - offset + markLen, rect.Top - offset);
                    e.Graphics.DrawLine(Pens.Gray, rect.Left - offset, rect.Top - offset,
                                        rect.Left - offset, rect.Top - offset + markLen);

                    // Trên - phải
                    e.Graphics.DrawLine(Pens.Gray, rect.Right + offset, rect.Top - offset,
                                        rect.Right + offset - markLen, rect.Top - offset);
                    e.Graphics.DrawLine(Pens.Gray, rect.Right + offset, rect.Top - offset,
                                        rect.Right + offset, rect.Top - offset + markLen);

                    // Dưới - trái
                    e.Graphics.DrawLine(Pens.Gray, rect.Left - offset, rect.Bottom + offset,
                                        rect.Left - offset + markLen, rect.Bottom + offset);
                    e.Graphics.DrawLine(Pens.Gray, rect.Left - offset, rect.Bottom + offset,
                                        rect.Left - offset, rect.Bottom + offset - markLen);

                    // Dưới - phải
                    e.Graphics.DrawLine(Pens.Gray, rect.Right + offset, rect.Bottom + offset,
                                        rect.Right + offset - markLen, rect.Bottom + offset);
                    e.Graphics.DrawLine(Pens.Gray, rect.Right + offset, rect.Bottom + offset,
                                        rect.Right + offset, rect.Bottom + offset - markLen);
                }
            }

            printedIndex += cols * rows;
            e.HasMorePages = printedIndex < printQueue.Count;
        }*/

        //private void printDocument_PrintPage(object sender, PrintPageEventArgs e)
        //{
        //    if (printQueue == null || printQueue.Count == 0)
        //    {
        //        return;
        //    }

        //    int cols = 2;
        //    int rows = 4;

        //    // Kích thước tem
        //    int temWidth = 345;   // ~8.8 cm
        //    int temHeight = 175;  // ~4.4 cm

        //    // Khoảng cách dọc giữa các tem
        //    //int paddingY = 8; 
        //    int gapHalfY = 34;
        //    int paddingY = gapHalfY * 2;

        //    // Lấy chiều ngang thật của trang giấy
        //    // KHÔNG dùng e.MarginBounds.Left vì nó cộng thêm lề mặc định của máy in
        //    int pageLeft = e.PageBounds.Left;
        //    int pageTop = e.PageBounds.Top;
        //    int pageWidth = e.PageBounds.Width;

        //    // Tính phần dư ngang còn lại
        //    int totalTemWidth = cols * temWidth;
        //    int remainWidth = pageWidth - totalTemWidth;

        //    // Chia lại phần dư:
        //    // Trái = Phải = 1 phần
        //    // Giữa = 2 phần
        //    int sideMargin = remainWidth / 4;
        //    int middleGap = remainWidth / 2;

        //    // Nếu muốn cố định đúng theo số đo đã tính, có thể dùng:
        //    // int sideMargin = 34;
        //    // int middleGap = 69;

        //    // Chừa thêm một khoảng nhỏ phía trên để boong cắt top không bị sát mép giấy    
        //    // Boong cắt top sẽ được tính từ mép trên của tem đi lên gapHalfY
        //    // topCutSafeOffset = 90 sẽ canh đều footer và top phần khoảng cách từ mép giấy đến boong cắt
        //    //int topCutSafeOffset = 90;
        //    int topCutSafeOffset = 0;
        //    int topMargin = gapHalfY + topCutSafeOffset;

        //    // Điều chỉnh thủ công2
        //    // Điều chỉnh X âm di chuyển tất cả các nhãn sang trái
        //    // Điều chỉnh X dương di chuyển tất cả các nhãn sang phải
        //    int adjustX = -6;

        //    // Điều chỉnh khoảng cách giữa bằng tay
        //    // Giá trị dương làm cho khoảng cách giữa rộng hơn
        //    // Giá trị âm làm cho khoảng cách giữa hẹp hơn
        //    int adjustMiddleGap = 0;

        //    // Đường cắt giữa để chia tờ giấy A4 theo chiều dọc
        //    // Vị trí này là trung tâm của khoảng trống giữa nhãn bên trái và nhãn bên phải
        //    int leftColumnX = pageLeft + sideMargin + adjustX;
        //    int rightColumnX = pageLeft + sideMargin + temWidth + middleGap + adjustMiddleGap + adjustX;
        //    int centerCutX = leftColumnX + temWidth + ((middleGap + adjustMiddleGap) / 2);

        //    // Tính toán số hàng thực tế trên trang hiện tại.
        //    int remainingItems = printQueue.Count - printedIndex;
        //    int itemsOnThisPage = Math.Min(cols * rows, remainingItems);
        //    int actualRows = (itemsOnThisPage + cols - 1) / cols;

        //    for (int r = 0; r < rows; r++)
        //    {
        //        for (int c = 0; c < cols; c++)
        //        {
        //            int index = printedIndex + r * cols + c;

        //            if (index >= printQueue.Count)
        //            {
        //                continue;
        //            }

        //            int x;

        //            if (c == 0)
        //            {
        //                // Cột trái
        //                //x = pageLeft + sideMargin;
        //                x = pageLeft + sideMargin + adjustX;
        //            }
        //            else
        //            {
        //                // Cột phải
        //                //x = pageLeft + sideMargin + temWidth + middleGap;
        //                x = pageLeft + sideMargin + temWidth + middleGap + adjustMiddleGap + adjustX;
        //            }

        //            int y = pageTop + topMargin + r * (temHeight + paddingY);

        //            Rectangle rect = new Rectangle(x, y, temWidth, temHeight);

        //            // Vẽ nội dung tem
        //            DrawPrintLayout(e.Graphics, printQueue[index], rect);
        //        }
        //    }

        //    using (Pen cutPen = new Pen(Color.FromArgb(200, 200, 200), 1))
        //    {
        //        cutPen.DashStyle = DashStyle.Custom;
        //        cutPen.DashPattern = new float[] { 10, 7 };
        //        cutPen.DashCap = DashCap.Flat;

        //        if (actualRows > 0)
        //        {
        //            //Vị trí trên cùng và dưới cùng của toàn bộ vùng nhãn
        //            int firstRowTop = pageTop + topMargin;
        //            int lastRowTop = pageTop + topMargin + (actualRows - 1) * (temHeight + paddingY);
        //            int lastRowBottom = lastRowTop + temHeight;

        //            // Boong cắt top được đo từ mép trên của tem đi lên
        //            int topCutY = firstRowTop - gapHalfY;

        //            // Boong cắt footer được đo từ mép dưới của tem cuối đi xuống
        //            int footerCutY = lastRowBottom + gapHalfY;

        //            // Cho đường boong dài hơn một chút so với vùng tem
        //            int cutLineExtend = 10;

        //            int fullCutStartX = leftColumnX - cutLineExtend;
        //            int fullCutEndX = rightColumnX + temWidth + cutLineExtend;

        //            // Vẽ đường cắt ngang phía trên, chạy dài hết vùng tem
        //            e.Graphics.DrawLine(cutPen, fullCutStartX, topCutY, fullCutEndX, topCutY);

        //            // Vẽ đường cắt dọc ở giữa GAP, chạy từ boong top tới boong footer
        //            e.Graphics.DrawLine(cutPen, centerCutX, topCutY, centerCutX, footerCutY);

        //            // Vẽ các đường cắt ngang giữa các hàng
        //            // Vị trí cắt là tâm của khoảng trống dọc 0,85 cm x 2
        //            for (int r = 0; r < actualRows - 1; r++)
        //            {
        //                int currentRowTop = pageTop + topMargin + r * (temHeight + paddingY);
        //                int cutY = currentRowTop + temHeight + gapHalfY;

        //                e.Graphics.DrawLine(cutPen, fullCutStartX, cutY, fullCutEndX, cutY);
        //            }

        //            // Vẽ dấu cắt chân trang
        //            // Vị trí cắt chân trang cũng dựa trên khoảng cách 0,85 cm
        //            e.Graphics.DrawLine(cutPen, fullCutStartX, footerCutY, fullCutEndX, footerCutY);
        //        }
        //    }

        //    printedIndex += cols * rows;
        //    e.HasMorePages = printedIndex < printQueue.Count;
        //}

        private void printDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (printQueue == null || printQueue.Count == 0)
            {
                return;
            }

            int cols = 2;
            int rows = 6;

            // Kích thước tem 
            int temWidth = 345;   // ~8.8 cm
            int temHeight = 175;  // ~4.4 cm

            // Layout: 6 hàng, 2 cột
            int paddingX = 12;     // khoảng cách ngang giữa 2 tem
            int paddingY = 12;     // khoảng cách dọc giữa các hàng

            // Dùng layout theo MarginBounds
            int startX = e.MarginBounds.Left;
            int startY = e.MarginBounds.Top;

            // Điều chỉnh thủ công nếu cần
            int adjustX = 0;
            int adjustY = 0;

            startX += adjustX;
            startY += adjustY;

            int remainingItems = printQueue.Count - printedIndex;
            int itemsOnThisPage = Math.Min(cols * rows, remainingItems);
            int actualRows = (itemsOnThisPage + cols - 1) / cols;

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    int index = printedIndex + r * cols + c;

                    if (index >= printQueue.Count)
                    {
                        continue;
                    }

                    int x = startX + c * (temWidth + paddingX);
                    int y = startY + r * (temHeight + paddingY);

                    Rectangle rect = new Rectangle(x, y, temWidth, temHeight);

                    // Vẽ nội dung tem
                    DrawPrintLayout(e.Graphics, printQueue[index], rect);
                }
            }

            // Vẽ boong / đường cắt
            using (Pen cutPen = new Pen(Color.FromArgb(200, 200, 200), 1))
            {
                cutPen.DashStyle = DashStyle.Custom;
                cutPen.DashPattern = new float[] { 10, 7 };
                cutPen.DashCap = DashCap.Flat;

                if (actualRows > 0)
                {
                    int gapHalfX = paddingX / 2;
                    int gapHalfY = paddingY / 2;

                    int firstRowTop = startY;
                    int lastRowTop = startY + (actualRows - 1) * (temHeight + paddingY);
                    int lastRowBottom = lastRowTop + temHeight;

                    int leftColumnX = startX;
                    int rightColumnX = startX + temWidth + paddingX;

                    // Vị trí đường cắt
                    int topCutY = firstRowTop - gapHalfY;
                    int footerCutY = lastRowBottom + gapHalfY;

                    int centerCutX = startX + temWidth + gapHalfX;

                    // Thêm boong dọc 2 bên trái/phải
                    // Đặt cách mép tem ra ngoài một nửa khoảng padding
                    int leftCutX = leftColumnX - gapHalfX;
                    int rightCutX = rightColumnX + temWidth + gapHalfX;

                    // Cho đường boong ngang dài ra ngoài một chút
                    int cutLineExtend = 10;

                    int fullCutStartX = leftCutX - cutLineExtend;
                    int fullCutEndX = rightCutX + cutLineExtend;

                    // 1. Đường cắt ngang phía trên
                    e.Graphics.DrawLine(cutPen, fullCutStartX, topCutY, fullCutEndX, topCutY);

                    // 2. Đường cắt dọc bên trái ngoài cùng
                    e.Graphics.DrawLine(cutPen, leftCutX, topCutY, leftCutX, footerCutY);

                    // 3. Đường cắt dọc ở giữa 2 cột
                    e.Graphics.DrawLine(cutPen, centerCutX, topCutY, centerCutX, footerCutY);

                    // 4. Đường cắt dọc bên phải ngoài cùng
                    e.Graphics.DrawLine(cutPen, rightCutX, topCutY, rightCutX, footerCutY);

                    // 5. Các đường cắt ngang giữa các hàng
                    for (int r = 0; r < actualRows - 1; r++)
                    {
                        int currentRowTop = startY + r * (temHeight + paddingY);
                        int cutY = currentRowTop + temHeight + gapHalfY;

                        e.Graphics.DrawLine(cutPen, fullCutStartX, cutY, fullCutEndX, cutY);
                    }

                    // 6. Đường cắt ngang phía dưới
                    e.Graphics.DrawLine(cutPen, fullCutStartX, footerCutY, fullCutEndX, footerCutY);
                }
            }

            printedIndex += cols * rows;
            e.HasMorePages = printedIndex < printQueue.Count;
        }

        private void FrmPackingSealMontbell_Load(object sender, EventArgs e)
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
        }

        private void BuildPrintQueue()
        {
            printQueue = new List<MontbellModel>();
            foreach (DataGridViewRow row in dgvData.Rows)
            {
                if (row.IsNewRow) continue;
                var newModel = new MontbellModel
                {
                    OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty,
                    PoNo = row.Cells["PONo"].Value?.ToString() ?? string.Empty,
                    ItemCode = row.Cells["ItemCode"].Value?.ToString() ?? string.Empty,
                    MaterialCode = row.Cells["MaterialCode"].Value?.ToString() ?? string.Empty,
                    StyleNo = row.Cells["StyleNo"].Value?.ToString() ?? string.Empty,
                    Color = row.Cells["ColorName"].Value?.ToString() ?? string.Empty,
                    Size = row.Cells["SizeName"].Value?.ToString() ?? string.Empty,
                    TypeProduct = typeProduct,
                    Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString()),
                    SizeQuantity = string.IsNullOrEmpty(row.Cells["SizeQuantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["SizeQuantity"].Value.ToString()),
                    QuantityInBox = string.IsNullOrEmpty(row.Cells["QuantityInBox"].Value.ToString()) ? 0 : double.Parse(row.Cells["QuantityInBox"].Value.ToString()),
                    LotNo = row.Cells["LotNo"].Value.ToString() ?? string.Empty,
                    Index = Convert.ToInt32(row.Cells["PCS"].Value)
                };

                // Lặp theo PCS để add đủ số tem
                for (int i = 0; i < newModel.Index; i++)
                {
                    printQueue.Add(new MontbellModel
                    {
                        OrderNo = newModel.OrderNo,
                        PoNo = newModel.PoNo,
                        ItemCode = newModel.ItemCode,
                        MaterialCode = newModel.MaterialCode,
                        StyleNo = newModel.StyleNo,
                        Color = newModel.Color,
                        Size = newModel.Size,
                        TypeProduct = newModel.TypeProduct,
                        Quantity = newModel.Quantity,
                        SizeQuantity = newModel.SizeQuantity,
                        QuantityInBox = newModel.QuantityInBox,
                        LotNo = newModel.LotNo,
                        Index = newModel.Index
                    });
                }
            }
            printedIndex = 0;
        }

        /*private void btnPrint_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            if (dgvData.Rows.Count == 0)
            {
                MessageBox.Show("Please enter data before printing.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Cursor = Cursors.Default;
                return;
            }
            try
            {
                BuildPrintQueue(); // tạo danh sách tem từ GridView
                printedIndex = 0; // reset index in queue

                printDocument.DefaultPageSettings.PaperSize = new System.Drawing.Printing.PaperSize("A4", 827, 1169);
                printDocument.DefaultPageSettings.Margins = new Margins(55, 0, 30, 0);

                if (Common.Consts.SettingINI.OnlyPrint != "True" && Common.Consts.SettingINI.OutputAndPrint != "True")
                {
                    if (string.IsNullOrEmpty(Common.Consts.SettingINI.PDFOutputFolderPath))
                    {
                        Common.Consts.SettingINI.PDFOutputFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    }
                    pdfOutputPath = Path.Combine(Common.Consts.SettingINI.PDFOutputFolderPath,
                        $"ONEGAI_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");

                    printDocument.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                    printDocument.PrinterSettings.PrintToFile = true;
                    printDocument.PrinterSettings.PrintFileName = pdfOutputPath;

                    printDocument.Print();
                }
                else if (Common.Consts.SettingINI.OnlyPrint == "True")
                {
                    printDocument.PrinterSettings.PrinterName = Common.Consts.SettingINI.PrinterName;
                    printDocument.PrinterSettings.PrintToFile = false;
                    printDocument.Print();
                }
                else if (Common.Consts.SettingINI.OutputAndPrint == "True")
                {
                    if (string.IsNullOrEmpty(Common.Consts.SettingINI.PDFOutputFolderPath))
                    {
                        Common.Consts.SettingINI.PDFOutputFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    }
                    pdfOutputPath = Path.Combine(Common.Consts.SettingINI.PDFOutputFolderPath,
                        $"ONEGAI_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");

                    printDocument.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                    printDocument.PrinterSettings.PrintToFile = true;
                    printDocument.PrinterSettings.PrintFileName = pdfOutputPath;
                    printDocument.Print();

                    Utils.PDFSetting.PrintPdfFileWithPDFViewer(pdfOutputPath, Common.Consts.SettingINI.PrinterName);
                }
                AutoCloseMessageBox.ShowMessage($"{Environment.MachineName} is print command has been successfully sent to the printer {Common.Consts.SettingINI.PrinterName}.", 2000);
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cursor = Cursors.Default;
                return;
            }
        }*/

        private void btnPrint_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            if (dgvData.Rows.Count == 0)
            {
                MessageBox.Show("Please enter data before printing.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Cursor = Cursors.Default;
                return;
            }

            string pdfPath = null;
            bool isTempPdf = false;

            try
            {
                // Case 1:
                // OnlyPrint = False, OutputAndPrint = False
                // Export PDF only. Keep PDF file.

                // Case 2:
                // OnlyPrint = True, OutputAndPrint = False
                // Export PDF to temp folder. Print PDF. Delete temp PDF.

                // Case 3:
                // OutputAndPrint = True
                // Export PDF to output folder. Print PDF. Keep PDF file.

                bool onlyPrint = Common.Consts.SettingINI.OnlyPrint == "True";
                bool outputAndPrint = Common.Consts.SettingINI.OutputAndPrint == "True";

                bool needPrint = onlyPrint == true || outputAndPrint == true;
                bool needKeepPdf = onlyPrint == false || outputAndPrint == true;

                // Build data before export or print
                BuildPrintQueue();
                printedIndex = 0;

                // Setup A4 portrait
                // PaperSize unit is 1/100 inch
                printDocument.DefaultPageSettings.PaperSize = new System.Drawing.Printing.PaperSize("A4", 827, 1169);

                // Keep current margins because PDF layout is already close
                printDocument.DefaultPageSettings.Margins = new Margins(55, 0, 30, 0);
                printDocument.DefaultPageSettings.Landscape = false;
                printDocument.OriginAtMargins = false;

                if (needKeepPdf)
                {
                    // Export PDF to output folder
                    // If output folder is empty, use Desktop
                    string outputFolder = Common.Consts.SettingINI.PDFOutputFolderPath;

                    if (string.IsNullOrEmpty(outputFolder))
                    {
                        outputFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        Common.Consts.SettingINI.PDFOutputFolderPath = outputFolder;
                    }

                    if (!Directory.Exists(outputFolder))
                    {
                        Directory.CreateDirectory(outputFolder);
                    }

                    pdfPath = Path.Combine(
                        outputFolder,
                        "ONEGAI_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".pdf");

                    isTempPdf = false;
                }
                else
                {
                    // Only print mode
                    // Create PDF in temp folder and delete after print
                    pdfPath = Path.Combine(Path.GetTempPath(),
                        "ONEGAI_PRINT_" + DateTime.Now.ToString("yyyyMMdd_HHmmssfff") + ".pdf");

                    isTempPdf = true;
                }

                pdfOutputPath = pdfPath;

                // Always create PDF first
                // This makes printing more stable than direct PrintDocument to printer
                printDocument.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                printDocument.PrinterSettings.PrintToFile = true;
                printDocument.PrinterSettings.PrintFileName = pdfPath;

                printDocument.Print();

                // Wait until PDF file is created
                Utils.PDFSetting.WaitUntilPdfReady(pdfPath, 30000);
                Thread.Sleep(500);
                // Print only when required
                // If user selected nothing, this block will not run
                if (needPrint)
                {
                    Utils.PDFSetting.PrintPdfFileWithPDFViewer(pdfPath, Common.Consts.SettingINI.PrinterName);
                }

                // Delete temp PDF only for only print mode
                //if (isTempPdf)
                //{
                //    DeleteFileSilent(pdfPath, 5000);
                //}

                string message;

                if (!needPrint)
                {
                    message = "PDF has been created successfully:\n" + pdfPath;
                }
                else if (outputAndPrint)
                {
                    message = "PDF has been created and print command has been sent to the printer "
                        + Common.Consts.SettingINI.PrinterName
                        + ".\n"
                        + pdfPath;
                }
                else
                {
                    message = Environment.MachineName
                        + " print command has been successfully sent to the printer "
                        + Common.Consts.SettingINI.PrinterName
                        + ".";
                }

                AutoCloseMessageBox.ShowMessage(message, 2000);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void WaitUntilPdfCreated(string filePath, int timeoutMilliseconds)
        {
            DateTime startTime = DateTime.Now;
            long lastLength = -1;
            int stableCount = 0;

            while (true)
            {
                try
                {
                    if (File.Exists(filePath))
                    {
                        FileInfo fileInfo = new FileInfo(filePath);

                        if (fileInfo.Length > 0)
                        {
                            // Microsoft Print to PDF may still be writing the file
                            // Check file length until it is stable
                            if (fileInfo.Length == lastLength)
                            {
                                stableCount++;
                            }
                            else
                            {
                                stableCount = 0;
                                lastLength = fileInfo.Length;
                            }

                            if (stableCount >= 2)
                            {
                                return;
                            }
                        }
                    }
                }
                catch
                {
                    // Ignore temporary access errors
                }

                if ((DateTime.Now - startTime).TotalMilliseconds > timeoutMilliseconds)
                {
                    throw new Exception("PDF file was not created or is not ready: " + filePath);
                }

                System.Threading.Thread.Sleep(300);
                Application.DoEvents();
            }
        }

        private void DeleteFileSilent(string filePath, int timeoutMilliseconds)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                return;
            }

            DateTime startTime = DateTime.Now;

            while (true)
            {
                try
                {
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }

                    return;
                }
                catch
                {
                    // Ignore delete error for temp PDF
                    if ((DateTime.Now - startTime).TotalMilliseconds > timeoutMilliseconds)
                    {
                        return;
                    }

                    System.Threading.Thread.Sleep(300);
                    Application.DoEvents();
                }
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
            zoomFactor -= 0.1f;
            if (zoomFactor < 1.0f)
            {
                zoomFactor = 1.0f;
            }
            panel12.Invalidate();
            ResizeRuler(zoomFactor);
        }

        private void toolStripButtonZoomIn_Click(object sender, EventArgs e)
        {
            zoomFactor += 0.1f;
            panel12.Invalidate();
            ResizeRuler(zoomFactor);
        }

        private void toolStripButtonReset_Click(object sender, EventArgs e)
        {
            zoomFactor = 1.0f;
            panel12.Invalidate();
            ResizeRuler(zoomFactor);
        }

        private void ResizeRuler(double ZoomFactor)
        {
            verticalRuler1.ZoomFactor = ZoomFactor;
            verticalRuler1.Invalidate();
            horizontalRuler1.ZoomFactor = zoomFactor;
            horizontalRuler1.Invalidate();
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
                btnPrint_Click(sender, e);
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
            dataTable.Columns.Add("LotNo");
            dataTable.Columns.Add("PCS");
            dataTable.Columns.Add("Key");
        }

        private void ShowImportNotification(ImportResponseModel resModel)
        {
            try
            {
                if (notifyIcon1 == null)
                {
                    notifyIcon1 = new NotifyIcon();
                }

                string status = resModel.IsSuccess ? "Success" : "Failed";

                notifyIcon1.BalloonTipTitle = "Data Import Notification";
                notifyIcon1.Icon = this.Icon; // nhớ gán icon cho form trước
                notifyIcon1.Visible = true;
                notifyIcon1.BalloonTipText =
                    $"Status: {status}\n" +
                    $"File: {resModel.SourceName}\n" +
                    $"Success: {resModel.RecordsImported}, Error: {resModel.RecordsFailed}\n" +
                    $"{resModel.Message}";

                notifyIcon1.ShowBalloonTip(4000);
                // KHÔNG dispose ở đây, chỉ dispose khi đóng form
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
                    //int boxSize = 3200;
                    int boxSize = Convert.ToInt32(comboBoxSize.SelectedValue.ToString());
                    if (boxSize == 0)
                    {
                        dgvData.Rows[rowIndex].Cells["SizeQuantity"].Value = quantity.ToString("N0");
                        dgvData.Rows[rowIndex].Cells["QuantityInBox"].Value = quantity.ToString("N0");
                        dgvData.Rows[rowIndex].Cells["LotNo"].Value = "1/1";
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
                    int totalBoxes = boxQuantities.Count;
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
                    dgvData.Rows[rowIndex].Cells["LotNo"].Value = $"1/{totalBoxes}";

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
                        newRow["LotNo"] = $"{i + 1}/{totalBoxes}";
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
            comboBoxSize.SelectedValue = "0";
            comboBoxSize.Enabled = false;
        }

        private void comboBoxSize_SelectionChangeCommitted(object sender, EventArgs e)
        {
            // Nếu chọn Size = 0 (All in one box) thì không cần gom lại
            if (comboBoxSize.SelectedValue.ToString() == "0" && originalData != null)
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
                    newRow["LotNo"] = "1/1";
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
                    int totalBoxes = fullBoxes + (lastBoxQty > 0 ? 1 : 0);

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
                        newRow["SizeQuantity"] = sizeQuantity.ToString("N0");
                        newRow["Quantity"] = quantitykeyProductItemCode;
                        newRow["PCS"] = "1";
                        newRow["LotNo"] = $"{i + 1}/{totalBoxes}";
                        newRow["Key"] = key;
                        if (dataTable.Columns.Contains("QuantityInBox"))
                        {
                            newRow["QuantityInBox"] = QuantityInboxValue.ToString("N0");
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
                        newRow["SizeQuantity"] = sizeQuantity.ToString("N0");
                        newRow["Quantity"] = quantitykeyProductItemCode;
                        newRow["LotNo"] = $"{totalBoxes}/{totalBoxes}";
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

        private void panel12_Paint(object sender, PaintEventArgs e)
        {
            if (currentModel != null)
            {
                e.Graphics.ScaleTransform(zoomFactor, zoomFactor);
                DrawSingleLayout(e.Graphics, currentModel);
            }

            if (layoutCache != null)
            {
                e.Graphics.DrawImage(layoutCache, Point.Empty);
            }

            if (!crosshairPoint.IsEmpty)
            {
                using (Pen pen = new Pen(Color.DarkGreen, 1))
                {
                    e.Graphics.ResetTransform();

                    e.Graphics.DrawLine(pen, 0, crosshairPoint.Y, panel12.Width, crosshairPoint.Y);
                    e.Graphics.DrawLine(pen, crosshairPoint.X, 0, crosshairPoint.X, panel12.Height);
                }
            }
        }

        private void panel12_MouseMove(object sender, MouseEventArgs e)
        {
            crosshairPoint = e.Location;
            panel12.Invalidate();
        }
    }
}
