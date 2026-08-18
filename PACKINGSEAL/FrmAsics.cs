using Microsoft.VisualBasic.FileIO;
using PACKINGSEAL.Models;
using PACKINGSEAL.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Deployment.Application;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Media.Media3D;

namespace PACKINGSEAL
{
    public partial class FrmAsics : Form
    {
        public string imagePath = string.Empty;
        private string pdfOutputPath = string.Empty;
        private Assembly assembly = Assembly.GetExecutingAssembly();
        private PrintDocument printDocument = new PrintDocument();
        private AsicsModel currentModel;
        private int printedIndex = 0;
        private float zoomFactor = 1.0f;
        private List<AsicsModel> printQueue;
        private Point crosshairPoint = Point.Empty;
        private Bitmap layoutCache;

        public FrmAsics(FilePathModel filePathModel)
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
            printDocument = new PrintDocument();
            printDocument.PrintPage += printDocument_PrintPage;
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

        private void LoadDataGirdview(string filePath)
        {
            try
            {
                AppCommonModule appCommon = new AppCommonModule();
                DataTable dt = new DataTable();
                dt = new DataTable();
                dt.Columns.Add("Customer");
                dt.Columns.Add("OrderNo");
                dt.Columns.Add("POLine");
                dt.Columns.Add("SKUNo");
                dt.Columns.Add("ItemCode");
                dt.Columns.Add("ProductName");
                dt.Columns.Add("Unit");
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
                                string unit = appCommon.GetUnitByItemCode(fields[1].Trim().Replace("\"", ""));
                                dataRow["Unit"] = unit;
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
            currentModel = new AsicsModel();
            currentModel.ImagePath = imagePath;
            currentModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
            currentModel.Customer = row.Cells["Customer"].Value?.ToString() ?? string.Empty;
            currentModel.POLine = row.Cells["POLine"].Value?.ToString() ?? string.Empty;
            string SkuNo = row.Cells["SKUNo"].Value.ToString() ?? string.Empty;
            if (!SkuNo.StartsWith("-") && !string.IsNullOrEmpty(SkuNo))
            {
                SkuNo = "-" + SkuNo;
            }
            currentModel.SKUNo = SkuNo;
            currentModel.OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty;
            currentModel.SampleNo = row.Cells["SampleNo"].Value?.ToString() ?? string.Empty;
            currentModel.ItemCode = row.Cells["ItemCode"].Value?.ToString() ?? string.Empty;
            currentModel.Unit = row.Cells["Unit"].Value?.ToString() ?? "PCS";
            string sizeName = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
            if (!sizeName.StartsWith(".") && !string.IsNullOrEmpty(sizeName))
            {
                sizeName = "." + sizeName;
            }
            currentModel.SizeName = sizeName;
            string colorCode = row.Cells["ColorCode"].Value?.ToString() ?? string.Empty;
            if (!colorCode.StartsWith(".") && !string.IsNullOrEmpty(colorCode))
            {
                colorCode = "." + colorCode;
            }
            currentModel.ColorCode = colorCode;
            currentModel.Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString());
            currentModel.KeyNo = row.Cells["KeyNo"].Value?.ToString() ?? string.Empty;
            currentModel.ContractNo = row.Cells["ContractNo"].Value?.ToString() ?? string.Empty;
            currentModel.Vendors = row.Cells["Vendors"].Value?.ToString() ?? string.Empty;
            currentModel.ProductName = row.Cells["ProductName"].Value?.ToString() ?? string.Empty;
            currentModel.Unit = row.Cells["Unit"].Value?.ToString() ?? "PCS";

            printPreviewControl1.Document = printDocument;
            printPreviewControl1.InvalidatePreview();
            panel12.Invalidate();
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

        //2026/07/30 Thay đổi thành in theo layout 8 con tem có khoảng cách 0.875cm cho từng Itemcode
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

        private void DrawPrintLayout(Graphics g, AsicsModel model, Rectangle rect)
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
            float contentHeight = 92; // vùng nội dung
            float rowHeight = contentHeight / 5f;
            int spacing = 2;
            int titleWidth = 75;

            // ---------------- Customer ----------------
            float yRow1 = rect.Top + padding;
            Rectangle labelRectRow1 = new Rectangle(rect.Left + padding, (int)yRow1, titleWidth, (int)(rowHeight - 2));
            g.DrawRectangle(borderPen, labelRectRow1);
            g.DrawString("Customer", fontRegular, textBrush,
                             labelRectRow1.Left + 2, labelRectRow1.Top + (labelRectRow1.Height - fontRegular.Height) / 2);

            Rectangle valueRectCustomer = new Rectangle(rect.Left + padding + titleWidth, (int)yRow1,
                rect.Width - titleWidth - 2 * padding, (int)(rowHeight - 2));
            g.DrawLine(borderPen, valueRectCustomer.Left, valueRectCustomer.Bottom,
                       valueRectCustomer.Right, valueRectCustomer.Bottom);

            string textCustomer = model.Customer;
            SizeF sizeCustomer = g.MeasureString(textCustomer, fontRegular);
            Font fontCustomer = fontRegular;
            if (sizeCustomer.Width > valueRectCustomer.Width - 5)
            {
                float shrink = (valueRectCustomer.Width - 5) / sizeCustomer.Width;
                fontCustomer = new Font(fontRegular.FontFamily, fontRegular.Size * shrink, fontRegular.Style);
            }
            g.DrawString(textCustomer, fontCustomer, textBrush,
                         valueRectCustomer.Left + 5,
                         valueRectCustomer.Top + (valueRectCustomer.Height - fontCustomer.Height) / 2);

            // ---------------- Order No (3-3-4) ----------------
            float yRow2 = rect.Top + padding + rowHeight + spacing; // ngay sau Customer

            // Label
            Rectangle labelRectRow2 = new Rectangle(rect.Left + padding, (int)yRow2, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow2);
            g.DrawString("Order No", fontRegular, textBrush,
                         labelRectRow2.Left + 2,
                         labelRectRow2.Top + (labelRectRow2.Height - fontRegular.Height) / 2);

            // Value rectangle
            int totalValueWidth = rect.Width - titleWidth - 2 * padding;
            Rectangle valueRectRow2 = new Rectangle(rect.Left + padding + titleWidth, (int)yRow2,
                                                    totalValueWidth, (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow2.Left, valueRectRow2.Bottom, valueRectRow2.Right, valueRectRow2.Bottom);

            // Chia 3 phần: 3-3-4
            int partWidthRow2 = totalValueWidth / 10;
            Rectangle valueRectLeft = new Rectangle(valueRectRow2.Left, valueRectRow2.Top, partWidthRow2 * 3, valueRectRow2.Height);
            Rectangle valueRectCenter = new Rectangle(valueRectLeft.Right, valueRectRow2.Top, partWidthRow2 * 3, valueRectRow2.Height);
            Rectangle valueRectRight = new Rectangle(valueRectCenter.Right, valueRectRow2.Top, partWidthRow2 * 4, valueRectRow2.Height);

            // Order No (căn trái)
            SizeF textSizeOrderNo = g.MeasureString(model.OrderNo, fontBold);
            Font drawFontOrderNo = fontBold;
            if (textSizeOrderNo.Width > valueRectLeft.Width - 10)
            {
                float shrinkRatio = (valueRectLeft.Width - 10) / textSizeOrderNo.Width;
                float newSize = fontBold.Size * shrinkRatio;
                drawFontOrderNo = new Font(fontBold.FontFamily, newSize, fontBold.Style);
            }
            float valueTextYOrderNo = valueRectLeft.Top + (valueRectLeft.Height - drawFontOrderNo.Height) / 2;
            g.DrawString(model.OrderNo, drawFontOrderNo, textBrush, valueRectLeft.Left + 5, valueTextYOrderNo);

            // SKU (căn giữa)
            string valueTextSKU = model.SKUNo;
            SizeF textSizeSKU = g.MeasureString(valueTextSKU, fontRegular);
            Font drawFontSKU = fontRegular;
            if (textSizeSKU.Width > valueRectCenter.Width - 10)
            {
                float shrinkRatio = (valueRectCenter.Width - 10) / textSizeSKU.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontSKU = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            StringFormat sfCenter = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString(valueTextSKU, drawFontSKU, textBrush, valueRectCenter, sfCenter);

            // Contract No (căn phải)
            string valueTextContractNo = model.ContractNo;
            SizeF textSizeContractNo = g.MeasureString(valueTextContractNo, fontRegular);
            Font drawFontContractNo = fontRegular;
            if (textSizeContractNo.Width > valueRectRight.Width - 10)
            {
                float shrinkRatio = (valueRectRight.Width - 10) / textSizeContractNo.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontContractNo = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            StringFormat sfRight = new StringFormat
            {
                Alignment = StringAlignment.Far,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString(valueTextContractNo, drawFontContractNo, textBrush, valueRectRight, sfRight);

            // ---------------- Item Code (gộp dòng 2 + 3) ----------------
            float yRow3 = rect.Top + padding + 2 * rowHeight + spacing * 2;

            // Label (cao bằng 2 dòng)
            Rectangle labelRectItemCode = new Rectangle(
                rect.Left + padding,
                (int)yRow3,
                titleWidth,
                (int)(2 * rowHeight - spacing) // gộp 2 dòng
            );
            g.DrawRectangle(borderPen, labelRectItemCode);

            // Chữ "Item Code" căn xuống đáy ô label
            float labelTextItemCode = labelRectItemCode.Bottom - fontRegular.Height;
            g.DrawString("Item Code", fontRegular, textBrush, labelRectItemCode.Left + 2, labelTextItemCode);

            // Value rectangle (merge, cao bằng 2 dòng)
            Rectangle valueRectItemCode = new Rectangle(
                rect.Left + padding + titleWidth,
                (int)yRow3,
                rect.Width - titleWidth - 2 * padding,
                (int)(2 * rowHeight - spacing) // gộp 2 dòng
            );
            g.DrawLine(borderPen, valueRectItemCode.Left, valueRectItemCode.Bottom, valueRectItemCode.Right, valueRectItemCode.Bottom);

            // Nội dung: ItemCode + ProductName stacked
            string textItemCode = model.ItemCode;
            string textProductName = model.ProductName;

            // Tính toạ độ chia đôi chiều cao
            float midY = valueRectItemCode.Top + valueRectItemCode.Height / 2;

            StringFormat sfLeft = new StringFormat
            {
                Alignment = StringAlignment.Near,
                LineAlignment = StringAlignment.Center,
                Trimming = StringTrimming.EllipsisCharacter,
                FormatFlags = StringFormatFlags.NoWrap
            };

            RectangleF itemCodeTextRect = new RectangleF(
                valueRectItemCode.Left + 5,
                valueRectItemCode.Top,
                valueRectItemCode.Width - 8,
                valueRectItemCode.Height / 2
            );

            RectangleF productNameTextRect = new RectangleF(
                valueRectItemCode.Left + 5,
                midY,
                valueRectItemCode.Width - 8,
                valueRectItemCode.Height / 2
            );

            // Vẽ ItemCode, chỉ co chữ, không xuống dòng
            DrawStringAutoFitSingleLine(g, textItemCode, fontRegular, textBrush, itemCodeTextRect, sfLeft, 5.0f);
            // Vẽ ProductName, chỉ co chữ, không xuống dòng
            DrawStringAutoFitSingleLine(g, textProductName, fontRegular, textBrush, productNameTextRect, sfLeft, 5.0f);

            // ---------------- Quantity ----------------
            float yRow4 = rect.Top + padding + 4 * rowHeight + spacing * 4;

            // Label
            Rectangle labelRectQuantity = new Rectangle(rect.Left + padding, (int)yRow4, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectQuantity);
            float labelTextYQuantity = labelRectQuantity.Top + (labelRectQuantity.Height - fontRegular.Height) / 2;
            g.DrawString("Quantity", fontRegular, textBrush,
                         labelRectQuantity.Left + 2, labelTextYQuantity);

            // Value rectangle
            Rectangle valueRectRow4 = new Rectangle(rect.Left + padding + titleWidth, (int)yRow4,
                rect.Width - titleWidth - 2 * padding, (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow4.Left, valueRectRow4.Bottom, valueRectRow4.Right, valueRectRow4.Bottom);

            // Chia thành (3-2-1-3-1)
            int totalWidthRow4 = valueRectRow4.Width;
            int partWidthRow4 = totalWidthRow4 / 10;

            Rectangle rectPartSampleNo = new Rectangle(valueRectRow4.Left, valueRectRow4.Top, partWidthRow4 * 3, valueRectRow4.Height);
            Rectangle rectPartColorCode = new Rectangle(rectPartSampleNo.Right, valueRectRow4.Top, partWidthRow4 * 2, valueRectRow4.Height);
            Rectangle rectPartColorSize = new Rectangle(rectPartColorCode.Right, valueRectRow4.Top, partWidthRow4 * 1, valueRectRow4.Height);
            Rectangle rectPartQuantity = new Rectangle(rectPartColorSize.Right, valueRectRow4.Top, partWidthRow4 * 3, valueRectRow4.Height);
            Rectangle rectPartUnit = new Rectangle(rectPartQuantity.Right, valueRectRow4.Top, partWidthRow4 * 1, valueRectRow4.Height);

            // Dữ liệu
            string sampleNo = model.SampleNo;
            string colorCode = model.ColorCode;
            string sizeName = model.SizeName;
            string quantity = model.Quantity.ToString("N0");
            string unit = model.Unit;

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

            // Vẽ các phần
            DrawTextInRect(g, rectPartSampleNo, "\u00A0" + sampleNo, fontRegular, textBrush);
            DrawTextInRect(g, rectPartColorCode, colorCode, fontRegular, textBrush);
            DrawTextInRect(g, rectPartColorSize, sizeName, fontRegular, textBrush);
            DrawTextInRect(g, rectPartQuantity, quantity, fontBold, textBrush, true);
            DrawTextInRect(g, rectPartUnit, unit, fontRegular, textBrush, true);

            // --- PO Line ---
            float yRow5 = rect.Top + padding + 5 * rowHeight + spacing * 5;
            Rectangle labelRectRow5 = new Rectangle(rect.Left + padding, (int)yRow5, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow5);

            // Nội dung PO Line (canh giữa)
            string textPOLine = model.POLine;
            SizeF textSizePOLine = g.MeasureString(textPOLine, fontRegular);
            Font drawFontPOLine = fontRegular;
            if (textSizePOLine.Width > labelRectRow5.Width - 10)
            {
                float shrinkRatio = (labelRectRow5.Width - 10) / textSizePOLine.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontPOLine = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            g.DrawString(textPOLine, drawFontPOLine, textBrush, labelRectRow5, sfCenter);

            // Vùng giá trị PO Line (thu gọn để chừa QR bên phải)
            int qrSize = 42;
            int qrOffset = 12;
            Rectangle valueRectRow5 = new Rectangle(
                rect.Left + padding + titleWidth,
                (int)yRow5,
                (int)(rect.Width - titleWidth - 2 * padding - qrSize - 10),
                (int)(rowHeight - spacing)
            );
            g.DrawLine(borderPen, valueRectRow5.Left, valueRectRow5.Bottom, valueRectRow5.Right, valueRectRow5.Bottom);

            // Nội dung KeyNo
            string textLeft = "\u00A0" + model.KeyNo;
            SizeF textSizeColorName = g.MeasureString(textLeft, fontRegular);
            Font drawFontColorName = fontRegular;
            if (textSizeColorName.Width > valueRectRow5.Width - 10)
            {
                float shrinkRatio = (valueRectRow5.Width - 10) / textSizeColorName.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontColorName = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            g.DrawString(textLeft, drawFontColorName, textBrush, valueRectRow5, sfLeft);

            // QR code bên phải hàng PO Line
            Rectangle qrRect = new Rectangle(
                (int)(rect.Right - padding - qrSize),
                (int)(yRow5 + ((rowHeight - qrSize) / 2) + qrOffset),
                qrSize,
                qrSize
            );
            g.DrawRectangle(borderPen, qrRect);

            // --- Vendors Ref ---
            float yRow6 = rect.Top + padding + 6 * rowHeight + spacing * 6;
            Rectangle labelRectRow6 = new Rectangle(rect.Left + padding, (int)yRow6, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow6);
            g.DrawString("Vendors Ref.", fontRegular, textBrush, labelRectRow6.Left + 2, labelRectRow6.Top + (labelRectRow6.Height - fontRegular.Height) / 2);

            Rectangle valueRectRow6 = new Rectangle(
                rect.Left + padding + titleWidth,
                (int)yRow6,
                (int)(rect.Width - titleWidth - 2 * padding - qrSize - 10),
                (int)(rowHeight - spacing)
            );
            g.DrawLine(borderPen, valueRectRow6.Left, valueRectRow6.Bottom, valueRectRow6.Right, valueRectRow6.Bottom);

            string textVendor = "\u00A0" + model.Vendors;
            SizeF textSizeVendor = g.MeasureString(textVendor, fontRegular);
            Font drawFontVendor = fontRegular;
            if (textSizeVendor.Width > valueRectRow6.Width - 10)
            {
                float shrinkRatio = (valueRectRow6.Width - 10) / textSizeVendor.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontVendor = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            g.DrawString(textVendor, drawFontVendor, textBrush, valueRectRow6, sfLeft);

            // --- Vùng dưới ---
            float bottomTop = rect.Top + padding + 7 * rowHeight + spacing * 7;
            string noticeText = "Please open and confirm this package immediately on receipt.";
            Font fontSmall = new Font(fontRegular.FontFamily, 7.2f, FontStyle.Regular);
            g.DrawString(noticeText, fontSmall, textBrush, rect.Left + padding, bottomTop);

            string footerText = "MADE IN VIETNAM";
            float footerTop = bottomTop + fontSmall.Height;
            g.DrawString(footerText, fontSmall, textBrush, rect.Left + padding, footerTop);

            // Logo bên phải Footer
            string logoPath = $@"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}\Template\logo.png";
            if (File.Exists(logoPath))
            {
                using (System.Drawing.Image logo = System.Drawing.Image.FromFile(logoPath))
                {
                    float aspectRatio = (float)logo.Width / logo.Height;
                    int targetHeight = 12;
                    int targetWidth = (int)(targetHeight * aspectRatio);

                    Rectangle logoRect = new Rectangle(
                        (int)(rect.Right - padding - targetWidth),
                        (int)(footerTop) - 3,
                        targetWidth,
                        targetHeight
                    );

                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                    g.DrawImage(logo, logoRect);
                }
            }
        }

        private void DrawSingleLayout(Graphics g, AsicsModel model)
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
            float contentHeight = 92; // vùng nội dung
            float rowHeight = contentHeight / 5f;
            int spacing = 2;
            int titleWidth = 75;

            // ---------------- Customer ----------------
            float yRow1 = padding;
            Rectangle labelRectRow1 = new Rectangle(padding, (int)yRow1, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow1);
            g.DrawString("Customer", fontRegular, textBrush,
                         labelRectRow1.Left + 2, labelRectRow1.Top + (labelRectRow1.Height - fontRegular.Height) / 2);

            Rectangle valueRectCustomer = new Rectangle(padding + titleWidth, (int)yRow1,
                (int)(width - titleWidth - 2 * padding), (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectCustomer.Left, valueRectCustomer.Bottom, valueRectCustomer.Right, valueRectCustomer.Bottom);

            string textCustomer = model.Customer;
            SizeF sizeCustomer = g.MeasureString(textCustomer, fontRegular);
            Font fontCustomer = fontRegular;
            if (sizeCustomer.Width > valueRectCustomer.Width - 5)
            {
                float shrink = (valueRectCustomer.Width - 5) / sizeCustomer.Width;
                fontCustomer = new Font(fontRegular.FontFamily, fontRegular.Size * shrink, fontRegular.Style);
            }
            g.DrawString(textCustomer, fontCustomer, textBrush,
                         valueRectCustomer.Left + 5, valueRectCustomer.Top + (valueRectCustomer.Height - fontCustomer.Height) / 2);
            //----------------Order No(3-3–4)----------------
            float yRow2 = padding + 1 * rowHeight + spacing * 1;
            Rectangle labelRectRow2 = new Rectangle(padding, (int)yRow2, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow2);
            g.DrawString("Order No", fontRegular, textBrush,
                         labelRectRow2.Left + 2, labelRectRow2.Top + (labelRectRow2.Height - fontRegular.Height) / 2);

            int totalValueWidth = (int)(width - titleWidth - 2 * padding);
            Rectangle valueRectRow2 = new Rectangle(padding + titleWidth, (int)yRow2, totalValueWidth, (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow2.Left, valueRectRow2.Bottom, valueRectRow2.Right, valueRectRow2.Bottom);

            int partWidthRow2 = totalValueWidth / 10;
            Rectangle valueRectLeft = new Rectangle(valueRectRow2.Left, valueRectRow2.Top, partWidthRow2 * 3, valueRectRow2.Height);
            Rectangle valueRectCenter = new Rectangle(valueRectLeft.Right, valueRectRow2.Top, partWidthRow2 * 3, valueRectRow2.Height);
            Rectangle valueRectRight = new Rectangle(valueRectCenter.Right, valueRectRow2.Top, partWidthRow2 * 4, valueRectRow2.Height);

            // Order No (căn trái)
            SizeF textSizeOrderNo = g.MeasureString(model.OrderNo, fontBold);
            Font drawFontOrderNo = fontBold;
            if (textSizeOrderNo.Width > valueRectLeft.Width - 10)
            {
                float shrinkRatio = (valueRectLeft.Width - 10) / textSizeOrderNo.Width;
                float newSize = fontBold.Size * shrinkRatio;
                drawFontOrderNo = new Font(fontBold.FontFamily, newSize, fontBold.Style);
            }
            float valueTextYOrderNo = valueRectLeft.Top + (valueRectLeft.Height - drawFontOrderNo.Height) / 2;
            g.DrawString(model.OrderNo, drawFontOrderNo, textBrush, valueRectLeft.Left + 5, valueTextYOrderNo);

            // SKU (căn giữa)
            string valueTextSKU = model.SKUNo;
            SizeF textSizeSKU = g.MeasureString(valueTextSKU, fontRegular);
            Font drawFontSKU = fontRegular;
            if (textSizeSKU.Width > valueRectCenter.Width - 10)
            {
                float shrinkRatio = (valueRectCenter.Width - 10) / textSizeSKU.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontSKU = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            float valueTextYSKU = valueRectCenter.Top + (valueRectCenter.Height - drawFontSKU.Height) / 2;
            StringFormat sfCenter = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString(valueTextSKU, drawFontSKU, textBrush, valueRectCenter, sfCenter);

            // Contract No (căn phải)
            string valueTextContractNo = model.ContractNo;
            SizeF textSizeContractNo = g.MeasureString(valueTextContractNo, fontRegular);
            Font drawFontContractNo = fontRegular;
            if (textSizeContractNo.Width > valueRectRight.Width - 10)
            {
                float shrinkRatio = (valueRectRight.Width - 10) / textSizeContractNo.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontContractNo = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }

            StringFormat sf = new StringFormat
            {
                Alignment = StringAlignment.Far,        // căn sát mép phải
                LineAlignment = StringAlignment.Center  // căn giữa theo chiều dọc
            };

            g.DrawString(valueTextContractNo, drawFontContractNo, textBrush, valueRectRight, sf);

            // ---------------- Item Code (gộp dòng 2 + 3) ----------------
            float yRow3 = padding + 2 * rowHeight + spacing * 2;

            // Label (cao bằng 2 dòng)
            Rectangle labelRectItemCode = new Rectangle(
                padding,
                (int)yRow3,
                titleWidth,
                (int)(2 * rowHeight - spacing) // gộp 2 dòng
            );
            g.DrawRectangle(borderPen, labelRectItemCode);
            float labelTextItemCode = labelRectItemCode.Bottom - fontRegular.Height;
            g.DrawString("Item Code", fontRegular, textBrush, labelRectItemCode.Left + 2, labelTextItemCode);

            // Value rectangle (merge, cao bằng 2 dòng)
            Rectangle valueRectItemCode = new Rectangle(
                padding + titleWidth,
                (int)yRow3,
                (int)(width - titleWidth - 2 * padding),
                (int)(2 * rowHeight - spacing) // gộp 2 dòng
            );
            g.DrawLine(borderPen, valueRectItemCode.Left, valueRectItemCode.Bottom, valueRectItemCode.Right, valueRectItemCode.Bottom);

            // Nội dung: ItemCode + ProductName stacked
            string textItemCode = model.ItemCode;
            string textProductName = model.ProductName;

            float midY = valueRectItemCode.Top + valueRectItemCode.Height / 2;

            StringFormat sfLeft = new StringFormat
            {
                Alignment = StringAlignment.Near,
                LineAlignment = StringAlignment.Center,
                Trimming = StringTrimming.EllipsisCharacter,
                FormatFlags = StringFormatFlags.NoWrap
            };

            RectangleF itemCodeRect = new RectangleF(
                valueRectItemCode.Left + 5,
                valueRectItemCode.Top,
                valueRectItemCode.Width - 8,
                valueRectItemCode.Height / 2
            );

            RectangleF productNameRect = new RectangleF(
                valueRectItemCode.Left + 5,
                midY,
                valueRectItemCode.Width - 8,
                valueRectItemCode.Height / 2
            );

            // Vẽ ItemCode, chỉ co chữ, không xuống dòng
            DrawStringAutoFitSingleLine(g, textItemCode, fontRegular, textBrush, itemCodeRect, sfLeft, 5.0f);
            // Vẽ ProductName, chỉ co chữ, không xuống dòng
            DrawStringAutoFitSingleLine(g, textProductName, fontRegular, textBrush, productNameRect, sfLeft, 5.0f);

            //----------------Quantity----------------
            float yRow4 = padding + 4 * rowHeight + spacing * 4;
            Rectangle labelRectQuantity = new Rectangle(padding, (int)yRow4, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectQuantity);
            float labelTextYQuantity = labelRectQuantity.Top + (labelRectQuantity.Height - fontRegular.Height) / 2;
            g.DrawString("Quantity", fontRegular, textBrush, labelRectQuantity.Left + 2, labelTextYQuantity);
            Rectangle valueRectRow4 = new Rectangle(padding + titleWidth, (int)yRow4,
                (int)(width - titleWidth - 2 * padding), (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow4.Left, valueRectRow4.Bottom, valueRectRow4.Right, valueRectRow4.Bottom);

            // Chia thành (3-2-1-3-1)
            int totalWidthRow4 = valueRectRow4.Width;
            int partWidthRow4 = totalWidthRow4 / 10;

            Rectangle rectPartSampleNo = new Rectangle(valueRectRow4.Left, valueRectRow4.Top, partWidthRow4 * 3, valueRectRow4.Height);
            Rectangle rectPartColorCode = new Rectangle(rectPartSampleNo.Right, valueRectRow4.Top, partWidthRow4 * 2, valueRectRow4.Height);
            Rectangle rectPartColorSize = new Rectangle(rectPartColorCode.Right, valueRectRow4.Top, partWidthRow4 * 1, valueRectRow4.Height);
            Rectangle rectPartQuantity = new Rectangle(rectPartColorSize.Right, valueRectRow4.Top, partWidthRow4 * 3, valueRectRow4.Height);
            Rectangle rectPartUnit = new Rectangle(rectPartQuantity.Right, valueRectRow4.Top, partWidthRow4 * 1, valueRectRow4.Height);

            string quantity = model.Quantity.ToString("N0");
            string unit = model.Unit;

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
            DrawTextInRect(g, rectPartSampleNo, "\u00A0" + model.SampleNo, fontRegular, textBrush);
            DrawTextInRect(g, rectPartColorCode, model.ColorCode, fontRegular, textBrush);
            DrawTextInRect(g, rectPartColorSize, model.SizeName, fontRegular, textBrush);
            DrawTextInRect(g, rectPartQuantity, quantity, fontBold, textBrush, true);
            DrawTextInRect(g, rectPartUnit, unit, fontRegular, textBrush, true);

            // --- PO Line ---
            float yRow5 = padding + 5 * rowHeight + spacing * 5;
            Rectangle labelRectRow5 = new Rectangle(padding, (int)yRow5, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow5);

            // Nội dung PO Line (canh giữa trong ô)
            string textPOLine = model.POLine;
            SizeF textSizePOLine = g.MeasureString(textPOLine, fontRegular);
            Font drawFontPOLine = fontRegular;
            if (textSizePOLine.Width > labelRectRow5.Width - 10)
            {
                float shrinkRatio = (labelRectRow5.Width - 10) / textSizePOLine.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontPOLine = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }

            // Vẽ chữ PO Line canh giữa
            g.DrawString(textPOLine, drawFontPOLine, textBrush, labelRectRow5, sfCenter);

            // Vùng giá trị PO Line (thu gọn để chừa QR bên phải)
            int qrSize = 42;
            int qrOffset = 12; // đẩy xuống một chút
            Rectangle valueRectRow5 = new Rectangle(
                padding + titleWidth,
                (int)yRow5,
                (int)(width - titleWidth - 2 * padding - qrSize - 10), // trừ chỗ QR
                (int)(rowHeight - spacing)
            );

            // Vẽ line chỉ dưới vùng chữ, không chạm QR
            g.DrawLine(borderPen, valueRectRow5.Left, valueRectRow5.Bottom, valueRectRow5.Right, valueRectRow5.Bottom);

            // Nội dung bên trái (KeyNo)
            string textLeft = "\u00A0" + model.KeyNo;
            SizeF textSizeColorName = g.MeasureString(textLeft, fontRegular);
            Font drawFontColorName = fontRegular;
            if (textSizeColorName.Width > valueRectRow5.Width - 10)
            {
                float shrinkRatio = (valueRectRow5.Width - 10) / textSizeColorName.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontColorName = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            g.DrawString(textLeft, drawFontColorName, textBrush, valueRectRow5, sfLeft);

            // QR code nằm ngay bên phải hàng PO Line
            Rectangle qrRect = new Rectangle(
                (int)(width - padding - qrSize),
                (int)(yRow5 + ((rowHeight - qrSize) / 2) + qrOffset), // căn giữa + thụt xuống
                qrSize,
                qrSize
            );
            g.DrawRectangle(borderPen, qrRect);

            // --- Vendors Ref ---
            float yRow6 = padding + 6 * rowHeight + spacing * 6;
            Rectangle labelRectRow6 = new Rectangle(padding, (int)yRow6, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow6);
            float labelTextYRow6 = labelRectRow6.Top + (labelRectRow6.Height - fontRegular.Height) / 2;
            g.DrawString("Vendors Ref.", fontRegular, textBrush, labelRectRow6.Left + 2, labelTextYRow6);

            // Vùng giá trị Vendor (thu gọn để không chạm QR)
            Rectangle valueRectRow6 = new Rectangle(
                padding + titleWidth,
                (int)yRow6,
                (int)(width - titleWidth - 2 * padding - qrSize - 10), // trừ chỗ QR
                (int)(rowHeight - spacing)
            );

            // Vẽ line chỉ dưới vùng chữ Vendor, không chạm QR
            g.DrawLine(borderPen, valueRectRow6.Left, valueRectRow6.Bottom, valueRectRow6.Right, valueRectRow6.Bottom);

            // Nội dung Vendor (co giãn font nếu quá dài)
            string textVendor = "\u00A0" + model.Vendors;
            SizeF textSizeVendor = g.MeasureString(textVendor, fontRegular);
            Font drawFontVendor = fontRegular;
            if (textSizeVendor.Width > valueRectRow6.Width - 10)
            {
                float shrinkRatio = (valueRectRow6.Width - 10) / textSizeVendor.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontVendor = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            g.DrawString(textVendor, drawFontVendor, textBrush, valueRectRow6, sfLeft);

            // --- Vùng dưới ---
            float bottomTop = padding + 7 * rowHeight + spacing * 7;

            // Notice (chỉ chữ, không khung)
            string noticeText = "Please open and confirm this package immediately on receipt.";
            Font fontSmall = new Font(fontRegular.FontFamily, 7.2f, FontStyle.Regular);
            g.DrawString(noticeText, fontSmall, textBrush, padding, bottomTop);

            // Footer (chỉ chữ, không khung, đặt ngay dưới notice)
            string footerText = "MADE IN VIETNAM";
            float footerTop = bottomTop + fontSmall.Height;
            g.DrawString(footerText, fontSmall, textBrush, padding, footerTop);

            // Logo bên phải Footer
            string logoPath = $@"{Path.GetDirectoryName(assembly.Location)}\Template\logo.png";
            if (File.Exists(logoPath))
            {
                using (System.Drawing.Image logo = System.Drawing.Image.FromFile(logoPath))
                {
                    // Tính tỉ lệ gốc
                    float aspectRatio = (float)logo.Width / logo.Height;

                    // Chiều cao mong muốn (ví dụ 12)
                    int targetHeight = 12;
                    int targetWidth = (int)(targetHeight * aspectRatio);

                    Rectangle logoRect = new Rectangle(
                        (int)(width - padding - targetWidth),
                        (int)(footerTop) - 3,
                        targetWidth,
                        targetHeight
                    );

                    // Vẽ với chất lượng cao
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                    g.DrawImage(logo, logoRect);
                }
            }
        }

        private void FrmPackingSealAsics_Load(object sender, EventArgs e)
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
        }

        private void BuildPrintQueue()
        {
            printQueue = new List<AsicsModel>();
            foreach (DataGridViewRow row in dgvData.Rows)
            {
                if (row.IsNewRow) continue;
                string SkuNo = row.Cells["SKUNo"].Value.ToString() ?? string.Empty;
                if (!SkuNo.StartsWith("-") && !string.IsNullOrEmpty(SkuNo))
                {
                    SkuNo = "-" + SkuNo;
                }
                string sizeName = row.Cells["SizeName"].Value?.ToString() ?? string.Empty;
                if (!sizeName.StartsWith(".") && !string.IsNullOrEmpty(sizeName))
                {
                    sizeName = "." + sizeName;
                }
                string colorCode = row.Cells["ColorCode"].Value?.ToString() ?? string.Empty;
                if (!colorCode.StartsWith(".") && !string.IsNullOrEmpty(colorCode))
                {
                    colorCode = "." + colorCode;
                }
                var newModel = new AsicsModel()
                {
                    ImagePath = imagePath,
                    Customer = row.Cells["Customer"].Value?.ToString(),
                    OrderNo = row.Cells["OrderNo"].Value?.ToString(),
                    SKUNo = SkuNo,
                    ContractNo = row.Cells["ContractNo"].Value?.ToString(),
                    ItemCode = row.Cells["ItemCode"].Value?.ToString(),
                    ProductName = row.Cells["ProductName"].Value?.ToString(),
                    Unit = row.Cells["Unit"].Value?.ToString(),
                    SampleNo = row.Cells["SampleNo"].Value?.ToString(),
                    ColorCode = colorCode,
                    SizeName = sizeName,
                    Quantity = int.TryParse(row.Cells["Quantity"].Value?.ToString(), out int qty) ? qty : 0,
                    POLine = row.Cells["POLine"].Value?.ToString(),
                    KeyNo = row.Cells["KeyNo"].Value?.ToString(),
                    Vendors = row.Cells["Vendors"].Value?.ToString(),
                    Index = Convert.ToInt32(row.Cells["PCS"].Value)
                };
                for (int i = 0; i < newModel.Index; i++)
                {
                    printQueue.Add(new AsicsModel
                    {
                        ImagePath = newModel.ImagePath,
                        Customer = newModel.Customer,
                        OrderNo = newModel.OrderNo,
                        SKUNo = newModel.SKUNo,
                        ContractNo = newModel.ContractNo,
                        ItemCode = newModel.ItemCode,
                        ProductName = newModel.ProductName,
                        Unit = newModel.Unit,
                        SampleNo = newModel.SampleNo,
                        ColorCode = newModel.ColorCode,
                        SizeName = newModel.SizeName,
                        Quantity = newModel.Quantity,
                        POLine = newModel.POLine,
                        KeyNo = newModel.KeyNo,
                        Vendors = newModel.Vendors
                    });
                }
            }
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

        private void FrmPackingSealAsics_KeyDown(object sender, KeyEventArgs e)
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

        private void DrawStringAutoFitSingleLine(Graphics g, string text, Font baseFont, Brush brush, RectangleF rect,
            StringFormat format, float minFontSize)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            // Khong cho text co ky tu xuong dong that
            text = text.Replace("\r", " ").Replace("\n", " ");

            float fontSize = baseFont.Size;
            Font drawFont = null;

            try
            {
                while (fontSize >= minFontSize)
                {
                    drawFont = new Font(baseFont.FontFamily, fontSize, baseFont.Style);

                    // Do text nhu 1 dong dai, khong cho wrap
                    SizeF textSize = g.MeasureString(text, drawFont, new SizeF(10000, rect.Height), format);

                    if (textSize.Width <= rect.Width && textSize.Height <= rect.Height)
                    {
                        break;
                    }

                    drawFont.Dispose();
                    drawFont = null;

                    fontSize -= 0.5f;
                }

                if (drawFont == null)
                {
                    drawFont = new Font(baseFont.FontFamily, minFontSize, baseFont.Style);
                }

                g.DrawString(text, drawFont, brush, rect, format);
            }
            finally
            {
                if (drawFont != null)
                {
                    drawFont.Dispose();
                }
            }
        }

    }
}
