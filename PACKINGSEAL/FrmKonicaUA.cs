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
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace PACKINGSEAL
{
    public partial class FrmKonicaUA : Form
    {
        public string imagePath = string.Empty;
        private string pdfOutputPath = string.Empty;
        private Assembly assembly = Assembly.GetExecutingAssembly();
        private PrintDocument printDocument = new PrintDocument();
        private KonicaUnderamour currentModel;
        private int printedIndex = 0;
        private float zoomFactor = 1.0f;
        private List<KonicaUnderamour> printQueue;
        private Point crosshairPoint = Point.Empty;
        private Bitmap layoutCache;

        public FrmKonicaUA(FilePathModel filePathModel)
        {
            InitializeComponent();
            string UserName = Environment.MachineName;
            lb_info_1.Text = "HC | KONICA Underamour Packing Onegai Seal- " + UserName;
            string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
            var version = "";
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                version = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            this.Text = "KONICA UNDERAMOUR PACKING ONEGAI SEAL (梱包シール) Ver: " + version;
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
                ExcelUtil excelUtil = new ExcelUtil();
                DataTable dt = new DataTable();
                dt = new DataTable();
                dt.Columns.Add("PONo");
                dt.Columns.Add("Customer");
                dt.Columns.Add("Product");
                dt.Columns.Add("Description");
                dt.Columns.Add("ColorName");
                dt.Columns.Add("SizeName");
                dt.Columns.Add("Quantity");
                dt.Columns.Add("OrderNo");
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
                AppCommonModule appCommonModule = new AppCommonModule();
                string customer = appCommonModule.GetCustomer(poNo);
                if (customer == null)
                {
                    MessageBox.Show($"This PONo customer name not found: {poNo}", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                for (int i = 1; i <= sheetAssortmentList.LastRowNum; i++)
                {
                    int quantity = (int)Math.Round(double.Parse(excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 6)));
                    string comment = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 7);
                    if (string.IsNullOrEmpty(comment))
                    {
                        comment = customer;
                    }
                    DataRow dataRow = dt.NewRow();
                    dataRow["PONo"] = poNo;
                    dataRow["Customer"] = comment;
                    dataRow["Product"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 0);
                    dataRow["Description"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 1);
                    dataRow["ColorName"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 4);
                    dataRow["SizeName"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 5);
                    dataRow["Quantity"] = quantity.ToString("N0");
                    dataRow["OrderNo"] = excelUtil.GetCellValueWithMerge(sheetAssortmentList, i, 9);
                    dataRow["PCS"] = "1";
                    dt.Rows.Add(dataRow.ItemArray);
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
            string htmlContent = string.Empty;
            HtmlResponse templateHTML = new HtmlResponse();
            currentModel = new KonicaUnderamour()
            {
                ImagePath = imagePath,
                PONo = row.Cells["PONo"].Value?.ToString() ?? string.Empty,
                Customer = row.Cells["Customer"].Value?.ToString() ?? string.Empty,
                Product = row.Cells["Product"].Value?.ToString() ?? string.Empty,
                Description = row.Cells["Description"].Value?.ToString() ?? string.Empty,
                ColorName = row.Cells["ColorName"].Value?.ToString() ?? string.Empty,
                SizeName = row.Cells["SizeName"].Value?.ToString() ?? string.Empty,
                OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty,
                Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString())
            };
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

        private void DrawPrintLayout(Graphics g, KonicaUnderamour model, Rectangle rect)
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
            int titleWidth = 65;
            int spacing = 2;

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

            //----------------Order No(3–3-4)----------------
            float yRow2 = rect.Top + padding + 1 * rowHeight + spacing * 1;

            // Label "Order No"
            Rectangle labelRectRow2 = new Rectangle(rect.Left + padding, (int)yRow2, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow2);
            g.DrawString("Order No", fontRegular, textBrush,
                         labelRectRow2.Left + 2,
                         labelRectRow2.Top + (labelRectRow2.Height - fontRegular.Height) / 2);

            // Vùng giá trị
            int totalValueWidth = rect.Width - titleWidth - 2 * padding;
            Rectangle valueRectRow2 = new Rectangle(rect.Left + padding + titleWidth, (int)yRow2,
                                                    totalValueWidth, (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow2.Left, valueRectRow2.Bottom, valueRectRow2.Right, valueRectRow2.Bottom);

            // Chia 3–3-4 theo tỉ lệ 10 phần để dễ tính toán
            int partWidthRow2 = totalValueWidth / 10;
            Rectangle valueRectLeft = new Rectangle(valueRectRow2.Left, valueRectRow2.Top, partWidthRow2 * 6, valueRectRow2.Height);
            Rectangle valueRectRight = new Rectangle(valueRectLeft.Right, valueRectRow2.Top, partWidthRow2 * 4, valueRectRow2.Height);

            // PoNo (căn trái)
            SizeF textSizeOrderNo = g.MeasureString(model.OrderNo, fontRegular);
            Font drawFontOrderNo = fontRegular;
            if (textSizeOrderNo.Width > valueRectLeft.Width - 10)
            {
                float shrinkRatio = (valueRectLeft.Width - 10) / textSizeOrderNo.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontOrderNo = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            float valueTextYOrderNo = valueRectLeft.Top + (valueRectLeft.Height - drawFontOrderNo.Height) / 2;
            g.DrawString(model.OrderNo, drawFontOrderNo, textBrush, valueRectLeft.Left + 5, valueTextYOrderNo);

            // PONo (căn phải)
            string valueTextPOCustomer = model.PONo;
            SizeF textSizePOCustomer = g.MeasureString(valueTextPOCustomer, fontBold);
            Font drawFontPOCustomer = fontBold;
            if (textSizePOCustomer.Width > valueRectRight.Width - 10)
            {
                float shrinkRatio = (valueRectRight.Width - 10) / textSizePOCustomer.Width;
                float newSize = fontBold.Size * shrinkRatio;
                drawFontPOCustomer = new Font(fontBold.FontFamily, newSize, fontBold.Style);
            }

            StringFormat sf = new StringFormat
            {
                Alignment = StringAlignment.Far,        // căn sát mép phải
                LineAlignment = StringAlignment.Center  // căn giữa theo chiều dọc
            };

            g.DrawString(valueTextPOCustomer, drawFontPOCustomer, textBrush, valueRectRight, sf);

            // ---------------- Item Code ----------------
            float yRow3 = rect.Top + padding + 2 * rowHeight + spacing * 2;

            // Label "Item Code"
            Rectangle labelRectItemCode = new Rectangle(rect.Left + padding, (int)yRow3, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectItemCode);
            float labelTextItemCode = labelRectItemCode.Top + (labelRectItemCode.Height - fontRegular.Height) / 2;
            g.DrawString("Item Code", fontRegular, textBrush, labelRectItemCode.Left + 2, labelTextItemCode);

            // Value rectangle (chia làm 2 phần)
            Rectangle valueRectItemCode = new Rectangle(rect.Left + padding + titleWidth, (int)yRow3,
                (int)(rect.Width - titleWidth - 2 * padding), (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectItemCode.Left, valueRectItemCode.Bottom, valueRectItemCode.Right, valueRectItemCode.Bottom);

            int halfWidth = valueRectItemCode.Width / 2;
            Rectangle rectProduct = new Rectangle(valueRectItemCode.Left, valueRectItemCode.Top, halfWidth, valueRectItemCode.Height);
            Rectangle rectCustom = new Rectangle(rectProduct.Right, valueRectItemCode.Top, halfWidth, valueRectItemCode.Height);

            // --- Product (căn trái, co font nếu quá dài) ---
            string valueTextProduct = model.Product;
            SizeF textSizeProduct = g.MeasureString(valueTextProduct, fontRegular);
            Font drawFontProduct = fontRegular;
            if (textSizeProduct.Width > rectProduct.Width - 10)
            {
                float shrinkRatio = (rectProduct.Width - 10) / textSizeProduct.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontProduct = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            float valueTextYProduct = rectProduct.Top + (rectProduct.Height - drawFontProduct.Height) / 2;
            g.DrawString(valueTextProduct, drawFontProduct, textBrush, rectProduct.Left + 5, valueTextYProduct);

            // --- Custom text (bạn tự nhập, căn phải) ---
            string valueTextDescription = model.Description; // thay bằng giá trị bạn muốn
            Font customFontDescription = new Font(fontRegular.FontFamily, 10, FontStyle.Regular);
            StringFormat sfRight = new StringFormat
            {
                Alignment = StringAlignment.Far,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString(valueTextDescription, customFontDescription, textBrush, rectCustom, sfRight);

            //---------------- Quantity ----------------
            float yRow4 = rect.Top + padding + 3 * rowHeight + spacing * 3;

            // Label "Quantity"
            Rectangle labelRectQuantity = new Rectangle(rect.Left + padding, (int)yRow4, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectQuantity);
            float labelTextYQuantity = labelRectQuantity.Top + (labelRectQuantity.Height - fontRegular.Height) / 2;
            g.DrawString("Quantity", fontRegular, textBrush, labelRectQuantity.Left + 2, labelTextYQuantity);

            // Vùng giá trị
            Rectangle valueRectRow4 = new Rectangle(rect.Left + padding + titleWidth, (int)yRow4,
                rect.Width - titleWidth - 2 * padding, (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow4.Left, valueRectRow4.Bottom, valueRectRow4.Right, valueRectRow4.Bottom);

            // Chia thành (9-1)
            int totalWidthRow4 = valueRectRow4.Width;
            int partWidthRow4 = totalWidthRow4 / 10;
            Rectangle rectPartQuantity = new Rectangle(valueRectRow4.Left, valueRectRow4.Top, partWidthRow4 * 9, valueRectRow4.Height);
            Rectangle rectPartUnit = new Rectangle(rectPartQuantity.Right, valueRectRow4.Top, partWidthRow4 * 1, valueRectRow4.Height);

            string quantity = model.Quantity.ToString("N0");
            string unit = "PCS";

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
            DrawTextInRect(g, rectPartQuantity, quantity, fontBold, textBrush, true);
            DrawTextInRect(g, rectPartUnit, unit, fontRegular, textBrush, true);

            //---------------- dòng riêng biệt ở dưới cùng ----------------
            float yRow5 = rect.Top + padding + 4 * rowHeight + spacing * 4;

            // Label (Color/Size)
            Rectangle labelRectRow5 = new Rectangle(rect.Left + padding, (int)yRow5, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow5);
            float labelTextYRow5 = labelRectRow5.Top + (labelRectRow5.Height - fontRegular.Height) / 2;
            g.DrawString("Color/Size", fontRegular, textBrush, labelRectRow5.Left + 2, labelTextYRow5);

            // Vùng giá trị
            Rectangle valueRectRow5 = new Rectangle(rect.Left + padding + titleWidth, (int)yRow5,
                rect.Width - titleWidth - 2 * padding, (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow5.Left, valueRectRow5.Bottom, valueRectRow5.Right, valueRectRow5.Bottom);

            int totalWidthRow5 = valueRectRow5.Width;
            int partWidthRow5 = totalWidthRow5 / 10;

            Rectangle rectLeft = new Rectangle(valueRectRow5.Left, valueRectRow5.Top, partWidthRow5 * 5, valueRectRow5.Height);
            Rectangle rectRight = new Rectangle(rectLeft.Right, valueRectRow5.Top, partWidthRow5 * 5, valueRectRow5.Height);

            // Nội dung bên trái (ColorName)
            string textColorName = "\u00A0" + model.ColorName;
            SizeF textSizeColorName = g.MeasureString(textColorName, fontBold);
            Font drawFontColorName = fontBold;
            if (textSizeColorName.Width > rectLeft.Width - 10)
            {
                float shrinkRatio = (rectLeft.Width - 10) / textSizeColorName.Width;
                float newSize = fontBold.Size * shrinkRatio;
                drawFontColorName = new Font(fontBold.FontFamily, newSize, fontBold.Style);
            }
            StringFormat sfLeft = new StringFormat
            {
                Alignment = StringAlignment.Near,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString(textColorName, drawFontColorName, textBrush, rectLeft, sfLeft);

            // Nội dung bên phải
            string textSizeName = model.SizeName;
            SizeF textSizeSizeName = g.MeasureString(textSizeName, fontRegular);
            Font drawFontRight = fontRegular;
            if (textSizeSizeName.Width > rectRight.Width - 10)
            {
                float shrinkRatio = (rectRight.Width - 10) / textSizeSizeName.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontRight = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            g.DrawString(textSizeName, drawFontRight, textBrush, rectRight, sfRight);
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
            float imageHeight = rect.Bottom - imageTop - 2;
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
            Rectangle qrRect = new Rectangle(rect.Right - padding - qrSize, (int)imageTop, qrSize, qrSize);
            g.DrawRectangle(borderPen, qrRect);
        }

        private void DrawSingleLayout(Graphics g, KonicaUnderamour model)
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
            int titleWidth = 65;

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
            //----------------Order No(6–4)----------------
            float yRow2 = padding + 1 * rowHeight + spacing * 1;
            Rectangle labelRectRow2 = new Rectangle(padding, (int)yRow2, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow2);
            g.DrawString("Order No", fontRegular, textBrush,
                         labelRectRow2.Left + 2, labelRectRow2.Top + (labelRectRow2.Height - fontRegular.Height) / 2);

            int totalValueWidth = (int)(width - titleWidth - 2 * padding);
            Rectangle valueRectRow2 = new Rectangle(padding + titleWidth, (int)yRow2, totalValueWidth, (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow2.Left, valueRectRow2.Bottom, valueRectRow2.Right, valueRectRow2.Bottom);

            int partWidthRow2 = totalValueWidth / 10;
            Rectangle valueRectLeft = new Rectangle(valueRectRow2.Left, valueRectRow2.Top, partWidthRow2 * 6, valueRectRow2.Height);
            Rectangle valueRectRight = new Rectangle(valueRectLeft.Right, valueRectRow2.Top, partWidthRow2 * 4, valueRectRow2.Height);

            // Order No (căn trái)
            SizeF textSizeOrderNo = g.MeasureString(model.OrderNo, fontRegular);
            Font drawFontOrderNo = fontRegular;
            if (textSizeOrderNo.Width > valueRectLeft.Width - 10)
            {
                float shrinkRatio = (valueRectLeft.Width - 10) / textSizeOrderNo.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontOrderNo = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            float valueTextYOrderNo = valueRectLeft.Top + (valueRectLeft.Height - drawFontOrderNo.Height) / 2;
            g.DrawString(model.OrderNo, drawFontOrderNo, textBrush, valueRectLeft.Left + 5, valueTextYOrderNo);

            // PONo (căn phải)
            string valueTextPONo = model.PONo;
            SizeF textSizePONo = g.MeasureString(valueTextPONo, fontBold);
            Font drawFontPONo = fontBold;
            if (textSizePONo.Width > valueRectRight.Width - 10)
            {
                float shrinkRatio = (valueRectRight.Width - 10) / textSizePONo.Width;
                float newSize = fontBold.Size * shrinkRatio;
                drawFontPONo = new Font(fontBold.FontFamily, newSize, fontBold.Style);
            }

            StringFormat sf = new StringFormat
            {
                Alignment = StringAlignment.Far,        // căn sát mép phải
                LineAlignment = StringAlignment.Center  // căn giữa theo chiều dọc
            };

            g.DrawString(valueTextPONo, drawFontPONo, textBrush, valueRectRight, sf);

            // ---------------- Item Code ----------------
            float yRow3 = padding + 2 * rowHeight + spacing * 2;

            // Label
            Rectangle labelRectItemCode = new Rectangle(padding, (int)yRow3, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectItemCode);
            float labelTextItemCode = labelRectItemCode.Top + (labelRectItemCode.Height - fontRegular.Height) / 2;
            g.DrawString("Item Code", fontRegular, textBrush, labelRectItemCode.Left + 2, labelTextItemCode);

            // Value rectangle (chia làm 2 phần)
            Rectangle valueRectItemCode = new Rectangle(padding + titleWidth, (int)yRow3,
                (int)(width - titleWidth - 2 * padding), (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectItemCode.Left, valueRectItemCode.Bottom, valueRectItemCode.Right, valueRectItemCode.Bottom);

            int halfWidth = valueRectItemCode.Width / 2;
            Rectangle rectProduct = new Rectangle(valueRectItemCode.Left, valueRectItemCode.Top, halfWidth, valueRectItemCode.Height);
            Rectangle rectCustom = new Rectangle(rectProduct.Right, valueRectItemCode.Top, halfWidth, valueRectItemCode.Height);

            // --- Product (căn trái, co font nếu quá dài) ---
            string valueTextProduct = model.Product;
            SizeF textSizeProduct = g.MeasureString(valueTextProduct, fontRegular);
            Font drawFontProduct = fontRegular;
            if (textSizeProduct.Width > rectProduct.Width - 10)
            {
                float shrinkRatio = (rectProduct.Width - 10) / textSizeProduct.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontProduct = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            float valueTextYProduct = rectProduct.Top + (rectProduct.Height - drawFontProduct.Height) / 2;
            g.DrawString(valueTextProduct, drawFontProduct, textBrush, rectProduct.Left + 5, valueTextYProduct);

            string valueTextDescription = model.Description;
            Font customFontDescription = new Font(fontRegular.FontFamily, 10, FontStyle.Regular);
            StringFormat sfRight = new StringFormat
            {
                Alignment = StringAlignment.Far,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString(valueTextDescription, customFontDescription, textBrush, rectCustom, sfRight);

            //----------------Quantity----------------
            float yRow4 = padding + 3 * rowHeight + spacing * 3;
            Rectangle labelRectQuantity = new Rectangle(padding, (int)yRow4, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectQuantity);
            float labelTextYQuantity = labelRectQuantity.Top + (labelRectQuantity.Height - fontRegular.Height) / 2;
            g.DrawString("Quantity", fontRegular, textBrush, labelRectQuantity.Left + 2, labelTextYQuantity);
            Rectangle valueRectRow4 = new Rectangle(padding + titleWidth, (int)yRow4,
                (int)(width - titleWidth - 2 * padding), (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow4.Left, valueRectRow4.Bottom, valueRectRow4.Right, valueRectRow4.Bottom);

            // Chia thành (8/2)
            int totalWidthRow4 = valueRectRow4.Width;
            int partWidthRow4 = totalWidthRow4 / 10;

            Rectangle rectPartQuantity = new Rectangle(valueRectRow4.Left, valueRectRow4.Top, partWidthRow4 * 9, valueRectRow4.Height);
            Rectangle rectPartUnit = new Rectangle(rectPartQuantity.Right, valueRectRow4.Top, partWidthRow4 * 1, valueRectRow4.Height);

            string quantity = model.Quantity.ToString("N0");
            string unit = "PCS";

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

            DrawTextInRect(g, rectPartQuantity, quantity, fontBold, textBrush, true);
            DrawTextInRect(g, rectPartUnit, unit, fontRegular, textBrush, true);

            //---------------- (dòng riêng biệt ở dưới cùng)----------------
            float yRow5 = padding + 4 * rowHeight + spacing * 4;
            Rectangle labelRectRow5 = new Rectangle(padding, (int)yRow5, titleWidth, (int)(rowHeight - spacing));
            g.DrawRectangle(borderPen, labelRectRow5);
            float labelTextYRow5 = labelRectRow5.Top + (labelRectRow5.Height - fontRegular.Height) / 2;
            g.DrawString("Color/Size", fontRegular, textBrush, labelRectRow5.Left + 2, labelTextYRow5);

            Rectangle valueRectRow5 = new Rectangle(padding + titleWidth, (int)yRow5,
                (int)(width - titleWidth - 2 * padding), (int)(rowHeight - spacing));
            g.DrawLine(borderPen, valueRectRow5.Left, valueRectRow5.Bottom, valueRectRow5.Right, valueRectRow5.Bottom);
            int totalWidthRow5 = valueRectRow5.Width;
            int partWidthRow5 = totalWidthRow5 / 10;

            Rectangle rectLeft = new Rectangle(valueRectRow5.Left, valueRectRow5.Top, partWidthRow5 * 5, valueRectRow5.Height);
            Rectangle rectRight = new Rectangle(rectLeft.Right, valueRectRow5.Top, partWidthRow5 * 5, valueRectRow5.Height);
            // Nội dung bên trái (ColorName)
            string textLeft = "\u00A0" + model.ColorName;
            SizeF textSizeColorName = g.MeasureString(textLeft, fontBold);
            Font drawFontColorName = fontBold;
            if (textSizeColorName.Width > rectLeft.Width - 10)
            {
                float shrinkRatio = (rectLeft.Width - 10) / textSizeColorName.Width;
                float newSize = fontBold.Size * shrinkRatio;
                drawFontColorName = new Font(fontBold.FontFamily, newSize, fontBold.Style);
            }
            StringFormat sfLeft = new StringFormat
            {
                Alignment = StringAlignment.Near,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString(textLeft, drawFontColorName, textBrush, rectLeft, sfLeft);

            // Nội dung bên phải
            string textSizeName = model.SizeName;
            SizeF textSizeSizeName = g.MeasureString(textSizeName, fontRegular);
            Font drawFontSizeName = fontRegular;
            if (textSizeSizeName.Width > rectRight.Width - 10)
            {
                float shrinkRatio = (rectRight.Width - 10) / textSizeSizeName.Width;
                float newSize = fontRegular.Size * shrinkRatio;
                drawFontSizeName = new Font(fontRegular.FontFamily, newSize, fontRegular.Style);
            }
            g.DrawString(textSizeName, drawFontSizeName, textBrush, rectRight, sfRight);

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

        private void BuildPrintQueue()
        {
            printQueue = new List<KonicaUnderamour>();
            foreach (DataGridViewRow row in dgvData.Rows)
            {
                if (row.IsNewRow) continue;
                var newModel = new KonicaUnderamour
                {
                    ImagePath = imagePath,
                    PONo = row.Cells["PONo"].Value?.ToString() ?? string.Empty,
                    Customer = row.Cells["Customer"].Value?.ToString() ?? string.Empty,
                    Product = row.Cells["Product"].Value?.ToString() ?? string.Empty,
                    Description = row.Cells["Description"].Value?.ToString() ?? string.Empty,
                    ColorName = row.Cells["ColorName"].Value?.ToString() ?? string.Empty,
                    SizeName = row.Cells["SizeName"].Value?.ToString() ?? string.Empty,
                    OrderNo = row.Cells["OrderNo"].Value?.ToString() ?? string.Empty,
                    Quantity = string.IsNullOrEmpty(row.Cells["Quantity"].Value.ToString()) ? 0 : double.Parse(row.Cells["Quantity"].Value.ToString()),
                    Index = Convert.ToInt32(row.Cells["PCS"].Value)
                };

                for (int i = 0; i < newModel.Index; i++)
                {
                    printQueue.Add(new KonicaUnderamour
                    {
                        ImagePath = newModel.ImagePath,
                        PONo = newModel.PONo,
                        Customer = newModel.Customer,
                        Product = newModel.Product,
                        Description = newModel.Description,
                        ColorName = newModel.ColorName,
                        SizeName = newModel.SizeName,
                        OrderNo = newModel.OrderNo,
                        Quantity = newModel.Quantity,
                        Index = newModel.Index
                    });
                }
            }
        }

        private void FrmPackingSealKonicaUA_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.frmKonicaUASize.Width != 0 && Properties.Settings.Default.frmKonicaUASize.Height != 0)
            {
                this.Size = Properties.Settings.Default.frmKonicaUASize;
            }
            if (Properties.Settings.Default.frmKonicaUALocation.X != 0 && Properties.Settings.Default.frmKonicaUALocation.Y != 0)
            {
                this.StartPosition = FormStartPosition.CenterScreen;
                this.Location = Properties.Settings.Default.frmKonicaUALocation;
            }
            this.WindowState = Properties.Settings.Default.frmKonicaUAState;
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

        private void FrmPackingSealKonicaUA_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.frmKonicaUAState = this.WindowState;
            if (this.WindowState == FormWindowState.Normal)
            {
                Properties.Settings.Default.frmKonicaUALocation = this.Location;
                Properties.Settings.Default.frmKonicaUASize = this.Size;
            }
            else
            {
                Properties.Settings.Default.frmKonicaUALocation = this.RestoreBounds.Location;
                Properties.Settings.Default.frmKonicaUASize = this.RestoreBounds.Size;
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

        private void FrmPackingSealUniqlo_KeyDown(object sender, KeyEventArgs e)
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

        private static string GetFirstPart(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            string[] parts = input.Split('_');
            return parts.Length > 0 ? parts[0] : string.Empty;
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
