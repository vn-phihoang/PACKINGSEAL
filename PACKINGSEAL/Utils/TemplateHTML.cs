using iText.IO.Font;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Layout;
using iText.Layout.Renderer;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Path = System.IO.Path;
using PdfFont = iText.Kernel.Font.PdfFont;


namespace PACKINGSEAL
{
    public class TemplateHTML
    {
        private Assembly assembly = Assembly.GetEntryAssembly();
        public string UQToPDF(string imagePath, string orderNo, string deliveryDestination, string poNo,
        string sampleNo, string materialCode, string sizeName, string colorName, string colorCode, double quantity,
        string keyNo, string contractNo, double quantityInBox, int index = 0)
        {
            string fontPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\MS-UIGothic.ttf";
            string fontBold = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\NotoSansJP-ExtraBold.ttf";
            float fontSizeCustomer = ResizeTextToFit(deliveryDestination, fontPath, 240);
            float fontSizeOrder = ResizeTextToFit(poNo, fontPath, 145);
            float fontSizeItem = ResizeTextToFit($"{sampleNo}/{materialCode}", fontPath, 230);
            string quantityDisplay = quantity.ToString("N0");
            if (materialCode.StartsWith("HT") || materialCode.StartsWith("WT"))
            {
                quantityDisplay = quantityInBox.ToString("N0") + "/" + quantity.ToString("N0");
            }
            float fontSizequantityDisplay = ResizeTextToFit($"{quantityDisplay}", fontBold, 90);
            string boxSide = index % 2 == 0 ? "box-left" : "box-right";
            string pageOpenTag = index % 2 == 0 ? "<page size=\"A4\">\r\n" : "";
            string pageCloseTag = index % 2 == 1 ? "</page>\r\n" : "";
            string tableContent = $@"{pageOpenTag}  
            <div class=""outer-box {boxSide}"">  
                <table class=""form-table""> 
                    <tr>  
                        <td class=""label"">Customer</td>  
                        <td colspan=""5"" style=""text-align:left; font-size:{fontSizeCustomer}px;"">
                            {deliveryDestination}
                        </td>  
                    </tr>  
                    <tr style=""height: 4px;""><td></td></tr>  
                    <tr>  
                        <td class=""label"">Order No</td>  
                        <td colspan=""3"" style=""font-size:{fontSizeOrder}px;"">{poNo}</td>  
                        <td colspan=""2"" class=""content-bold"" style=""text-align:right""> 
                            {(!string.IsNullOrEmpty(deliveryDestination) && deliveryDestination.StartsWith("REGENT GARMENT", StringComparison.OrdinalIgnoreCase) ? contractNo : "")} 
                        </td> 
                    </tr>  
                    <tr style=""height: 4px;""><td></td></tr>  
                    <tr>  
                        <td class=""label"">Item Code</td>  
                        <td colspan=""5"" style=""padding-left:8px; width:235px; font-size:{fontSizeItem}px;"">
                            {sampleNo}/{materialCode}
                        </td>  
                    </tr>  
                    <tr style=""height: 4px;""><td></td></tr>  
                    <tr>  
                        <td class=""label"">Quantity</td>  
                        <td style=""padding-left:8px;"">Col. {colorCode}</td>  
                        <td>Size {sizeName}</td>  
                        <td colspan=""3"" style=""text-align:right; width:110px"">
                            <label class=""content-bold"" style=""font-size:{fontSizequantityDisplay}px;"">{quantityDisplay}</label> 
                            <label>pcs</label> 
                        </td>
                    </tr>  
                    <tr style=""height: 4px;""><td></td></tr>  
                    <tr>  
                        <td class=""label""></td>  
                        <td colspan=""3"" class=""content-bold"" style=""padding-left:8px;"">{orderNo}</td>  
                        <td colspan=""2"" style=""text-align:left"">{keyNo}</td>  
                    </tr>  
                </table>  
                <div class=""notice"">  
                    Please open and confirm this package immediately on receipt.  
                </div>  
                <div class=""footer"">  
                    <span>MADE IN VIETNAM</span>  
                </div>  
                <div class=""container"">  
                    <div style=""width: 65%; display: inline-block;"">  
                        <img src=""{imagePath}"" class=""address-img"" />  
                    </div>  
                    <div class=""parent-container"">  
                        <div class=""qr-box""></div>  
                    </div>  
                </div>  
            </div>
            {pageCloseTag}";
            return tableContent;
        }

        public string MonbellToPDF(string imagePath, string orderNo, string poNo, string styleNo, string itemCode, string materialCode, double quantity, string color, string size, double sizeQuantity, int index = 0)
        {
            string fontPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\MS-UIGothic.ttf";
            string fontBoldPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\NotoSansJP-ExtraBold.ttf";
            float fontSizeOrderNo = ResizeTextToFit(orderNo, fontPath, 190);
            float fontQTY = ResizeTextToFit(quantity.ToString("N0"), fontPath, 55);
            float fontSizeStyleNo = ResizeTextToFit(styleNo, fontPath, 60);
            float fontSizeColor = ResizeTextToFit(color, fontPath, 60);
            float fontSizeQTY = ResizeTextToFit(sizeQuantity.ToString("N0"), fontBoldPath, 55);
            float fontItemcodeMametan = ResizeTextToFit(itemCode +" / " + materialCode, fontPath, 190);
            string boxSide = index % 2 == 0 ? "box-left" : "box-right";
            string pageOpenTag = index % 2 == 0 ? "<page size=\"A4\">\r\n" : "";
            string pageCloseTag = index % 2 == 1 ? "</page>\r\n" : "";
            string tableContent = $@"{pageOpenTag}  
            <div class=""outer-box {boxSide}"">
                <table class=""form-table"">
                    <tr>
                        <td colspan=""1"" class=""label"">Order No </td>
                        <td colspan=""4"" style=""padding-left:1.3em; font-size:{fontSizeOrderNo}px;"">
                             {orderNo}
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">TENTAC(PO NO)</td>
                        <td colspan=""1"" class=""content-bold"" style=""width:60px;"">{poNo}</td>
                        <td colspan=""1"" class=""title"">TTL Style</td>
                        <td colspan=""1"" style=""text-align:right; font-size:{fontQTY}px;"">
                           <label style=""text-align:right"">{quantity.ToString("N0")}</label>
                        </td>
                        <td colspan=""1"" style=""width:20px; text-align:right"">PCS</td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">ITEM Code/Material</td>
                        <td colspan=""4"" style=""font-size:{fontItemcodeMametan}px;"">
                            <label>{itemCode} / {materialCode}</label>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">Style No</td>
                        <td colspan=""1"" style=""font-size:{fontSizeStyleNo}px;"">
                            {styleNo}
                        </td>
                        <td colspan=""1"" class=""title"">Size</td>
                        <td colspan=""2"" class=""content-bold"" style=""text-align:right"">{size}</td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">Color</td>
                        <td colspan=""1"" style=""font-size:{fontSizeColor}px;"">
                             {color}
                        </td>
                        <td colspan=""1"" class=""title"">QTY</td>
                        <td colspan=""1"" style=""text-align:right;"">
                            <label class=""content-bold"" style=""font-size:{fontSizeQTY}px;"">{sizeQuantity.ToString("N0")}</label>
                        </td>
                        <td colspan=""1"" style=""width:20px;"">PCS</td>
                    </tr>
                </table>
                <div class=""notice"">
                    Please open and confirm this package immediately on receipt.
                </div>
                <div class=""footer"">
                    <span>MADE IN VIETNAM</span>
                </div>
                <div class=""container"">
                    <div style=""width: 65%; display: inline-block;"">
                        <img src=""{imagePath}"" class=""address-img"" />
                    </div>
                    <div class=""parent-container"">
                        <div class=""qr-box""></div>
                    </div>
                </div>
            </div>
            {pageCloseTag}";
            return tableContent;
        }
        public string UQToWebview(string imagePath, string orderNo, string deliveryDestination, string poNo,
                string sampleNo, string materialCode, string sizeName, string colorName, string colorCode, double quantity,
                string keyNo, string contractNo, double quantityInBox)
        {
            byte[] imageBytes = File.ReadAllBytes(imagePath);
            string base64 = Convert.ToBase64String(imageBytes);
            string imageSrc = $"data:image/png;base64,{base64}";
            string quantityDisplay = quantity.ToString("N0");
            if (materialCode.StartsWith("HT") || materialCode.StartsWith("WT"))
            {
                quantityDisplay = quantityInBox.ToString("N0") + "/" + quantity.ToString("N0");
            }
            string htmlContent = $@"
                    <div class=""crosshair-h"" id=""crosshair-h""></div>
                    <div class=""crosshair-v"" id=""crosshair-v""></div>
                    <div class=""outer-box box-left"">
                     <table class=""form-table"">
                         <tr>
                             <td class=""label"">Customer</td>
                             <td colspan=""5"" style=""text-align:left;"">
                                <div id=""fit-text"" class=""fit-text"" style=""width:245px;"">
                                    <span>{deliveryDestination}</span>
                                </div>
                             </td>
                         </tr>
                         <tr style=""height: 4px;""><td></td></tr>
                         <tr>
                             <td class=""label"">Order No</td>
                             <td colspan=""3"">
                                <div id=""fit-text"" class=""fit-text"" style=""width:160px;"">
                                   <span>{poNo}</span>
                                </div>
                            </td>
                             <td colspan=""2"" class=""content-bold"" style=""text-align:right"">
                               {(!string.IsNullOrEmpty(deliveryDestination) && deliveryDestination.StartsWith("REGENT GARMENT", StringComparison.OrdinalIgnoreCase) ? contractNo : "")} 
                             </td>
                         </tr>
                         <tr style=""height: 4px;""><td></td></tr>
                         <tr>
                             <td class=""label"">Item Code</td>
                             <td colspan=""5"" style=""padding-left:8px;"">
                                <div id=""fit-text"" class=""fit-text"" style=""width:235px;"">
                                   <span>{sampleNo}/{materialCode}</span>
                                </div>
                            </td>
                         </tr>
                         <tr style=""height: 4px;""><td></td></tr>
                         <tr>
                             <td class=""label"">Quantity</td>
                             <td colspan=""1"" style=""padding-left:8px;"">Col. {colorCode}</td>
                             <td colspan=""1"">Size {sizeName}</td>
                             <td colspan=""2"" style=""text-align:right"">
                                <span class=""content-bold"">{quantityDisplay}</span>
                            </td>
                             <td colspan=""1"" class=""pcs"">pcs</td>
                         </tr>
                         <tr style=""height: 4px;""><td></td></tr>
                         <tr>
                             <td class=""label""></td>
                             <td colspan=""3"" class=""content-bold"" style=""padding-left:8px;""> {orderNo}</td>
                             <td colspan=""2"" style=""text-align:left"">{keyNo}</td>
                         </tr>
                     </table>
                     <div class=""notice"">
                         Please open and confirm this package immediately on receipt.
                     </div>
                     <div class=""footer"">
                         <span>MADE IN VIETNAM</span>
                     </div>
                     <div class=""container"">
                         <div style=""width: 70%; display: inline-block;"">
                             <img src=""{imageSrc}"" class=""address-img"" />
                         </div>
                         <div class=""parent-container"">
                             <div class=""qr-box""></div>
                         </div>
                     </div>
                 </div>";
            return htmlContent;
        }

        public string MonbellToWebview(string imagePath, string orderNo, string poNo, string styleNo, string itemCode, string materialCode, double quantity, string color, string size, double sizeQuantity)
        {
            byte[] imageBytes = File.ReadAllBytes(imagePath);
            string base64 = Convert.ToBase64String(imageBytes);
            string imageSrc = $"data:image/png;base64,{base64}";
            string htmlContent = $@"
            <div class=""crosshair-h"" id=""crosshair-h""></div>
            <div class=""crosshair-v"" id=""crosshair-v""></div>
            <div class=""outer-box box-left"">
                <table class=""form-table"">
                    <tr>
                        <td colspan=""1"" class=""label"">Order No </td>
                        <td colspan=""4"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:205px;"">
                                <span style=""padding-left:1.3em;"">{orderNo}</span>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">TENTAC(PO NO)</td>
                        <td colspan=""1"" class=""content-bold"" style=""width:60px;"">{poNo}</td>
                        <td colspan=""1"" class=""title"">TTL Style</td>
                        <td colspan=""1"" style=""text-align:right"">
                            <div id=""fit-text"" class="" fit-text"" style=""width:55px"">
                                <label style=""text-align:right"">{quantity.ToString("N0")}</label>
                            </div>
                        </td>
                        <td colspan=""1"" style=""width:20px; text-align:right"">PCS</td>
                    </tr>
                    <tr style=""height: 4px;""></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">ITEM Code/Material</td>
                        <td colspan=""4"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:205px;"">
                                <label>{itemCode} / {materialCode}</label>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">Style No</td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:60px"">
                                <label>{styleNo}</label>
                            </div>
                        </td>
                        <td colspan=""1"" class=""title"">Size</td>
                        <td colspan=""2"" class=""content-bold"" style=""text-align:right"">{size}</td>
                    </tr>
                    <tr style="" height: 4px;""></tr>
                    <tr>
                        <td colspan=""1"" class="" label"">Color</td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class=""fit-text"" style="" width:60px;"">
                                <label>{color}</label>
                            </div>
                        </td>
                        <td colspan=""1"" class=""title"">QTY</td>
                        <td colspan=""1"" style=""text-align:right"">
                            <div id=""fit-text"" class="" fit-text"" style=""width:55px;"">
                                <label style=""text-align:right"" class=""content-bold"">{sizeQuantity.ToString("N0")}</label>
                            </div>
                        </td>
                        <td colspan=""1"" style=""width:20px;"">PCS</td>
                    </tr>
                </table>
                <div class=""notice"">
                    Please open and confirm this package immediately on receipt.
                </div>
                <div class=""footer"">
                    <span>MADE IN VIETNAM</span>
                </div>
                <div class=""container"">
                    <div style=""width: 65%; display: inline-block;"">
                        <img src=""{imageSrc}"" class="" address-img"" />
                    </div>
                    <div class=""parent-container"">
                        <div class=""qr-box""></div>
                    </div>
                </div>
            </div>";
            return htmlContent;
        }

        //private float ResizeTextToFit(string contentLayout, float maxWidth)
        //{
        //    float fontSize = 12f;
        //    float minFontSize = 6f;

        //    string fontFilePath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\MS-UIGothic.ttf";
        //    BaseFont baseFont = BaseFont.CreateFont(fontFilePath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

        //    if (!string.IsNullOrEmpty(contentLayout))
        //    {
        //        while (MeasureTextWidth(contentLayout, baseFont, fontSize) > maxWidth && fontSize > minFontSize)
        //        {
        //            fontSize -= 0.5f;
        //        }
        //    }

        //    return fontSize;
        //}

        //private float MeasureTextWidth(string text, BaseFont baseFont, float fontSize)
        //{
        //    // Use BaseFont's GetWidthPoint method to calculate the width of the text
        //    return baseFont.GetWidthPoint(text, fontSize);
        //}
        public float ResizeTextToFit(string text, string fontPath, float maxWidth, float maxFontSize = 12f, float minFontSize = 6f)
        {
            PdfFont font = PdfFontFactory.CreateFont(fontPath, PdfEncodings.IDENTITY_H);
            // Tạo một document giả để layout
            using (var ms = new MemoryStream())
            {
                PdfDocument pdfDoc = new PdfDocument(new PdfWriter(ms));
                Document doc = new Document(pdfDoc);

                for (float fontSize = maxFontSize; fontSize >= minFontSize; fontSize -= 0.5f)
                {
                    Text t = new Text(text).SetFont(font).SetFontSize(fontSize);
                    TextRenderer renderer = (TextRenderer)t.CreateRendererSubTree();
                    renderer.SetParent(doc.GetRenderer());

                    LayoutResult result = renderer.Layout(new LayoutContext(new LayoutArea(0, new Rectangle(1000, 1000))));
                    float width = result.GetOccupiedArea().GetBBox().GetWidth();

                    if (width <= maxWidth)
                    {
                        pdfDoc.Close(); // Giải phóng bộ nhớ
                        return fontSize;
                    }
                }

                pdfDoc.Close();
            }

            return minFontSize;
        }

        private string GenerateSafeFileName(string originalPath)
        {
            string baseName = Path.GetFileNameWithoutExtension(originalPath);
            string extension = Path.GetExtension(originalPath);
            string safeName = $@"PackingListXXX_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + extension;
            return Path.Combine(Path.GetTempPath(), safeName);
        }

        private bool ContainsJapaneseCharacters(string input)
        {
            return input.Any(c => (c >= '\u3040' && c <= '\u30FF') || // Hiragana & Katakana
                                  (c >= '\u4E00' && c <= '\u9FBF'));  // Kanji
        }

        public string PrepareFile(string originalPath)
        {
            string fileName = Path.GetFileName(originalPath);
            if (ContainsJapaneseCharacters(fileName) == true)
            {
                string newPath = GenerateSafeFileName(originalPath);
                File.Copy(originalPath, newPath, true);
                return newPath;
            }
            return originalPath;
        }
    }
}
