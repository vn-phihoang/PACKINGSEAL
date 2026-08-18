using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Layout;
using iText.Layout.Renderer;
using PACKINGSEAL.Models;
using QRCoder;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Path = System.IO.Path;
using PdfFont = iText.Kernel.Font.PdfFont;


namespace PACKINGSEAL
{
    public class HtmlResponse
    {
        private Assembly assembly = Assembly.GetEntryAssembly();

        public string UQToPDF(UniqloModel Uniqlo)
        {
            string fontPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\MS-UIGothic.ttf";
            string fontBold = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\NotoSansJP-ExtraBold.ttf";
            float fontSizeCustomer = ResizeTextToFit(Uniqlo.DeliveryDestination, fontPath, 240);
            float fontSizeOrder = ResizeTextToFit(Uniqlo.PoNo, fontPath, 145);
            float fontSizeItem = ResizeTextToFit($"{Uniqlo.SampleNo}/{Uniqlo.MaterialCode}", fontPath, 230);
            string quantityDisplay = Uniqlo.Quantity.ToString("N0");
            if (Uniqlo.MaterialCode.StartsWith("HT") || Uniqlo.MaterialCode.StartsWith("WT"))
            {
                quantityDisplay = Uniqlo.QuantityInBox.ToString("N0") + "/" + Uniqlo.Quantity.ToString("N0");
            }
            float fontSizequantityDisplay = ResizeTextToFit($"{quantityDisplay}", fontBold, 90);
            string boxSide = Uniqlo.Index % 2 == 0 ? "box-left" : "box-right";
            string pageOpenTag = Uniqlo.Index % 2 == 0 ? "<page size=\"A4\">\r\n" : "";
            string pageCloseTag = Uniqlo.Index % 2 == 1 ? "</page>\r\n" : "";
            string tableContent = $@"{pageOpenTag}  
            <div class=""outer-box {boxSide}"">  
                <table class=""form-table""> 
                    <tr>  
                        <td class=""label"">Customer</td>  
                        <td colspan=""5"" style=""text-align:left; font-size:{fontSizeCustomer}px;"">
                            {Uniqlo.DeliveryDestination}
                        </td>  
                    </tr>  
                    <tr style=""height: 4px;""><td></td></tr>  
                    <tr>  
                        <td class=""label"">Order No</td>  
                        <td colspan=""3"" style=""font-size:{fontSizeOrder}px;"">{Uniqlo.PoNo}</td>  
                        <td colspan=""2"" class=""content-bold"" style=""text-align:right""> 
                            {(!string.IsNullOrEmpty(Uniqlo.DeliveryDestination) && Uniqlo.DeliveryDestination.StartsWith("REGENT GARMENT", StringComparison.OrdinalIgnoreCase) ? Uniqlo.ContractNo : "")} 
                        </td> 
                    </tr>  
                    <tr style=""height: 4px;""><td></td></tr>  
                    <tr>  
                        <td class=""label"">Item Code</td>  
                        <td colspan=""5"" style=""padding-left:8px; width:235px; font-size:{fontSizeItem}px;"">
                            {Uniqlo.SampleNo}/{Uniqlo.MaterialCode}
                        </td>  
                    </tr>  
                    <tr style=""height: 4px;""><td></td></tr>  
                    <tr>  
                        <td class=""label"">Quantity</td>  
                        <td style=""padding-left:8px;"">Col. {Uniqlo.ColorCode}</td>  
                        <td>Size {Uniqlo.SizeName}</td>  
                        <td colspan=""3"" style=""text-align:right; width:110px"">
                            <label class=""content-bold"" style=""font-size:{fontSizequantityDisplay}px;"">{quantityDisplay}</label> 
                            <label>pcs</label> 
                        </td>
                    </tr>  
                    <tr style=""height: 4px;""><td></td></tr>  
                    <tr>  
                        <td class=""label""></td>  
                        <td colspan=""3"" class=""content-bold"" style=""padding-left:8px;"">{Uniqlo.OrderNo}</td>  
                        <td colspan=""2"" style=""text-align:left"">{Uniqlo.KeyNo}</td>  
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
                        <img src=""{Uniqlo.ImagePath}"" class=""address-img"" />  
                    </div>  
                    <div class=""parent-container"">  
                        <div class=""qr-box""></div>  
                    </div>  
                </div>  
            </div>
            {pageCloseTag}";
            return tableContent;
        }

        public string MonbellToPDF(MontbellModel montbellModel)
        {
            string fontPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\MS-UIGothic.ttf";
            string fontBoldPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\NotoSansJP-ExtraBold.ttf";
            float fontSizeOrderNo = ResizeTextToFit(montbellModel.OrderNo, fontPath, 190);
            float fontQTY = ResizeTextToFit(montbellModel.Quantity.ToString("N0"), fontPath, 60);
            float fontSizeStyleNo = ResizeTextToFit(montbellModel.StyleNo, fontPath, 60);
            float fontSizeColor = ResizeTextToFit(montbellModel.Color, fontPath, 60);
            float fontSizeQTY;
            if (montbellModel.TypeProduct == 1)
            {
                fontSizeQTY = ResizeTextToFit(montbellModel.QuantityInBox.ToString("N0") + "/" + montbellModel.SizeQuantity.ToString("N0"), fontBoldPath, 60);
            }
            else
            {
                fontSizeQTY = ResizeTextToFit(montbellModel.SizeQuantity.ToString("N0"), fontBoldPath, 60);
            }
            string quantityDisplay;
            if (montbellModel.TypeProduct == 1)
            {
                quantityDisplay = montbellModel.QuantityInBox.ToString("N0") + "/" + montbellModel.SizeQuantity.ToString("N0");
            }
            else
            {
                quantityDisplay = montbellModel.SizeQuantity.ToString("N0");
            }
            float fontItemcodeMametan = ResizeTextToFit(montbellModel.ItemCode + " / " + montbellModel.MaterialCode, fontPath, 190);
            string boxSide = montbellModel.Index % 2 == 0 ? "box-left" : "box-right";
            string pageOpenTag = montbellModel.Index % 2 == 0 ? "<page size=\"A4\">\r\n" : "";
            string pageCloseTag = montbellModel.Index % 2 == 1 ? "</page>\r\n" : "";
            string tableContent = $@"{pageOpenTag}  
            <div class=""outer-box {boxSide}"">
                <table class=""form-table"">
                    <tr>
                        <td colspan=""1"" class=""label"">Order No </td>
                        <td colspan=""4"" style=""padding-left:1.3em; font-size:{fontSizeOrderNo}px;"">
                             {montbellModel.OrderNo}
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">TENTAC(PO NO)</td>
                        <td colspan=""1"" class=""content-bold"" style=""width:60px;"">{montbellModel.PoNo}</td>
                        <td colspan=""1"" class=""title"">TTL Style</td>
                        <td colspan=""1"" style=""text-align:right; font-size:{fontQTY}px;"">
                           <label style=""text-align:right"">{montbellModel.Quantity.ToString("N0")}</label>
                        </td>
                        <td colspan=""1"" style=""width:20px; text-align:right"">PCS</td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">ITEM Code/Material</td>
                        <td colspan=""4"" style=""font-size:{fontItemcodeMametan}px;"">
                            <label>{montbellModel.ItemCode} / {montbellModel.MaterialCode}</label>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">Style No</td>
                        <td colspan=""1"" style=""font-size:{fontSizeStyleNo}px;"">
                            {montbellModel.StyleNo}
                        </td>
                        <td colspan=""1"" class=""title"">Size</td>
                        <td colspan=""2"" class=""content-bold"" style=""text-align:right"">{montbellModel.Size}</td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">Color</td>
                        <td colspan=""1"" style=""font-size:{fontSizeColor}px;"">
                             {montbellModel.Color}
                        </td>
                        <td colspan=""1"" class=""title"">QTY</td>
                        <td colspan=""1"" style=""text-align:right;"">
                            <label class=""content-bold"" style=""font-size:{fontSizeQTY}px;"">{quantityDisplay}</label>
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
                        <img src=""{montbellModel.ImagePath}"" class=""address-img"" />
                    </div>
                    <div class=""parent-container"">
                        <div class=""qr-box""></div>
                    </div>
                </div>
            </div>
            {pageCloseTag}";
            return tableContent;
        }

        public string MujiToPDF(MujiModel mujiModel)
        {
            string fontPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\MS-UIGothic.ttf";
            string fontBold = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\NotoSansJP-ExtraBold.ttf";
            float fontSizeCustomer = ResizeTextToFit(mujiModel.DeliveryDestination, fontPath, 240);
            float fontSizeOrder = ResizeTextToFit(mujiModel.PoNo, fontPath, 60);
            float fontSizePOCustomer = ResizeTextToFit(mujiModel.POCustomer, fontPath, 165);
            float fontSizeItem = ResizeTextToFit($"{mujiModel.ItemCode}", fontPath, 230);
            string quantityDisplay = mujiModel.Quantity.ToString("N0");
            if (mujiModel.QuantityInBox == 0)
            {
                quantityDisplay = mujiModel.Quantity.ToString("N0");
            }
            else
            {
                quantityDisplay = mujiModel.QuantityInBox.ToString("N0") + "/" + mujiModel.Quantity.ToString("N0");
            }
            float fontSizequantityDisplay = ResizeTextToFit($"{quantityDisplay}", fontBold, 210);
            string boxSide = mujiModel.Index % 2 == 0 ? "box-left" : "box-right";
            string pageOpenTag = mujiModel.Index % 2 == 0 ? "<page size=\"A4\">\r\n" : "";
            string pageCloseTag = mujiModel.Index % 2 == 1 ? "</page>\r\n" : "";
            string tableContent = $@"{pageOpenTag}  
            <div class=""outer-box {boxSide}"">  
                <table class=""form-table""> 
                    <tr>  
                        <td class=""label"">Customer</td>  
                        <td colspan=""3"" style=""text-align:left; font-size:{fontSizeCustomer}px;"">
                            {mujiModel.DeliveryDestination}
                        </td>  
                    </tr>  
                    <tr style=""height: 4px;""><td></td></tr>  
                    <tr>  
                        <td class=""label"">Order No</td>  
                        <td colspan=""1"" style=""font-size:{fontSizeOrder}px; width:65px;"">{mujiModel.PoNo}</td> 
                        <td colspan=""2"" style=""text-align:right; font-size:{fontSizePOCustomer}px; width:170px;"">{mujiModel.POCustomer}</td> 
                    </tr>  
                    <tr style=""height: 4px;""><td></td></tr>  
                    <tr>  
                        <td class=""label"">Item Code</td>  
                        <td colspan=""3"" style=""padding-left:8px; width:235px; font-size:{fontSizeItem}px;"">
                            {mujiModel.ItemCode}
                        </td>  
                    </tr>  
                    <tr style=""height: 4px;""><td></td></tr>  
                    <tr>  
                        <td class=""label"">Quantity</td>  
                        <td colspan=""3"" style=""text-align:right; width:110px"">
                            <label class=""content-bold"" style=""font-size:{fontSizequantityDisplay}px;"">{quantityDisplay}</label> 
                            <label>PCS</label> 
                        </td>
                    </tr>  
                    <tr style=""height: 4px;""><td></td></tr>  
                    <tr>  
                        <td class=""label""></td>  
                        <td colspan=""1"" style=""padding-left:8px;"">{mujiModel.KeyNo}</td>  
                        <td colspan=""2"" class=""content-bold"" style=""text-align:right"">FSC Mix Credit</td>
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
                        <img src=""{mujiModel.ImagePath}"" class=""address-img"" />  
                    </div>  
                    <div class=""parent-container"">  
                        <div class=""qr-box""></div>  
                    </div>  
                </div>  
            </div>
            {pageCloseTag}";
            return tableContent;
        }

        public string SakuraiToPDF(SakuraiModel sakuraiModel)
        {
            string qrData = $"{sakuraiModel.Date}#{sakuraiModel.InvoiNo}#{Convert.ToInt32(sakuraiModel.CartonNo.ToString()).ToString("D3")}";
            using (var qrGenerator = new QRCodeGenerator())
            using (var qrCodeData = qrGenerator.CreateQrCode(qrData, QRCodeGenerator.ECCLevel.Q))
            using (var qrCode = new QRCoder.QRCode(qrCodeData))
            using (var qrImage = qrCode.GetGraphic(20)) // 20 = độ phân giải
            {
                // Chuyển ảnh QR sang Base64
                using (var ms = new MemoryStream())
                {
                    qrImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    string base64Qr = Convert.ToBase64String(ms.ToArray());
                    string fontPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\BIZ-UDGothic-Bold.ttf";
                    float fontSizeInvoiNo = ResizeTextToFit(sakuraiModel.InvoiNo, fontPath, 135, 20f);

                    // Nhúng trực tiếp vào HTML
                    string boxSide = sakuraiModel.Index % 2 == 0 ? "box-left" : "box-right";
                    string pageOpenTag = sakuraiModel.Index % 2 == 0 ? "<page size=\"A4\">\r\n" : "";
                    string pageCloseTag = sakuraiModel.Index % 2 == 1 ? "</page>\r\n" : "";

                    string tableContent = $@"{pageOpenTag}
                    <div class=""outer-box {boxSide}"" style=""height:365px; width: 245px;"">
                        <div class='qr-box' style='text-align:center;'>
                            <img src='data:image/png;base64,{base64Qr}' width='240' height='240'/>
                        </div>
                        <div class=""box-content"">
                            <label class=""title"">DATE :</label>
                            <label class=""content"">{sakuraiModel.Date}</label>
                        </div>
                        <div class=""box-content"">
                            <label class=""title"">I/V No. :</label>
                            <label class=""content"" style=""font-size:{fontSizeInvoiNo}px;"">{sakuraiModel.InvoiNo}</label>
                        </div>
                        <div class=""box-content"">
                            <label class=""title"">CARTON No. :</label>
                            <label class=""content"">{sakuraiModel.CartonNo}</label>
                        </div>
                    </div>
                    {pageCloseTag}";

                    return tableContent;
                }
            }
        }

        public string ThermalUAToPDF(ThermalUnderamour thermalUnderamour)
        {
            string fontPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\MS-UIGothic.ttf";
            string fontBoldPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\NotoSansJP-ExtraBold.ttf";
            float fontSizeCustomer = ResizeTextToFit(thermalUnderamour.Customer, fontPath, 235);
            float fontSizeOrderNo = ResizeTextToFit(thermalUnderamour.OrderNo, fontPath, 75);
            float fontSizeStyleNo = ResizeTextToFit(thermalUnderamour.StyleNo, fontPath, 75);
            float fontSizePONo = ResizeTextToFit(thermalUnderamour.PONo, fontBoldPath, 80);
            float fontSizeProduct = ResizeTextToFit(thermalUnderamour.Product, fontPath, 75);
            float fontSizeDescription = ResizeTextToFit(thermalUnderamour.Description, fontPath, 160);
            float fontSizeColorSize = ResizeTextToFit(thermalUnderamour.ColorCode + " / " + thermalUnderamour.SizeName, fontBoldPath, 155);
            float fontSizeJan = ResizeTextToFit(thermalUnderamour.JAN, fontPath, 80);
            string boxSide = thermalUnderamour.Index % 2 == 0 ? "box-left" : "box-right";
            string pageOpenTag = thermalUnderamour.Index % 2 == 0 ? "<page size=\"A4\">\r\n" : "";
            string pageCloseTag = thermalUnderamour.Index % 2 == 1 ? "</page>\r\n" : "";
            string tableContent = $@"{pageOpenTag}  
            <div class=""outer-box {boxSide}"">
                <table class=""form-table"">
                    <tr>
                        <td colspan=""1"" class=""label"">Customer</td>
                        <td colspan=""3"">
                            <label style = ""font-size:{fontSizeCustomer}px;"">{thermalUnderamour.Customer}</label>
                        </td>
                    </tr>
                  <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">Order No</td>
                        <td colspan=""1"">
                           <label style = ""font-size:{fontSizeOrderNo}px;"">{thermalUnderamour.OrderNo}</label>
                        </td>
                        <td colspan=""1"">
                            <label style = ""font-size:{fontSizeStyleNo}px;"">{thermalUnderamour.StyleNo}</label>
                        </td>
                        <td colspan=""1"">
                            <div style = ""text-align:right;"">
                                <label style = ""font-size:{fontSizePONo}px;"" class=""content-bold"">{thermalUnderamour.PONo}</label>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">ITEM Code</td>
                        <td colspan=""1"">
                            <label style = ""font-size:{fontSizeProduct}px;"">{thermalUnderamour.Product}</label>
                        </td>
                        <td colspan=""2"">
                            <div style = ""text-align:right;"">
                                <label style = ""font-size:{fontSizeDescription}px;"">{thermalUnderamour.Description}</label>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">Quantity</td>
                        <td colspan=""3"">
                            <div style=""text-align:right;"">
                                <label class=""content-bold"" style=""text-align:right"">{thermalUnderamour.Quantity.ToString("N0")}</label>
                                <label class=""title"">pcs</label>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class="" label"">Color/Size</td>
                        <td colspan=""2"">
                            <label style = ""font-size:{fontSizeColorSize}px;"" class=""content-bold"">{thermalUnderamour.ColorCode} / {thermalUnderamour.SizeName}</label>
                        </td>
                        <td colspan=""1"">
                            <div style=""text-align:right;"">
                                <label style=""text-align:right; font-size:{fontSizeJan}px;"">{thermalUnderamour.JAN}</label>
                            </div>
                        </td>
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
                        <img src=""{thermalUnderamour.ImagePath}"" class=""address-img"" />
                    </div>
                    <div class=""parent-container"">
                        <div class=""qr-box""></div>
                    </div>
                </div>
            </div>
            {pageCloseTag}";
            return tableContent;
        }

        public string KonicaUAToPDF(KonicaUnderamour konicaUnderamour)
        {
            string fontPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\MS-UIGothic.ttf";
            string fontBoldPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\NotoSansJP-ExtraBold.ttf";
            float fontSizeCustomer = ResizeTextToFit(konicaUnderamour.Customer, fontPath, 235);
            float fontSizeOrderNo = ResizeTextToFit(konicaUnderamour.OrderNo, fontPath, 115);
            float fontSizePONo = ResizeTextToFit(konicaUnderamour.PONo, fontBoldPath, 115);
            float fontSizeProduct = ResizeTextToFit(konicaUnderamour.Product, fontPath, 115);
            float fontSizeDescription = ResizeTextToFit(konicaUnderamour.Description, fontPath, 115);
            float fontSizeColor = ResizeTextToFit(konicaUnderamour.ColorName, fontBoldPath, 115);
            float fontSizeSizeName = ResizeTextToFit(konicaUnderamour.SizeName, fontBoldPath, 115);
            string boxSide = konicaUnderamour.Index % 2 == 0 ? "box-left" : "box-right";
            string pageOpenTag = konicaUnderamour.Index % 2 == 0 ? "<page size=\"A4\">\r\n" : "";
            string pageCloseTag = konicaUnderamour.Index % 2 == 1 ? "</page>\r\n" : "";
            string tableContent = $@"{pageOpenTag}  
            <div class=""outer-box {boxSide}"">
                <table class=""form-table"">
                    <tr>
                        <td colspan=""1"" class=""label"">Customer</td>
                        <td colspan=""2"">
                            <label style = ""font-size:{fontSizeCustomer}px;"">{konicaUnderamour.Customer}</label>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">Order No</td>
                        <td colspan=""1"">
                           <label style = ""font-size:{fontSizeOrderNo}px;"">{konicaUnderamour.OrderNo}</label>
                        </td>
                        <td colspan=""1"" style = ""text-align:right;"">
                            <label style = ""font-size:{fontSizePONo}px;"" class=""content-bold"">{konicaUnderamour.PONo}</label>
                        </td>
                    </tr>
                   <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">ITEM Code</td>
                        <td colspan=""1"">
                            <label style = ""font-size:{fontSizeProduct}px;"">{konicaUnderamour.Product}</label>
                        </td>
                        <td colspan=""1"">
                            <div style = ""text-align:right;"">
                                <label style = ""font-size:{fontSizeDescription}px;"">{konicaUnderamour.Description}</label>
                            </div>
                        </td>
                    </tr>
                   <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">Quantity</td>
                        <td colspan=""2"">
                            <div style=""text-align:right;"">
                                <label class=""content-bold"" style=""text-align:right"">{konicaUnderamour.Quantity.ToString("N0")}</label>
                                <label class=""title"">pcs</label>
                            </div>
                        </td>
                    </tr>
                   <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class="" label"">Color/Size</td>
                        <td colspan=""1"">
                            <label style = ""font-size:{fontSizeColor}px;"" class=""content-bold"">{konicaUnderamour.ColorName}</label>
                        </td>
                        <td colspan=""1"">
                            <div style=""text-align:right;"">
                                <label style = ""font-size:{fontSizeSizeName}px;"" class=""content-bold"">{konicaUnderamour.SizeName}</label>
                            </div>
                        </td>
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
                        <img src=""{konicaUnderamour.ImagePath}"" class=""address-img"" />
                    </div>
                    <div class=""parent-container"">
                        <div class=""qr-box""></div>
                    </div>
                </div>
            </div>
            {pageCloseTag}";
            return tableContent;
        }

        public string AsicsToPDF(AsicsModel asicsModel)
        {
            string fontPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\MS-UIGothic.ttf";
            string fontBoldPath = $@"{Path.GetDirectoryName(assembly.Location)}\Fonts\NotoSansJP-ExtraBold.ttf";
            float fontSizeCustomer = ResizeTextToFit(asicsModel.Customer, fontPath, 225);
            float fontSizeOrderNo = ResizeTextToFit(asicsModel.OrderNo, fontPath, 55);
            float fontSizeSKUNo = ResizeTextToFit(asicsModel.SKUNo, fontPath, 45);
            float fontSizeContractNo = ResizeTextToFit(asicsModel.ContractNo, fontPath, 95);
            float fontSizeSampleNo = ResizeTextToFit(asicsModel.SampleNo, fontPath, 55);
            float fontSizeColorCode = ResizeTextToFit(asicsModel.ColorCode, fontPath, 25);
            float fontSizeSizeName = ResizeTextToFit(asicsModel.SizeName, fontPath, 25);
            float fontSizeQuantity = ResizeTextToFit(asicsModel.Quantity.ToString("N0"), fontBoldPath, 70);
            float fontSizeKeyNo = ResizeTextToFit(asicsModel.KeyNo, fontPath, 225);
            float fontSizeVendors = ResizeTextToFit(asicsModel.Vendors, fontPath, 225);
            float fontSizeItemCode = ResizeTextToFit(asicsModel.ItemCode, fontPath, 225);
            float fontSizeProductName = ResizeTextToFit(asicsModel.ProductName, fontPath, 225);
            //float fontSizeStyleNo = ResizeTextToFit(thermalUnderamour.StyleNo, fontPath, 75);
            //float fontSizePONo = ResizeTextToFit(thermalUnderamour.PONo, fontBoldPath, 80);
            //float fontSizeProduct = ResizeTextToFit(thermalUnderamour.Product, fontPath, 75);
            //float fontSizeDescription = ResizeTextToFit(thermalUnderamour.Description, fontPath, 160);
            //float fontSizeColorSize = ResizeTextToFit(thermalUnderamour.ColorCode + " / " + thermalUnderamour.SizeName, fontBoldPath, 155);
            //float fontSizeJan = ResizeTextToFit(thermalUnderamour.JAN, fontPath, 80);
            string boxSide = asicsModel.Index % 2 == 0 ? "box-left" : "box-right";
            string pageOpenTag = asicsModel.Index % 2 == 0 ? "<page size=\"A4\">\r\n" : "";
            string pageCloseTag = asicsModel.Index % 2 == 1 ? "</page>\r\n" : "";
            string tableContent = $@"{pageOpenTag}  
            <div class=""outer-box {boxSide}"">
                <table class=""form-table"">
                    <tr>
                        <td class=""label"">Customer</td>
                        <td colspan=""5"" style=""text-align:left;"">
                            <label style = ""font-size:{fontSizeCustomer}px;"">{asicsModel.Customer}</label>
                        </td>
                    </tr>
                    <tr style=""height: 2px;""><td></td></tr>
                    <tr>
                        <td class=""label"">Order No</td>
                        <td colspan=""1"">
                            <label style = ""font-size:{fontSizeOrderNo}px;"">{asicsModel.OrderNo}</label>
                        </td>
                        <td colspan=""2"">
                            <label style = ""font-size:{fontSizeSKUNo}px;"">{asicsModel.SKUNo}</label>
                        </td>
                        <td colspan=""2"" style=""text-align:right"">
                            <label style = ""font-size:{fontSizeContractNo}px;"">{asicsModel.ContractNo}</label>
                        </td>
                    </tr>
                    <tr style=""height: 2px;""><td></td></tr>
                    <tr>
                        <td class=""label"" rowspan=""2"">Item Code</td>
                        <td colspan=""4"" rowspan=""2"">
                            <div style=""white-space: normal; word-wrap: break-word;"">
                                <label style = ""font-size:{fontSizeItemCode}px;"">{asicsModel.ItemCode}</label>
                                <label style = ""font-size:{fontSizeProductName}px;"">{asicsModel.ProductName}</label>
                            </div>
                        </td>
                    </tr>
                    <tr><td></td></tr>
                    <tr style=""height: 2px;""><td></td></tr>
                    <tr>
                        <td class=""label"">Quantity</td>
                        <td colspan=""1"">
                            <label style = ""font-size:{fontSizeSampleNo}px;"">{asicsModel.SampleNo}</label>
                        </td>
                        <td colspan=""1"">
                            <label style = ""font-size:{fontSizeColorCode}px;"">{asicsModel.ColorCode}</label>
                        </td>
                        <td colspan=""1"">
                            <label style = ""font-size:{fontSizeSizeName}px;"">{asicsModel.SizeName}</label>
                        </td>
                        <td colspan=""1"" style=""text-align:right;"">
                            <label class=""content-bold"" style = ""font-size:{fontSizeQuantity}px;"">{asicsModel.Quantity}</label>
                        </td>
                        <td colspan=""1"" class=""pcs"">PCS</td>
                    </tr>
                    <tr style=""height: 2px;""><td></td></tr>
                    <tr>
                        <td class=""label"" style=""text-align:center"">{asicsModel.POLine}</td>
                        <td colspan=""5"">
                            <label style = ""font-size:{fontSizeKeyNo}px;"">{asicsModel.KeyNo}</label>
                        </td>
                    </tr>
                    <tr style=""height: 2px;""><td></td></tr>
                    <tr>
                        <td class=""label"">Vendors Ref.</td>
                        <td colspan=""5"">
                            <label style = ""font-size:{fontSizeVendors}px;"">{asicsModel.Vendors}</label>
                        </td>
                    </tr>
                </table>
                <div class=""notice"">
                    <span>Please open and confirm this package immediately on receipt.</span>
                </div>
                <div class=""footer"">
                    <span>MADE IN VIETNAM</span>
                    <div class=""parent-container"">
                        <div class=""qr-box""></div>
                    </div>
                </div>
                <div class=""container"">
                    <div style=""width: 30%; display: inline-block;"">
                        <img src=""{asicsModel.ImagePath}"" class="" address-img"" />
                    </div>
                </div>
            </div>
            {pageCloseTag}";
            return tableContent;
        }

        public string UQToWebview(UniqloModel Uniqlo)
        {
            byte[] imageBytes = File.ReadAllBytes(Uniqlo.ImagePath);
            string base64 = Convert.ToBase64String(imageBytes);
            string imageSrc = $"data:image/png;base64,{base64}";
            string quantityDisplay = Uniqlo.Quantity.ToString("N0");
            if (Uniqlo.MaterialCode.StartsWith("HT") || Uniqlo.MaterialCode.StartsWith("WT"))
            {
                quantityDisplay = Uniqlo.QuantityInBox.ToString("N0") + "/" + Uniqlo.Quantity.ToString("N0");
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
                                    <span>{Uniqlo.DeliveryDestination}</span>
                                </div>
                             </td>
                         </tr>
                         <tr style=""height: 4px;""><td></td></tr>
                         <tr>
                             <td class=""label"">Order No</td>
                             <td colspan=""3"">
                                <div id=""fit-text"" class=""fit-text"" style=""width:160px;"">
                                   <span>{Uniqlo.PoNo}</span>
                                </div>
                            </td>
                             <td colspan=""2"" class=""content-bold"" style=""text-align:right"">
                               {(!string.IsNullOrEmpty(Uniqlo.DeliveryDestination) && Uniqlo.DeliveryDestination.StartsWith("REGENT GARMENT", StringComparison.OrdinalIgnoreCase) ? Uniqlo.ContractNo : "")} 
                             </td>
                         </tr>
                         <tr style=""height: 4px;""><td></td></tr>
                         <tr>
                             <td class=""label"">Item Code</td>
                             <td colspan=""5"" style=""padding-left:8px;"">
                                <div id=""fit-text"" class=""fit-text"" style=""width:235px;"">
                                   <span>{Uniqlo.SampleNo}/{Uniqlo.MaterialCode}</span>
                                </div>
                            </td>
                         </tr>
                         <tr style=""height: 4px;""><td></td></tr>
                         <tr>
                             <td class=""label"">Quantity</td>
                             <td colspan=""1"" style=""padding-left:8px;"">Col. {Uniqlo.ColorCode}</td>
                             <td colspan=""1"">Size {Uniqlo.SizeName}</td>
                             <td colspan=""2"" style=""text-align:right"">
                                <span class=""content-bold"">{quantityDisplay}</span>
                            </td>
                             <td colspan=""1"" class=""pcs"">pcs</td>
                         </tr>
                         <tr style=""height: 4px;""><td></td></tr>
                         <tr>
                             <td class=""label""></td>
                             <td colspan=""3"" class=""content-bold"" style=""padding-left:8px;""> {Uniqlo.OrderNo}</td>
                             <td colspan=""2"" style=""text-align:left"">{Uniqlo.KeyNo}</td>
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
                             <img src=""{imageSrc}"" class=""address-img"" />
                         </div>
                         <div class=""parent-container"">
                             <div class=""qr-box""></div>
                         </div>
                     </div>
                 </div>";
            return htmlContent;
        }

        public string SakuraiToWebview(SakuraiModel sakuraiModel)
        {
            string qrData = $"{sakuraiModel.Date}#{sakuraiModel.InvoiNo}#{Convert.ToInt32(sakuraiModel.CartonNo.ToString()).ToString("D3")}";
            string htmlContent = $@"
            <div class=""crosshair-h"" id=""crosshair-h""></div>
            <div class=""crosshair-v"" id=""crosshair-v""></div>
            <div class=""outer-box box-left"">
                <div id=""qr-box"" style=""text-align:center;""></div>
                <div class=""box-content"">
                    <label class=""title"">DATE :</label>
                    <label class=""content"">{sakuraiModel.Date}</label>
                </div>
                <div class=""box-content"">
                    <label class=""title"">I/V No. :</label>
                   <label id=""fit-text"" class=""content fit-text"" style=""width:150px;"">{sakuraiModel.InvoiNo}</label>
                </div>
                <div class=""box-content"">
                    <label class=""title"">CARTON No. :</label>
                    <label class=""content"">{sakuraiModel.CartonNo}</label>
                </div>
            </div>
            <script type='module'>
                import {{ BrowserQRCodeSvgWriter }} from 'https://cdn.jsdelivr.net/npm/@zxing/library@latest/+esm';
                const writer = new BrowserQRCodeSvgWriter();
                const qrBox = document.querySelector('.qr-box');
                const svg = writer.write('{qrData}', 250, 250);
                qrBox.appendChild(svg);
            </script>
            ";
            return htmlContent;
        }

        public string MonbellToWebview(MontbellModel montbellModel)
        {
            byte[] imageBytes = File.ReadAllBytes(montbellModel.ImagePath);
            string base64 = Convert.ToBase64String(imageBytes);
            string imageSrc = $"data:image/png;base64,{base64}";
            string quantityDisplay = string.Empty;
            if (montbellModel.TypeProduct == 1)
            {
                quantityDisplay = montbellModel.QuantityInBox.ToString("N0") + "/" + montbellModel.SizeQuantity.ToString("N0");
            }
            else
            {
                quantityDisplay = montbellModel.SizeQuantity.ToString("N0");
            }
            string htmlContent = $@"
            <div class=""crosshair-h"" id=""crosshair-h""></div>
            <div class=""crosshair-v"" id=""crosshair-v""></div>
            <div class=""outer-box box-left"">
                <table class=""form-table"">
                    <tr>
                        <td colspan=""1"" class=""label"">Order No </td>
                        <td colspan=""4"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:210px;"">
                                <span style=""padding-left:1.3em;"">{montbellModel.OrderNo}</span>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">TENTAC(PO NO)</td>
                        <td colspan=""1"" class=""content-bold"" style=""width:60px;"">{montbellModel.PoNo}</td>
                        <td colspan=""1"" class=""title"">TTL Style</td>
                        <td colspan=""1"" style=""text-align:right"">
                            <div id=""fit-text"" class="" fit-text"" style=""width:60px"">
                                <label style=""text-align:right"">{montbellModel.Quantity.ToString("N0")}</label>
                            </div>
                        </td>
                        <td colspan=""1"" style=""width:20px; text-align:right"">PCS</td>
                    </tr>
                    <tr style=""height: 4px;""></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">ITEM Code/Material</td>
                        <td colspan=""4"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:210px;"">
                                <label>{montbellModel.ItemCode} / {montbellModel.MaterialCode}</label>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">Style No</td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:60px"">
                                <label>{montbellModel.StyleNo}</label>
                            </div>
                        </td>
                        <td colspan=""1"" class=""title"">Size</td>
                        <td colspan=""2"" class=""content-bold"" style=""text-align:right"">{montbellModel.Size}</td>
                    </tr>
                    <tr style="" height: 4px;""></tr>
                    <tr>
                        <td colspan=""1"" class="" label"">Color</td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class=""fit-text"" style="" width:60px;"">
                                <label>{montbellModel.Color}</label>
                            </div>
                        </td>
                        <td colspan=""1"" class=""title"">QTY</td>
                        <td colspan=""1"" style=""text-align:right"">
                            <div id=""fit-text"" class="" fit-text"" style=""width:60px;"">
                                <label style=""text-align:right"" class=""content-bold""> {quantityDisplay}</label>
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

        public string MujiToWebview(MujiModel mujiModel)
        {
            byte[] imageBytes = File.ReadAllBytes(mujiModel.ImagePath);
            string base64 = Convert.ToBase64String(imageBytes);
            string imageSrc = $"data:image/png;base64,{base64}";
            string quantityDisplay = mujiModel.Quantity.ToString("N0");
            if (mujiModel.QuantityInBox == 0)
            {
                quantityDisplay = mujiModel.Quantity.ToString("N0");
            }
            else
            {
                quantityDisplay = mujiModel.QuantityInBox.ToString("N0") + "/" + mujiModel.Quantity.ToString("N0");
            }
            string htmlContent = $@"
                    <div class=""crosshair-h"" id=""crosshair-h""></div>
                    <div class=""crosshair-v"" id=""crosshair-v""></div>
                    <div class=""outer-box box-left"">
                     <table class=""form-table"">
                         <tr>
                             <td class=""label"">Customer</td>
                             <td colspan=""3"" style=""text-align:left;"">
                                <div id=""fit-text"" class=""fit-text"" style=""width:235px;"">
                                    <span>{mujiModel.DeliveryDestination}</span>
                                </div>
                             </td>
                         </tr>
                         <tr style=""height: 4px;""><td></td></tr>
                         <tr>
                             <td class=""label"">Order No</td>
                             <td colspan=""1"">
                                <div id=""fit-text"" class=""fit-text"" style=""width:65px;"">
                                   <span>{mujiModel.PoNo}</span>
                                </div>
                            </td>
                            <td colspan=""2"" style=""text-align:right"">
                                <div id=""fit-text"" class=""fit-text"" style=""width:160px;"">
                                   <span>{mujiModel.POCustomer}</span>
                                </div>
                            </td>
                         </tr>
                         <tr style=""height: 4px;""><td></td></tr>
                         <tr>
                             <td class=""label"">Item Code</td>
                             <td colspan=""3"" style=""padding-left:8px;"">
                                <div id=""fit-text"" class=""fit-text"" style=""width:225px;"">
                                   <span>{mujiModel.ItemCode}</span>
                                </div>
                            </td>
                         </tr>
                         <tr style=""height: 4px;""><td></td></tr>
                         <tr>
                             <td class=""label"">Quantity</td>
                             <td colspan=""2"" style=""text-align:right"">
                                <span class=""content-bold"">{quantityDisplay}</span>
                            </td>
                             <td colspan=""1"" class=""pcs"">PCS</td>
                         </tr>
                         <tr style=""height: 4px;""><td></td></tr>
                         <tr>
                             <td class=""label""></td>
                             <td colspan=""1"" style=""padding-left:8px;""> {mujiModel.KeyNo}</td>
                             <td colspan=""2"" class=""content-bold"" style=""text-align:right"">FSC Mix Credit</td>
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
                             <img src=""{imageSrc}"" class=""address-img"" />
                         </div>
                         <div class=""parent-container"">
                             <div class=""qr-box""></div>
                         </div>
                     </div>
                 </div>";
            return htmlContent;
        }

        public string ThermalUAToWebview(ThermalUnderamour thermalUnderamour)
        {
            byte[] imageBytes = File.ReadAllBytes(thermalUnderamour.ImagePath);
            string base64 = Convert.ToBase64String(imageBytes);
            string imageSrc = $"data:image/png;base64,{base64}";
            string htmlContent = $@"
            <div class=""crosshair-h"" id=""crosshair-h""></div>
            <div class=""crosshair-v"" id=""crosshair-v""></div>
            <div class=""outer-box box-left"">
                <table class=""form-table"">
                    <tr>
                        <td colspan=""1"" class=""label"">Customer</td>
                        <td colspan=""3"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:240px;"">
                                <label>{thermalUnderamour.Customer}</label>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">Order No</td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class="" fit-text"" style=""width:80px"">
                                <label>{thermalUnderamour.OrderNo}</label>
                            </div>
                        </td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class="" fit-text""; style=""width:70px;"">
                                <label>{thermalUnderamour.StyleNo}</label>
                            </div>
                        </td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class=""fit-text"" style=""text-align:right; width:90px;"">
                                <label class=""content-bold"">{thermalUnderamour.PONo}</label>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">ITEM Code</td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:80px;"">
                                <label>{thermalUnderamour.Product}</label>
                            </div>
                        </td>
                        <td colspan=""2"">
                            <div id=""fit-text"" class=""fit-text"" style=""text-align:right; width:160px;"">
                                <label>{thermalUnderamour.Description}</label>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">Quantity</td>
                        <td colspan=""3"">
                            <div id=""fit-text"" class=""fit-text"" style=""text-align:right;"">
                                <label class=""content-bold"" style=""text-align:right"">{thermalUnderamour.Quantity.ToString("N0")}</label>
                                <label class=""title"">pcs</label>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class="" label"">Color/Size</td>
                        <td colspan=""2"">
                            <div id=""fit-text"" class=""fit-text"" style="" width:150px;"">
                                <label class=""content-bold"">{thermalUnderamour.ColorCode}/{thermalUnderamour.SizeName}</label>
                            </div>
                        </td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class="" fit-text"" style=""text-align:right; width:90px;"">
                                <label style=""text-align:right"">{thermalUnderamour.JAN}</label>
                            </div>
                        </td>
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

        public string KonicaUAToWebview(KonicaUnderamour konicaUnderamour)
        {
            byte[] imageBytes = File.ReadAllBytes(konicaUnderamour.ImagePath);
            string base64 = Convert.ToBase64String(imageBytes);
            string imageSrc = $"data:image/png;base64,{base64}";
            string htmlContent = $@"
            <div class=""crosshair-h"" id=""crosshair-h""></div>
            <div class=""crosshair-v"" id=""crosshair-v""></div>
            <div class=""outer-box box-left"">
                <table class=""form-table"">
                    <tr>
                        <td colspan=""1"" class=""label"">Customer</td>
                        <td colspan=""2"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:240px;"">
                                <label>{konicaUnderamour.Customer}</label>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""><td></td></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">Order No</td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class="" fit-text"" style=""width:120px"">
                                <label>{konicaUnderamour.OrderNo}</label>
                            </div>
                        </td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class=""fit-text"" style=""text-align:right; width:120px;"">
                                <label class=""content-bold"">{konicaUnderamour.PONo}</label>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">ITEM Code</td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:120px;"">
                                <label>{konicaUnderamour.Product}</label>
                            </div>
                        </td>
                        <td colspan=""2"">
                            <div id=""fit-text"" class=""fit-text"" style=""text-align:right; width:120px;"">
                                <label>{konicaUnderamour.Description}</label>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 4px;""></tr>
                    <tr>
                        <td colspan=""1"" class=""label"">Quantity</td>
                        <td colspan=""2"">
                            <div id=""fit-text"" class=""fit-text"" style=""text-align:right;"">
                                <label class=""content-bold"">{konicaUnderamour.Quantity.ToString("N0")}</label>
                                <label class=""title"">pcs</label>
                            </div>
                        </td>
                    </tr>
                    <tr style="" height: 4px;""></tr>
                    <tr>
                        <td colspan=""1"" class="" label"">Color/Size</td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class=""fit-text"" style="" width:120px;"">
                                <label class=""content-bold"">{konicaUnderamour.ColorName}</label>
                            </div>
                        </td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class=""fit-text"" style="" width:120px; text-align:right;"">
                                <label class=""content-bold"">{konicaUnderamour.SizeName}</label>
                            </div>
                        </td>
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

        public string AsicsToWebview(AsicsModel asicsModel)
        {
            byte[] imageBytes = File.ReadAllBytes(asicsModel.ImagePath);
            string base64 = Convert.ToBase64String(imageBytes);
            string imageSrc = $"data:image/png;base64,{base64}";
            string htmlContent = $@"
            <div class=""crosshair-h"" id=""crosshair-h""></div>
            <div class=""crosshair-v"" id=""crosshair-v""></div>
            <div class=""outer-box box-left"">
                <table class=""form-table"">
                    <tr>
                        <td class=""label"">Customer</td>
                        <td colspan=""5"" style=""text-align:left;"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:225px;"">
                                <span>{asicsModel.Customer}</span>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 3px;""><td></td></tr>
                    <tr>
                        <td class=""label"">Order No</td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:60px;"">
                                <span>{asicsModel.OrderNo}</span>
                            </div>
                        </td>
                        <td colspan=""2"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:50px;"">
                                <span>{asicsModel.SKUNo}</span>
                            </div>
                        </td>
                        <td colspan=""3"" style=""text-align:right"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:100px;"">
                                <span>{asicsModel.ContractNo}</span>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 3px;""><td></td></tr>
                    <tr>
                        <td class=""label"">Item Code</td>
                        <td colspan=""5"" rowspan=""2"">
                            <div style=""width:225px; white-space: normal; word-wrap: break-word;"">
                                <div id=""fit-text"" class=""fit-text"" style=""width:225px;"">{asicsModel.ItemCode}</div>
                                <div id=""fit-text"" class=""fit-text"" style=""width:225px;"">{asicsModel.ProductName}</div>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 3px;""><td></td></tr>
                    <tr style=""height: 3px;""><td></td></tr>
                    <tr>
                        <td class=""label"">Quantity</td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:60px;"">
                                <span>{asicsModel.SampleNo}</span>
                            </div>
                        </td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:30px;"">
                                <span>{asicsModel.ColorCode}</span>
                            </div>
                        </td>
                        <td colspan=""1"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:30px;"">
                                <span>{asicsModel.SizeName}</span>
                            </div>
                        </td>
                        <td colspan=""1"" style=""text-align:right; width:75px;"">
                            <span class=""content-bold"">{asicsModel.Quantity}</span>
                        </td>
                        <td colspan=""1"" class=""pcs"">PCS</td>
                    </tr>
                    <tr style=""height: 3px;""><td></td></tr>
                    <tr>
                        <td class=""label"" style=""text-align:center"">{asicsModel.POLine}</td>
                        <td colspan=""5"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:225px;"">
                                <span>{asicsModel.KeyNo}</span>
                            </div>
                        </td>
                    </tr>
                    <tr style=""height: 3px;""><td></td></tr>
                    <tr>
                        <td class=""label"">Vendors Ref.</td>
                        <td colspan=""5"">
                            <div id=""fit-text"" class=""fit-text"" style=""width:225px;"">
                                <span>{asicsModel.Vendors}</span>
                            </div>
                        </td>
                    </tr>
                </table>
                <div class=""notice"">
                    <span>Please open and confirm this package immediately on receipt.</span>
                </div>
                <div class=""footer"">
                    <span>MADE IN VIETNAM</span>
                    <div class=""parent-container"">
                        <div class=""qr-box""></div>
                    </div>
                </div>
                <div class=""container"">
                    <div style=""width: 30%; display: inline-block;"">
                        <img src=""{imageSrc}"" class="" address-img"" />
                    </div>
                </div>
            </div>";
            return htmlContent;
        }

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