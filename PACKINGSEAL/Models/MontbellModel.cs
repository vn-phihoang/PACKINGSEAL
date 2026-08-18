namespace PACKINGSEAL.Models
{
    public class MontbellModel
    {
        public string ImagePath { get; set; }
        public string OrderNo { get; set; }
        public string PoNo { get; set; }
        public string StyleNo { get; set; }
        public string ItemCode { get; set; }
        public string MaterialCode { get; set; }
        public double Quantity { get; set; }
        public string Color { get; set; }
        public string Size { get; set; }
        public double SizeQuantity { get; set; }
        public double QuantityInBox { get; set; }
        public string LotNo { get; set; }
        public int Index { get; set; } = 0;
        public int TypeProduct { get; set; } = 1;
    }
}
