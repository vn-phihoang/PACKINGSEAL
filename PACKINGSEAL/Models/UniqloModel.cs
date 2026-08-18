namespace PACKINGSEAL.Models
{
    public class UniqloModel
    {
        public string ImagePath { get; set; }
        public string OrderNo { get; set; }
        public string DeliveryDestination { get; set; }
        public string PoNo { get; set; }
        public string SampleNo { get; set; }
        public string MaterialCode { get; set; }
        public string SizeName { get; set; }
        public string ColorName { get; set; }
        public string ColorCode { get; set; }
        public double Quantity { get; set; }
        public string KeyNo { get; set; }
        public string ContractNo { get; set; }
        public double QuantityInBox { get; set; }

        public string LotNo { get; set; }
        public int Index { get; set; } = 0;
    }
}
