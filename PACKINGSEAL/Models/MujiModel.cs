namespace PACKINGSEAL.Models
{
    public class MujiModel
    {
        public string ImagePath { get; set; }
        public string DeliveryDestination { get; set; }
        public string ItemCode { get; set; }
        public string KeyNo { get; set; }
        public string PoNo { get; set; }
        public string POCustomer { get; set; }
        public double Quantity { get; set; }
        public double QuantityInBox { get; set; }
        public string LotNo {  get; set; }
        public int Index { get; set; } = 0;
    }
}
