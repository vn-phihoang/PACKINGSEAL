using System;
using System.Drawing;

namespace PACKINGSEAL.Models
{
    public class PrintResponseModel
    {
        public string Brand { get; set; }
        public string PoNumber { get; set; }
        public string PrinterName { get; set; }
        public DateTime PrintTime { get; set; }
        public string MsgVal { get; set; }
    }

}
