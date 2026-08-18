using System.Drawing;

namespace PACKINGSEAL.Models
{
    public class ImportResponseModel
    {
        public bool IsSuccess { get; set; }

        public string Message { get; set; }

        public int RecordsImported { get; set; }

        public int RecordsFailed { get; set; }

        public string SourceName { get; set; }

        public string ErrorDetail { get; set; }

    }
}
