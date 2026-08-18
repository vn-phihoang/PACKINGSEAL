using NPOI.SS.UserModel;

namespace PACKINGSEAL.Utils
{
    public class ExcelUtil
    {
        /// <summary>
        /// Lấy giá trị của ô, có xử lý vùng merge nếu có.
        /// </summary>
        /// <param name="sheet">Sheet cần đọc</param>
        /// <param name="rowIndex">Chỉ số dòng (bắt đầu từ 0)</param>
        /// <param name="colIndex">Chỉ số cột (bắt đầu từ 0)</param>
        /// <returns>Giá trị của ô dưới dạng chuỗi</returns>
        public string GetCellValueWithMerge(ISheet sheet, int rowIndex, int colIndex)
        {
            // Kiểm tra vùng merge
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                var region = sheet.GetMergedRegion(i);
                if (region.IsInRange(rowIndex, colIndex))
                {
                    var firstRow = region.FirstRow;
                    var firstCol = region.FirstColumn;
                    var mergedCell = sheet.GetRow(firstRow)?.GetCell(firstCol);
                    return mergedCell?.ToString().Trim() ?? string.Empty;
                }
            }

            // Nếu không nằm trong vùng merge, lấy trực tiếp
            var cell = sheet.GetRow(rowIndex)?.GetCell(colIndex);
            return cell?.ToString() ?? string.Empty;
        }

        /// <summary>
        /// Đếm số dòng trong sheet có ít nhất một ô chứa dữ liệu thực sự.
        /// Bỏ qua các dòng chỉ có định dạng (ví dụ: kẻ khung, màu nền) nhưng không có nội dung.
        /// </summary>
        /// <param name="sheet">Sheet cần kiểm tra</param>
        /// <returns>Số dòng có dữ liệu thực sự</returns>
        public int GetActualRowCount(ISheet sheet)
        {
            int count = 0;

            for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                {
                    count++;
                    continue;
                }

                bool hasData = false;

                for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                {
                    var cell = row.GetCell(j);
                    if (cell == null) continue;

                    if (cell.CellType == CellType.String && !string.IsNullOrWhiteSpace(cell.StringCellValue) ||
                        cell.CellType == CellType.Numeric ||
                        cell.CellType == CellType.Boolean ||
                        cell.CellType == CellType.Formula && !string.IsNullOrWhiteSpace(cell.ToString()))
                    {
                        hasData = true;
                        break;
                    }
                }

                if (hasData) count++;
            }

            return count;
        }
    }
}
