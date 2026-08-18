using System.Drawing;
using System.Drawing.Drawing2D;

namespace PACKINGSEAL.Utils
{
    public class RoundedRectangleHelper
    {
        // Tạo GraphicsPath cho rectangle bo góc
        public static GraphicsPath Create(float x, float y, float w, float h, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int d = radius * 2;

            // Góc trên trái
            path.AddArc(x, y, d, d, 180, 90);
            // Cạnh trên
            path.AddLine(x + radius, y, x + w - radius, y);
            // Góc trên phải
            path.AddArc(x + w - d, y, d, d, 270, 90);
            // Cạnh phải
            path.AddLine(x + w, y + radius, x + w, y + h - radius);
            // Góc dưới phải
            path.AddArc(x + w - d, y + h - d, d, d, 0, 90);
            // Cạnh dưới
            path.AddLine(x + w - radius, y + h, x + radius, y + h);
            // Góc dưới trái
            path.AddArc(x, y + h - d, d, d, 90, 90);
            // Cạnh trái
            path.AddLine(x, y + h - radius, x, y + radius);

            path.CloseFigure();
            return path;
        }

        // Vẽ khung bo góc
        public static void Draw(Graphics g, Pen pen, Rectangle rect, int radius)
        {
            using (GraphicsPath path = Create(rect.X, rect.Y, rect.Width, rect.Height, radius))
            {
                g.DrawPath(pen, path);
            }
        }

        // Tô nền bo góc (nếu cần fill)
        public static void Fill(Graphics g, Brush brush, Rectangle rect, int radius)
        {
            using (GraphicsPath path = Create(rect.X, rect.Y, rect.Width, rect.Height, radius))
            {
                g.FillPath(brush, path);
            }
        }
    }
}
