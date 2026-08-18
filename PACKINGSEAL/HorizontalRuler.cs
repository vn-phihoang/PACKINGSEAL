using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace PACKINGSEAL
{
    [ToolboxItem(true)]
    public class HorizontalRuler : Control
    {
        [Category("Ruler Settings")]
        //public float PixelsPerCm { get; set; } = 40.25f;
        public float PixelsPerCm { get; set; } = 39.55f;

        [Category("Ruler Settings")]
        public float MinorTickCm { get; set; } = 0.1f; // Vạch nhỏ mỗi 0.1 cm

        [Category("Ruler Settings")]
        public float MajorTickCm { get; set; } = 1.0f; // Vạch lớn mỗi 1 cm
        public double ZoomFactor { get; set; } = 1.0;
        public HorizontalRuler()
        {
            this.DoubleBuffered = true;
            this.BackColor = Color.White;
            this.Size = new Size(400, 25);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            float scaledPixelsPerCm = PixelsPerCm * (float)ZoomFactor;
            using (Pen tickPen = new Pen(Color.Black))
            using (Pen redTickPen = new Pen(Color.Red))
            using (Font font = new Font("Segoe UI", 9))
            using (Brush textBrush = new SolidBrush(Color.Black))
            {
                float minorStep = MinorTickCm * scaledPixelsPerCm;   // ví dụ: 0.1 cm
                float majorStep = MajorTickCm * scaledPixelsPerCm;   // ví dụ: 1.0 cm
                float redStep = 0.5f * scaledPixelsPerCm;            // vạch đỏ mỗi 0.5 cm

                int rulerHeight = this.Height;
                int length = this.Width;
                int maxCm = (int)(length / scaledPixelsPerCm);

                float majorTickTop = rulerHeight - 18;
                float redTickTop = rulerHeight - 12;   // vạch 0.5 cm cao hơn vạch nhỏ
                float minorTickTop = rulerHeight - 8;

                // Vạch lớn và số đo
                for (int cm = 0; cm <= maxCm; cm++)
                {
                    float pos = cm * scaledPixelsPerCm;

                    g.DrawLine(redTickPen, new PointF(pos, majorTickTop), new PointF(pos, rulerHeight));

                    string label = $"{cm}";
                    SizeF textSize = g.MeasureString(label, font);
                    float textX = pos - textSize.Width / 2 + 7;
                    float textY = rulerHeight - textSize.Height - 10;

                    g.DrawString(label, font, textBrush, new PointF(textX, textY));
                }

                // Vạch nhỏ
                for (float i = 0; i < length; i += minorStep)
                {
                    if (Math.Abs(i % majorStep) < 0.01f || Math.Abs(i % redStep) < 0.01f) continue;
                    g.DrawLine(tickPen, new PointF(i, minorTickTop), new PointF(i, rulerHeight));
                }

                // Vạch đỏ mỗi 0.5 cm
                for (float i = 0; i < length; i += redStep)
                {
                    if (Math.Abs(i % majorStep) < 0.01f) continue;
                    g.DrawLine(redTickPen, new PointF(i, redTickTop), new PointF(i, rulerHeight));
                }
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.ResumeLayout(false);

        }
    }
}
