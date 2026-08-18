using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace PACKINGSEAL
{
    [ToolboxItem(true)]
    public class VerticalRuler : Control
    {
        [Category("Ruler Settings")]
        //public float PixelsPerCm { get; set; } = 40.25f;
        public float PixelsPerCm { get; set; } = 39.55f;

        [Category("Ruler Settings")]
        public float MinorTickCm { get; set; } = 0.1f; // Vạch nhỏ mỗi 0.1 cm

        [Category("Ruler Settings")]
        public float MajorTickCm { get; set; } = 1.0f; // Vạch lớn mỗi 1 cm
        public double ZoomFactor { get; set; } = 0.95;
        public VerticalRuler()
        {
            this.DoubleBuffered = true;
            this.BackColor = Color.White;
            this.Size = new Size(25, 400);
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

                int rulerWidth = this.Width;
                int length = this.Height;
                int maxCm = (int)(length / scaledPixelsPerCm);

                float majorTickLeft = rulerWidth - 18;
                float redTickLeft = rulerWidth - 12;   // vạch đỏ dài hơn vạch nhỏ
                float minorTickLeft = rulerWidth - 8;

                // Vạch lớn và số đo
                for (int cm = 0; cm <= maxCm; cm++)
                {
                    float pos = cm * scaledPixelsPerCm;

                    g.DrawLine(redTickPen, new PointF(majorTickLeft, pos), new PointF(rulerWidth, pos));

                    string label = $"{cm}";
                    SizeF textSize = g.MeasureString(label, font);
                    float textX = rulerWidth - textSize.Width - 10;
                    float textY = pos - textSize.Height / 2 + 8;

                    g.DrawString(label, font, textBrush, new PointF(textX, textY));
                }

                // Vạch nhỏ
                for (float i = 0; i < length; i += minorStep)
                {
                    if (Math.Abs(i % majorStep) < 0.01f || Math.Abs(i % redStep) < 0.01f) continue;
                    g.DrawLine(tickPen, new PointF(minorTickLeft, i), new PointF(rulerWidth, i));
                }

                // Vạch đỏ mỗi 0.5 cm
                for (float i = 0; i < length; i += redStep)
                {
                    if (Math.Abs(i % majorStep) < 0.01f) continue;
                    g.DrawLine(redTickPen, new PointF(redTickLeft, i), new PointF(rulerWidth, i));
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
