using System.Drawing;
using System.Windows.Forms;

namespace PACKINGSEAL.Utils
{
    public class AutoCloseMessageBox : Form
    {
        public AutoCloseMessageBox(string message, int timeout)
        {
            // Tạo label hiển thị nội dung
            Label lbl = new Label()
            {
                Text = message,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 12, FontStyle.Regular),
                ForeColor = Color.LightGoldenrodYellow,
                BackColor = ColorTranslator.FromHtml("#10274B")
            };
            Controls.Add(lbl);

            // Tạo timer để tự đóng
            Timer timer = new Timer();
            timer.Interval = timeout; // thời gian tính bằng ms
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                this.Close();
            };
            timer.Start();

            // Cấu hình form
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size(400, 150);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.ControlBox = false;
        }

        public static void ShowMessage(string message, int timeout = 2000)
        {
            AutoCloseMessageBox box = new AutoCloseMessageBox(message, timeout);
            box.Show();
        }
    }
}
