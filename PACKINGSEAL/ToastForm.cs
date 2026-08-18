using System;
using System.Drawing;
using System.Windows.Forms;

namespace PACKINGSEAL
{
    public partial class ToastForm : Form
    {
        int toastX, toastY;
        public ToastForm(string type, string message)
        {
            this.Opacity = 0.9;
            InitializeComponent();
            labelType.Text = type;
            labelMessage.Text = message;
            switch (type)
            {
                case "SUCCESS":
                    panel1.BackColor = Color.FromArgb(57, 155, 53);
                    pictureBox1.Image = Properties.Resources.happy_32x32;
                    break;
                case "ERROR":
                    panel1.BackColor = Color.FromArgb(227, 50, 45);
                    pictureBox1.Image = Properties.Resources.shock_32x32; 
                    break;
                case "INFO":
                    panel1.BackColor = Color.FromArgb(18, 136, 191);
                    pictureBox1.Image = Properties.Resources.smile_32x32;
                    break;
                case "WARRING":
                    panel1.BackColor = Color.FromArgb(245, 171, 35);
                    pictureBox1.Image = Properties.Resources.annoyed_32x32; 
                    break;
            }
        }

        private void timerToast_Tick(object sender, EventArgs e)
        {
            toastY -= 2;
            this.Location = new Point(toastX, toastY);
            if (toastY <= 970)
            {
                timerToastShow.Stop();
                // Start the hide timer after the show timer stops
                timerToastHide.Start();
            }
        }
        int y = 30;
        private void timerToastHide_Tick(object sender, EventArgs e)
        {
            y--;
            if (y <= 0)
            {
                toastY += 3; // trượt xuống nhanh hơn
                this.Location = new Point(toastX, toastY);

                if (toastY >= 1100)
                {
                    timerToastHide.Stop();
                    this.Close();
                }
            }
        }

        private void ToastForm_Load(object sender, EventArgs e)
        {
            Position();
        }

        private void Position()
        {
            int ScreenWidth = Screen.PrimaryScreen.WorkingArea.Width;
            int ScreenHeight = Screen.PrimaryScreen.WorkingArea.Height;

            toastX = ScreenWidth - this.Width - 2;
            toastY = ScreenHeight - this.Height + 5;
            this.Location = new Point(toastX, toastY);
        }
    }
}
