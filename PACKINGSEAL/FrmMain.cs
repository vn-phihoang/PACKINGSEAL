using PACKINGSEAL.Common.Consts;
using PACKINGSEAL.Models;
using System;
using System.Deployment.Application;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace PACKINGSEAL
{
    public partial class FrmMain : Form
    {
        FilePathModel filePathModel = new FilePathModel();
        public Assembly assembly = Assembly.GetExecutingAssembly();
        private System.Drawing.Image gifImage;

        public FrmMain()
        {
            InitializeComponent();
            string UserName = Environment.MachineName;
            lb_info_1.Text = "HC | Packing Onegai Seal- " + UserName;
            string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
            var version = "";
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                version = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            this.Text = "PACKING ONEGAI SEAL (梱包シール) Ver: " + version;
            LoadLocation();

            gifImage = System.Drawing.Image.FromFile($@"{Path.GetDirectoryName(assembly.Location)}\Resources\gif1.gif");
            // Bắt đầu animate
            ImageAnimator.Animate(gifImage, OnFrameChanged);
        }
        private void OnFrameChanged(object sender, EventArgs e)
        {
            // Mỗi khi frame thay đổi, vẽ lại panel
            panel4.Invalidate();
        }

        private void LoadLocation()
        {
            if (SettingINI.Location == "HCM")
            {
                radioButtonHCM.Checked = true;
                filePathModel.ImagePath = $@"{Path.GetDirectoryName(assembly.Location)}\Template\HCMAddressTentacEng.png";
                label9.Text = "WELCOME TO TENTAC (HO CHI MINH)";
            }
            else if (SettingINI.Location == "HNI")
            {
                radioButtonHN.Checked = true;
                filePathModel.ImagePath = $@"{Path.GetDirectoryName(assembly.Location)}\Template\HNAddressTentacEng.jpg";
                label9.Text = "WELCOME TO TENTAC (HA NOI)";
            }
        }

        private void radioButtonLocation_CheckedChanged(object sender, EventArgs e)
        {
            string Location = string.Empty;
            if (radioButtonHCM.Checked)
            {
                Location = "HCM";
                filePathModel.ImagePath = $@"{Path.GetDirectoryName(assembly.Location)}\Template\HCMAddressTentacEng.png";
                label9.Text = "WELLCOME TO TENTAC (HO CHI MINH)";
            }
            else if (radioButtonHN.Checked)
            {
                Location = "HNI";
                filePathModel.ImagePath = $@"{Path.GetDirectoryName(assembly.Location)}\Template\HNAddressTentacEng.jpg";
                label9.Text = "WELLCOME TO TENTAC (HA NOI)";
            }
            SettingINI.SetLocation(Location);
        }

        private void buttonUniqlo_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePathModel.ImagePath))
            {
                MessageBox.Show("Please select a location (HCM or HNI) before using the application.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            filePathModel.HtmlPath = $@"{Path.GetDirectoryName(assembly.Location)}\Template\TemplateUniqlo.html";

            FrmUniqlo frm = new FrmUniqlo(filePathModel);
            if (!Application.OpenForms.OfType<FrmSettings>().Any())
            {
                if (frm == null)
                {
                    frm = new FrmUniqlo(filePathModel);
                }
                frm.Show();
                this.Hide();
            }
            else
            {
                frm.Focus();
            }
        }

        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void buttonMontbell_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePathModel.ImagePath))
            {
                MessageBox.Show("Please select a location (HCM or HNI) before using the application.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            filePathModel.HtmlPath = $@"{Path.GetDirectoryName(assembly.Location)}\Template\TemplateMontbell.html";

            FrmMontbell frm = new FrmMontbell(filePathModel);
            if (!Application.OpenForms.OfType<FrmSettings>().Any())
            {
                if (frm == null)
                {
                    frm = new FrmMontbell(filePathModel);
                }
                frm.Show();
                this.Hide();
            }
            else
            {
                frm.Focus();
            }
        }

        private void buttonMuji_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePathModel.ImagePath))
            {
                MessageBox.Show("Please select a location (HCM or HNI) before using the application.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            filePathModel.HtmlPath = $@"{Path.GetDirectoryName(assembly.Location)}\Template\TemplateMuji.html";

            FrmMuji frm = new FrmMuji(filePathModel);
            if (!Application.OpenForms.OfType<FrmSettings>().Any())
            {
                if (frm == null)
                {
                    frm = new FrmMuji(filePathModel);
                }
                frm.Show();
                this.Hide();
            }
            else
            {
                frm.Focus();
            }
        }

        private void buttonSakurai_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePathModel.ImagePath))
            {
                MessageBox.Show("Please select a location (HCM or HNI) before using the application.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            filePathModel.HtmlPath = $@"{Path.GetDirectoryName(assembly.Location)}\Template\TemplateSakurai.html";

            FrmPackingSealSakurai frm = new FrmPackingSealSakurai(filePathModel);
            if (!Application.OpenForms.OfType<FrmSettings>().Any())
            {
                if (frm == null)
                {
                    frm = new FrmPackingSealSakurai(filePathModel);
                }
                frm.Show();
                this.Hide();
            }
            else
            {
                frm.Focus();
            }
        }

        private void buttonThermalUA_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePathModel.ImagePath))
            {
                MessageBox.Show("Please select a location (HCM or HNI) before using the application.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            filePathModel.HtmlPath = $@"{Path.GetDirectoryName(assembly.Location)}\Template\TemplateThermalUnderamour.html";

            FrmThermalUA frm = new FrmThermalUA(filePathModel);
            if (!Application.OpenForms.OfType<FrmSettings>().Any())
            {
                if (frm == null)
                {
                    frm = new FrmThermalUA(filePathModel);
                }
                frm.Show();
                this.Hide();
            }
            else
            {
                frm.Focus();
            }
        }

        private void buttonKonicaUA_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePathModel.ImagePath))
            {
                MessageBox.Show("Please select a location (HCM or HNI) before using the application.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            filePathModel.HtmlPath = $@"{Path.GetDirectoryName(assembly.Location)}\Template\TemplateKonicaUnderamour.html";

            FrmKonicaUA frm = new FrmKonicaUA(filePathModel);
            if (!Application.OpenForms.OfType<FrmSettings>().Any())
            {
                if (frm == null)
                {
                    frm = new FrmKonicaUA(filePathModel);
                }
                frm.Show();
                this.Hide();
            }
            else
            {
                frm.Focus();
            }
        }

        private void buttonAsics_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePathModel.ImagePath))
            {
                MessageBox.Show("Please select a location (HCM or HNI) before using the application.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            filePathModel.HtmlPath = $@"{Path.GetDirectoryName(assembly.Location)}\Template\TemplateAsics.html";

            FrmAsics frm = new FrmAsics(filePathModel);
            if (!Application.OpenForms.OfType<FrmSettings>().Any())
            {
                if (frm == null)
                {
                    frm = new FrmAsics(filePathModel);
                }
                frm.Show();
                this.Hide();
            }
            else
            {
                frm.Focus();
            }
        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {
            if (gifImage != null)
            {
                ImageAnimator.UpdateFrames(gifImage);
                // Vẽ fill vừa panel
                e.Graphics.DrawImage(gifImage, new Rectangle(0, 0, panel4.Width, panel4.Height));
            }
        }
    }
}
