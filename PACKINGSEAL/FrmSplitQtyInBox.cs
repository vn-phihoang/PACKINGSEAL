using System;
using System.Windows.Forms;

namespace PACKINGSEAL
{
    public partial class FrmSplitQtyInBox : Form
    {
        public int QuantityInBox { get; private set; }
        public FrmSplitQtyInBox(int currentQuantity)
        {
            InitializeComponent();
            numericUpDownQtyInBox.Value = currentQuantity;
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            QuantityInBox = (int)numericUpDownQtyInBox.Value;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
