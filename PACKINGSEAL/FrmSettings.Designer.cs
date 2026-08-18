namespace PACKINGSEAL
{
    partial class FrmSettings
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmSettings));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnImportPath = new System.Windows.Forms.Button();
            this.textBoxImportPath = new System.Windows.Forms.TextBox();
            this.labelImportPath = new System.Windows.Forms.Label();
            this.chk_only_print = new System.Windows.Forms.CheckBox();
            this.chk_output_print = new System.Windows.Forms.CheckBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnRegister = new System.Windows.Forms.Button();
            this.btnPDFOutputPath = new System.Windows.Forms.Button();
            this.txtPDFOutputPath = new System.Windows.Forms.TextBox();
            this.lbl_excel_output_path = new System.Windows.Forms.Label();
            this.cbo_printer_list = new System.Windows.Forms.ComboBox();
            this.lbl_printer_list = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnImportPath);
            this.groupBox1.Controls.Add(this.textBoxImportPath);
            this.groupBox1.Controls.Add(this.labelImportPath);
            this.groupBox1.Controls.Add(this.chk_only_print);
            this.groupBox1.Controls.Add(this.chk_output_print);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnRegister);
            this.groupBox1.Controls.Add(this.btnPDFOutputPath);
            this.groupBox1.Controls.Add(this.txtPDFOutputPath);
            this.groupBox1.Controls.Add(this.lbl_excel_output_path);
            this.groupBox1.Controls.Add(this.cbo_printer_list);
            this.groupBox1.Controls.Add(this.lbl_printer_list);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(456, 204);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            // 
            // btnImportPath
            // 
            this.btnImportPath.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnImportPath.Location = new System.Drawing.Point(404, 125);
            this.btnImportPath.Name = "btnImportPath";
            this.btnImportPath.Size = new System.Drawing.Size(37, 29);
            this.btnImportPath.TabIndex = 7;
            this.btnImportPath.Text = ".....";
            this.btnImportPath.UseVisualStyleBackColor = true;
            this.btnImportPath.Click += new System.EventHandler(this.btnImportPath_Click);
            // 
            // textBoxImportPath
            // 
            this.textBoxImportPath.Location = new System.Drawing.Point(149, 126);
            this.textBoxImportPath.Name = "textBoxImportPath";
            this.textBoxImportPath.Size = new System.Drawing.Size(249, 26);
            this.textBoxImportPath.TabIndex = 6;
            // 
            // labelImportPath
            // 
            this.labelImportPath.AutoSize = true;
            this.labelImportPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelImportPath.Location = new System.Drawing.Point(12, 129);
            this.labelImportPath.Name = "labelImportPath";
            this.labelImportPath.Size = new System.Drawing.Size(92, 20);
            this.labelImportPath.TabIndex = 12;
            this.labelImportPath.Text = "Import Path";
            // 
            // chk_only_print
            // 
            this.chk_only_print.AutoSize = true;
            this.chk_only_print.Location = new System.Drawing.Point(345, 58);
            this.chk_only_print.Name = "chk_only_print";
            this.chk_only_print.Size = new System.Drawing.Size(95, 24);
            this.chk_only_print.TabIndex = 3;
            this.chk_only_print.Text = "Only Print";
            this.chk_only_print.UseVisualStyleBackColor = true;
            this.chk_only_print.CheckedChanged += new System.EventHandler(this.CheckedChanged);
            // 
            // chk_output_print
            // 
            this.chk_output_print.AutoSize = true;
            this.chk_output_print.Location = new System.Drawing.Point(148, 58);
            this.chk_output_print.Name = "chk_output_print";
            this.chk_output_print.Size = new System.Drawing.Size(177, 24);
            this.chk_output_print.TabIndex = 2;
            this.chk_output_print.Text = "Print and output PDF";
            this.chk_output_print.UseVisualStyleBackColor = true;
            this.chk_output_print.CheckedChanged += new System.EventHandler(this.CheckedChanged);
            // 
            // btnCancel
            // 
            this.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCancel.Location = new System.Drawing.Point(335, 164);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(105, 33);
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnRegister
            // 
            this.btnRegister.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRegister.Location = new System.Drawing.Point(224, 164);
            this.btnRegister.Name = "btnRegister";
            this.btnRegister.Size = new System.Drawing.Size(105, 33);
            this.btnRegister.TabIndex = 9;
            this.btnRegister.Text = "Save";
            this.btnRegister.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnRegister.UseVisualStyleBackColor = true;
            this.btnRegister.Click += new System.EventHandler(this.btnRegister_Click);
            // 
            // btnPDFOutputPath
            // 
            this.btnPDFOutputPath.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnPDFOutputPath.Location = new System.Drawing.Point(403, 87);
            this.btnPDFOutputPath.Name = "btnPDFOutputPath";
            this.btnPDFOutputPath.Size = new System.Drawing.Size(37, 29);
            this.btnPDFOutputPath.TabIndex = 5;
            this.btnPDFOutputPath.Text = ".....";
            this.btnPDFOutputPath.UseVisualStyleBackColor = true;
            this.btnPDFOutputPath.Click += new System.EventHandler(this.btnPDFOutputPath_Click);
            // 
            // txtPDFOutputPath
            // 
            this.txtPDFOutputPath.Location = new System.Drawing.Point(148, 88);
            this.txtPDFOutputPath.Name = "txtPDFOutputPath";
            this.txtPDFOutputPath.Size = new System.Drawing.Size(249, 26);
            this.txtPDFOutputPath.TabIndex = 4;
            // 
            // lbl_excel_output_path
            // 
            this.lbl_excel_output_path.AutoSize = true;
            this.lbl_excel_output_path.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_excel_output_path.Location = new System.Drawing.Point(11, 91);
            this.lbl_excel_output_path.Name = "lbl_excel_output_path";
            this.lbl_excel_output_path.Size = new System.Drawing.Size(131, 20);
            this.lbl_excel_output_path.TabIndex = 3;
            this.lbl_excel_output_path.Text = "PDF Output Path";
            // 
            // cbo_printer_list
            // 
            this.cbo_printer_list.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbo_printer_list.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbo_printer_list.FormattingEnabled = true;
            this.cbo_printer_list.Location = new System.Drawing.Point(148, 24);
            this.cbo_printer_list.Name = "cbo_printer_list";
            this.cbo_printer_list.Size = new System.Drawing.Size(292, 28);
            this.cbo_printer_list.TabIndex = 1;
            // 
            // lbl_printer_list
            // 
            this.lbl_printer_list.AutoSize = true;
            this.lbl_printer_list.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_printer_list.Location = new System.Drawing.Point(11, 28);
            this.lbl_printer_list.Name = "lbl_printer_list";
            this.lbl_printer_list.Size = new System.Drawing.Size(84, 20);
            this.lbl_printer_list.TabIndex = 0;
            this.lbl_printer_list.Text = "Printer List";
            // 
            // FrmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(456, 204);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Settings";
            this.Load += new System.EventHandler(this.FrmOptions_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox chk_only_print;
        private System.Windows.Forms.CheckBox chk_output_print;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnRegister;
        private System.Windows.Forms.Button btnPDFOutputPath;
        private System.Windows.Forms.TextBox txtPDFOutputPath;
        private System.Windows.Forms.Label lbl_excel_output_path;
        private System.Windows.Forms.ComboBox cbo_printer_list;
        private System.Windows.Forms.Label lbl_printer_list;
        private System.Windows.Forms.Button btnImportPath;
        private System.Windows.Forms.TextBox textBoxImportPath;
        private System.Windows.Forms.Label labelImportPath;
    }
}