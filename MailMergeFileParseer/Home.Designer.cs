namespace MailMergeFileParseer
{
    partial class Home
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
            this.metroLabel1 = new MetroFramework.Controls.MetroLabel();
            this.metroLabel2 = new MetroFramework.Controls.MetroLabel();
            this.TemplateFileBtn = new MetroFramework.Controls.MetroButton();
            this.ExcelFileBtn = new MetroFramework.Controls.MetroButton();
            this.ExcelFileLabel = new MetroFramework.Controls.MetroLabel();
            this.TmplateFileLabel = new MetroFramework.Controls.MetroLabel();
            this.Transform = new MetroFramework.Controls.MetroButton();
            this.Result = new MetroFramework.Controls.MetroLabel();
            this.metroProgressBar1 = new MetroFramework.Controls.MetroProgressBar();
            this.SuspendLayout();
            // 
            // metroLabel1
            // 
            this.metroLabel1.AutoSize = true;
            this.metroLabel1.Location = new System.Drawing.Point(28, 123);
            this.metroLabel1.Name = "metroLabel1";
            this.metroLabel1.Size = new System.Drawing.Size(99, 20);
            this.metroLabel1.TabIndex = 0;
            this.metroLabel1.Text = "Excel File Path:";
            // 
            // metroLabel2
            // 
            this.metroLabel2.AutoSize = true;
            this.metroLabel2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.metroLabel2.Location = new System.Drawing.Point(28, 169);
            this.metroLabel2.Name = "metroLabel2";
            this.metroLabel2.Size = new System.Drawing.Size(123, 20);
            this.metroLabel2.TabIndex = 0;
            this.metroLabel2.Text = "Template File Path:";
            // 
            // TemplateFileBtn
            // 
            this.TemplateFileBtn.Location = new System.Drawing.Point(238, 166);
            this.TemplateFileBtn.Name = "TemplateFileBtn";
            this.TemplateFileBtn.Size = new System.Drawing.Size(97, 23);
            this.TemplateFileBtn.TabIndex = 1;
            this.TemplateFileBtn.Text = "Browse";
            this.TemplateFileBtn.Click += new System.EventHandler(this.TemplateFileBtn_Click);
            // 
            // ExcelFileBtn
            // 
            this.ExcelFileBtn.Location = new System.Drawing.Point(238, 123);
            this.ExcelFileBtn.Name = "ExcelFileBtn";
            this.ExcelFileBtn.Size = new System.Drawing.Size(97, 23);
            this.ExcelFileBtn.TabIndex = 1;
            this.ExcelFileBtn.Text = "Browse";
            this.ExcelFileBtn.Click += new System.EventHandler(this.ExcelFileBtn_Click);
            // 
            // ExcelFileLabel
            // 
            this.ExcelFileLabel.AutoSize = true;
            this.ExcelFileLabel.Location = new System.Drawing.Point(543, 126);
            this.ExcelFileLabel.Name = "ExcelFileLabel";
            this.ExcelFileLabel.Size = new System.Drawing.Size(87, 20);
            this.ExcelFileLabel.TabIndex = 2;
            this.ExcelFileLabel.Text = "metroLabel3";
            // 
            // TmplateFileLabel
            // 
            this.TmplateFileLabel.AutoSize = true;
            this.TmplateFileLabel.Location = new System.Drawing.Point(543, 169);
            this.TmplateFileLabel.Name = "TmplateFileLabel";
            this.TmplateFileLabel.Size = new System.Drawing.Size(87, 20);
            this.TmplateFileLabel.TabIndex = 2;
            this.TmplateFileLabel.Text = "metroLabel3";
            // 
            // Transform
            // 
            this.Transform.Location = new System.Drawing.Point(1090, 508);
            this.Transform.Name = "Transform";
            this.Transform.Size = new System.Drawing.Size(145, 23);
            this.Transform.TabIndex = 3;
            this.Transform.Text = "Transform Files";
            this.Transform.UseMnemonic = false;
            this.Transform.Click += new System.EventHandler(this.Transform_Click);
            // 
            // Result
            // 
            this.Result.AutoSize = true;
            this.Result.CustomBackground = true;
            this.Result.CustomForeColor = true;
            this.Result.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Result.Location = new System.Drawing.Point(28, 69);
            this.Result.Name = "Result";
            this.Result.Size = new System.Drawing.Size(87, 20);
            this.Result.TabIndex = 2;
            this.Result.Text = "metroLabel3";
            // 
            // metroProgressBar1
            // 
            this.metroProgressBar1.Location = new System.Drawing.Point(449, 257);
            this.metroProgressBar1.Name = "metroProgressBar1";
            this.metroProgressBar1.Size = new System.Drawing.Size(292, 23);
            this.metroProgressBar1.TabIndex = 4;
            // 
            // Home
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1322, 582);
            this.Controls.Add(this.metroProgressBar1);
            this.Controls.Add(this.Transform);
            this.Controls.Add(this.TmplateFileLabel);
            this.Controls.Add(this.Result);
            this.Controls.Add(this.ExcelFileLabel);
            this.Controls.Add(this.ExcelFileBtn);
            this.Controls.Add(this.TemplateFileBtn);
            this.Controls.Add(this.metroLabel2);
            this.Controls.Add(this.metroLabel1);
            this.Name = "Home";
            this.Text = "Mail Merged File Creator";
            this.Load += new System.EventHandler(this.Home_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private MetroFramework.Controls.MetroLabel metroLabel1;
        private MetroFramework.Controls.MetroLabel metroLabel2;
        private MetroFramework.Controls.MetroButton TemplateFileBtn;
        private MetroFramework.Controls.MetroButton ExcelFileBtn;
        private MetroFramework.Controls.MetroLabel ExcelFileLabel;
        private MetroFramework.Controls.MetroLabel TmplateFileLabel;
        private MetroFramework.Controls.MetroButton Transform;
        private MetroFramework.Controls.MetroLabel Result;
        private MetroFramework.Controls.MetroProgressBar metroProgressBar1;
    }
}