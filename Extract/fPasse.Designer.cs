namespace Extract
{
    partial class fPasse
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
            this.bOK = new System.Windows.Forms.Button();
            this.bAnnul = new System.Windows.Forms.Button();
            this.tPasse = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // bOK
            // 
            this.bOK.Location = new System.Drawing.Point(111, 67);
            this.bOK.Name = "bOK";
            this.bOK.Size = new System.Drawing.Size(77, 25);
            this.bOK.TabIndex = 1;
            this.bOK.Text = "OK";
            this.bOK.UseVisualStyleBackColor = true;
            this.bOK.Click += new System.EventHandler(this.bOK_Click);
            // 
            // bAnnul
            // 
            this.bAnnul.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.bAnnul.Location = new System.Drawing.Point(12, 67);
            this.bAnnul.Name = "bAnnul";
            this.bAnnul.Size = new System.Drawing.Size(77, 25);
            this.bAnnul.TabIndex = 2;
            this.bAnnul.Text = "Annuler";
            this.bAnnul.UseVisualStyleBackColor = true;
            this.bAnnul.Click += new System.EventHandler(this.bAnnul_Click);
            // 
            // tPasse
            // 
            this.tPasse.Location = new System.Drawing.Point(60, 27);
            this.tPasse.Name = "tPasse";
            this.tPasse.PasswordChar = '*';
            this.tPasse.Size = new System.Drawing.Size(91, 20);
            this.tPasse.TabIndex = 0;
            this.tPasse.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // fPasse
            // 
            this.AcceptButton = this.bOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.bAnnul;
            this.ClientSize = new System.Drawing.Size(207, 111);
            this.ControlBox = false;
            this.Controls.Add(this.tPasse);
            this.Controls.Add(this.bAnnul);
            this.Controls.Add(this.bOK);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "fPasse";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Paramétrage";
            this.Load += new System.EventHandler(this.fPasse_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bOK;
        private System.Windows.Forms.Button bAnnul;
        private System.Windows.Forms.TextBox tPasse;
    }
}