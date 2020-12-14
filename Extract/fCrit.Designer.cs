namespace Extract
{
    partial class fCrit
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
            this.bAnnul = new System.Windows.Forms.Button();
            this.bOK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // bAnnul
            // 
            this.bAnnul.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bAnnul.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.bAnnul.Location = new System.Drawing.Point(15, 22);
            this.bAnnul.Name = "bAnnul";
            this.bAnnul.Size = new System.Drawing.Size(75, 23);
            this.bAnnul.TabIndex = 0;
            this.bAnnul.Text = "Annuler";
            this.bAnnul.UseVisualStyleBackColor = true;
            this.bAnnul.Click += new System.EventHandler(this.bAnnul_Click);
            // 
            // bOK
            // 
            this.bOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bOK.Location = new System.Drawing.Point(170, 22);
            this.bOK.Name = "bOK";
            this.bOK.Size = new System.Drawing.Size(75, 23);
            this.bOK.TabIndex = 1;
            this.bOK.Text = "OK";
            this.bOK.UseVisualStyleBackColor = true;
            this.bOK.Click += new System.EventHandler(this.bOK_Click);
            // 
            // fCrit
            // 
            this.AcceptButton = this.bOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.bAnnul;
            this.ClientSize = new System.Drawing.Size(263, 57);
            this.Controls.Add(this.bOK);
            this.Controls.Add(this.bAnnul);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "fCrit";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Critères";
            this.Load += new System.EventHandler(this.fCrit_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button bAnnul;
        private System.Windows.Forms.Button bOK;
    }
}