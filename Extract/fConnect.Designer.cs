namespace Extract
{
    partial class fConnect
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
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.tID = new System.Windows.Forms.Label();
            this.tConStr = new System.Windows.Forms.TextBox();
            this.tNom = new System.Windows.Forms.TextBox();
            this.lCon = new System.Windows.Forms.ListBox();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(94, 230);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(77, 25);
            this.button4.TabIndex = 13;
            this.button4.Text = "Nouveau";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(15, 230);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(73, 25);
            this.button3.TabIndex = 14;
            this.button3.Text = "Supprimer";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(448, 230);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 25);
            this.button1.TabIndex = 12;
            this.button1.Text = "Enregistrer";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // tID
            // 
            this.tID.AutoSize = true;
            this.tID.Location = new System.Drawing.Point(189, 15);
            this.tID.Name = "tID";
            this.tID.Size = new System.Drawing.Size(18, 13);
            this.tID.TabIndex = 11;
            this.tID.Tag = "ConStr_Id,k";
            this.tID.Text = "ID";
            // 
            // tConStr
            // 
            this.tConStr.AcceptsReturn = true;
            this.tConStr.Location = new System.Drawing.Point(192, 38);
            this.tConStr.Multiline = true;
            this.tConStr.Name = "tConStr";
            this.tConStr.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tConStr.Size = new System.Drawing.Size(331, 186);
            this.tConStr.TabIndex = 10;
            this.tConStr.Tag = "ConnectionString,t";
            // 
            // tNom
            // 
            this.tNom.Location = new System.Drawing.Point(213, 12);
            this.tNom.Name = "tNom";
            this.tNom.Size = new System.Drawing.Size(310, 20);
            this.tNom.TabIndex = 9;
            this.tNom.Tag = "nom,t";
            // 
            // lCon
            // 
            this.lCon.FormattingEnabled = true;
            this.lCon.Location = new System.Drawing.Point(15, 12);
            this.lCon.Name = "lCon";
            this.lCon.Size = new System.Drawing.Size(156, 212);
            this.lCon.TabIndex = 8;
            this.lCon.DoubleClick += new System.EventHandler(this.dbl);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(352, 230);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 25);
            this.button2.TabIndex = 15;
            this.button2.Text = "Tester";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // fConnect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(543, 267);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.tID);
            this.Controls.Add(this.tConStr);
            this.Controls.Add(this.tNom);
            this.Controls.Add(this.lCon);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "fConnect";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Load += new System.EventHandler(this.fConnect_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label tID;
        private System.Windows.Forms.TextBox tConStr;
        private System.Windows.Forms.TextBox tNom;
        private System.Windows.Forms.ListBox lCon;
        private System.Windows.Forms.Button button2;
    }
}