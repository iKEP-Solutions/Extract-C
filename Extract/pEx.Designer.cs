namespace Extract
{
    partial class pEx
    {
        /// <summary> 
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary> 
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas 
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.StatusStrip1 = new System.Windows.Forms.StatusStrip();
            this.tInit = new System.Windows.Forms.ToolStripStatusLabel();
            this.i_info = new System.Windows.Forms.ToolStripStatusLabel();
            this.lVue = new System.Windows.Forms.ListBox();
            this.Button1 = new System.Windows.Forms.Button();
            this.StatusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // StatusStrip1
            // 
            this.StatusStrip1.ImageScalingSize = new System.Drawing.Size(22, 22);
            this.StatusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tInit,
            this.i_info});
            this.StatusStrip1.Location = new System.Drawing.Point(0, 492);
            this.StatusStrip1.Name = "StatusStrip1";
            this.StatusStrip1.Size = new System.Drawing.Size(200, 22);
            this.StatusStrip1.TabIndex = 29;
            this.StatusStrip1.Text = "StatusStrip1";
            // 
            // tInit
            // 
            this.tInit.DoubleClickEnabled = true;
            this.tInit.Name = "tInit";
            this.tInit.Size = new System.Drawing.Size(24, 17);
            this.tInit.Text = "init";
            this.tInit.Click += new System.EventHandler(this.tInit_Click);
            // 
            // i_info
            // 
            this.i_info.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.i_info.DoubleClickEnabled = true;
            this.i_info.Image = global::Extract.Properties.Resources.info3_16_fw;
            this.i_info.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.i_info.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.i_info.IsLink = true;
            this.i_info.Name = "i_info";
            this.i_info.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.i_info.Size = new System.Drawing.Size(130, 17);
            this.i_info.Spring = true;
            this.i_info.Tag = "";
            this.i_info.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.i_info.Click += new System.EventHandler(this.i_info_Click);
            this.i_info.DoubleClick += new System.EventHandler(this.Param);
            // 
            // lVue
            // 
            this.lVue.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lVue.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.lVue.FormattingEnabled = true;
            this.lVue.Location = new System.Drawing.Point(3, 10);
            this.lVue.Name = "lVue";
            this.lVue.Size = new System.Drawing.Size(194, 442);
            this.lVue.Sorted = true;
            this.lVue.TabIndex = 28;
            this.lVue.SelectedIndexChanged += new System.EventHandler(this.lVue_SelectedIndexChanged);
            this.lVue.DoubleClick += new System.EventHandler(this.lVue_MouseDoubleClick);
            // 
            // Button1
            // 
            this.Button1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Button1.Location = new System.Drawing.Point(3, 457);
            this.Button1.Name = "Button1";
            this.Button1.Size = new System.Drawing.Size(194, 32);
            this.Button1.TabIndex = 26;
            this.Button1.Text = "Extraire";
            this.Button1.UseVisualStyleBackColor = true;
            this.Button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // pEx
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.StatusStrip1);
            this.Controls.Add(this.lVue);
            this.Controls.Add(this.Button1);
            this.Name = "pEx";
            this.Size = new System.Drawing.Size(200, 514);
            this.Load += new System.EventHandler(this.pEx_Load);
            this.StatusStrip1.ResumeLayout(false);
            this.StatusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.StatusStrip StatusStrip1;
        internal System.Windows.Forms.ToolStripStatusLabel tInit;
        internal System.Windows.Forms.ToolStripStatusLabel i_info;
        internal System.Windows.Forms.ListBox lVue;
        internal System.Windows.Forms.Button Button1;
    }
}
