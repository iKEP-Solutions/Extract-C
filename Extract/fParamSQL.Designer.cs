namespace Extract
{
    partial class fParamSQL
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
            this.lVue = new System.Windows.Forms.ListBox();
            this.tNom = new System.Windows.Forms.TextBox();
            this.tSQL = new System.Windows.Forms.TextBox();
            this.tID = new System.Windows.Forms.Label();
            this.lCon = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.gUser = new System.Windows.Forms.DataGridView();
            this.ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.User = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CritSQL = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MOD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.gUser)).BeginInit();
            this.SuspendLayout();
            // 
            // lVue
            // 
            this.lVue.FormattingEnabled = true;
            this.lVue.Location = new System.Drawing.Point(12, 7);
            this.lVue.Name = "lVue";
            this.lVue.Size = new System.Drawing.Size(212, 537);
            this.lVue.TabIndex = 0;
            this.lVue.SelectedIndexChanged += new System.EventHandler(this.lVue_SelectedIndexChanged);
            this.lVue.DoubleClick += new System.EventHandler(this.lVue_DoubleClick);
            // 
            // tNom
            // 
            this.tNom.Location = new System.Drawing.Point(264, 8);
            this.tNom.Name = "tNom";
            this.tNom.Size = new System.Drawing.Size(310, 20);
            this.tNom.TabIndex = 1;
            this.tNom.Tag = "nom,t";
            // 
            // tSQL
            // 
            this.tSQL.AcceptsReturn = true;
            this.tSQL.Location = new System.Drawing.Point(243, 39);
            this.tSQL.Multiline = true;
            this.tSQL.Name = "tSQL";
            this.tSQL.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tSQL.Size = new System.Drawing.Size(533, 194);
            this.tSQL.TabIndex = 2;
            this.tSQL.Tag = "CmdSQL,t";
            // 
            // tID
            // 
            this.tID.AutoSize = true;
            this.tID.Location = new System.Drawing.Point(240, 11);
            this.tID.Name = "tID";
            this.tID.Size = new System.Drawing.Size(18, 13);
            this.tID.TabIndex = 3;
            this.tID.Tag = "vue_Id,k";
            this.tID.Text = "ID";
            // 
            // lCon
            // 
            this.lCon.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.lCon.FormattingEnabled = true;
            this.lCon.Location = new System.Drawing.Point(580, 7);
            this.lCon.Name = "lCon";
            this.lCon.Size = new System.Drawing.Size(148, 21);
            this.lCon.TabIndex = 4;
            this.lCon.Tag = "constr_id";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(701, 550);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 28);
            this.button1.TabIndex = 5;
            this.button1.Text = "Enregistrer";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(744, 6);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(32, 21);
            this.button2.TabIndex = 6;
            this.button2.Text = "...";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(12, 550);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 28);
            this.button3.TabIndex = 7;
            this.button3.Text = "Supprimer";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(149, 550);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 28);
            this.button4.TabIndex = 7;
            this.button4.Text = "Nouveau";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // gUser
            // 
            this.gUser.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.gUser.BackgroundColor = System.Drawing.Color.White;
            this.gUser.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.gUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gUser.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ID,
            this.User,
            this.CritSQL,
            this.MOD});
            this.gUser.Location = new System.Drawing.Point(243, 239);
            this.gUser.Name = "gUser";
            this.gUser.Size = new System.Drawing.Size(533, 305);
            this.gUser.TabIndex = 8;
            this.gUser.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.gUser_CellBeginEdit);
            this.gUser.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.gUser_CellEndEdit);
            // 
            // ID
            // 
            this.ID.FillWeight = 20F;
            this.ID.HeaderText = "ID";
            this.ID.Name = "ID";
            this.ID.ReadOnly = true;
            this.ID.Visible = false;
            // 
            // User
            // 
            this.User.FillWeight = 50F;
            this.User.HeaderText = "User";
            this.User.Name = "User";
            // 
            // CritSQL
            // 
            this.CritSQL.HeaderText = "CritSQL";
            this.CritSQL.Name = "CritSQL";
            // 
            // MOD
            // 
            this.MOD.FillWeight = 20F;
            this.MOD.HeaderText = "MOD";
            this.MOD.Name = "MOD";
            this.MOD.ReadOnly = true;
            // 
            // fParamSQL
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(792, 583);
            this.Controls.Add(this.gUser);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lCon);
            this.Controls.Add(this.tID);
            this.Controls.Add(this.tSQL);
            this.Controls.Add(this.tNom);
            this.Controls.Add(this.lVue);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Name = "fParamSQL";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Paramétrage";
            this.Load += new System.EventHandler(this.fParamSQL_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gUser)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox lVue;
        private System.Windows.Forms.TextBox tNom;
        private System.Windows.Forms.TextBox tSQL;
        private System.Windows.Forms.Label tID;
        private System.Windows.Forms.ComboBox lCon;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.DataGridView gUser;
        private System.Windows.Forms.DataGridViewTextBoxColumn ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn User;
        private System.Windows.Forms.DataGridViewTextBoxColumn CritSQL;
        private System.Windows.Forms.DataGridViewTextBoxColumn MOD;
    }
}