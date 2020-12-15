using System;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Extract
{
    public partial class fParamSQL : Form
    {
        private string LaCelVal;

        private OleDbConnection LaConnect = new OleDbConnection();

        private void Initialise()
        {
            try
            {
                Extract.Properties.Settings.Default.Reload();
                Common.ConnexionInit(Extract.Properties.Settings.Default.VueConStr2, ref LaConnect);          
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message);
            }
        }

        public fParamSQL()
        {
            InitializeComponent();
        }

        private void fParamSQL_Load(object sender, EventArgs e)
        {
            Initialise();
            Common.ListFill("Select vue.vue_id,vue.nom from extract.vue order by vue.nom", this.lVue, ref LaConnect);
            Common.ComboFill("Select constr_id, Nom from extract.ConnectionString", this.lCon, ref LaConnect);
        }

        private void lVue_DoubleClick(object sender, EventArgs e)
        {
            Common.FormRempli(this, "SELECT Vue_id,Vue.nom,CmdSQL,ConStr_id FROM Extract.Vue where vue_id=" + ((Common.ListItem)lVue.SelectedItem).Val, ref LaConnect);

            OleDbDataReader LeRs;
//            LeRs = Common.SqlLit("SELECT Vue_id,Vue.Nom,CmdSQL,ConStr_id FROM Extract.Vue where vue_id=" + ((Common.ListItem)lVue.SelectedItem).Val, ref LaConnect);
//            while (LeRs.Read()) {
//                this.tID.Text = LeRs.GetInt32(0).ToString();
//                this.tNom.Text = LeRs.GetString(1);
//                this.tSQL.Text = LeRs.GetString(2);
//            }
//            LeRs.Close();

            this.gUser.Rows.Clear();
            LeRs = Common.SqlLit("SELECT VueUser_Id,UserLogin,CritSql FROM Extract.VueUser where vue_id=" + ((Common.ListItem)lVue.SelectedItem).Val, ref LaConnect);
            while (LeRs.Read())
            {
                this.gUser.Rows.Add(LeRs.GetInt32(0), LeRs.GetString(1), LeRs.GetString(2),"");
            }
            LeRs.Close();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Common.FormVide(this);
            this.tID.Text = "";
            this.tNom.Text = "";
            this.lCon.SelectedIndex = -1;
            this.tSQL.Text = "";
            this.gUser.Rows.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            fConnect fc = new fConnect();
            fc.ShowDialog();
            string laVal = Common.ComboGetVal(this.lCon);
            Common.ComboFill("Select constr_id, Nom from extract.ConnectionString", this.lCon, ref LaConnect);
            Common.ComboSetVal(laVal,this.lCon);
            fc.Dispose();
        }

        private void lVue_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //List<Common.SQLchamp> LesChamps = new List<Common.SQLchamp>(Common.FormRecupereChamp(this));
            int Lid = Common.FormEnreg(this, "extract.vue", ref LaConnect);
            if (Lid !=0)  {
                //Enregistre les Users
                string LesUser = "0";
                string LesNouveaux = "";
                string LesModifs = "";
                foreach (DataGridViewRow L in (this.gUser.Rows))
                {
                    if (Common.Nz(L.Cells["User"].Value,"").ToString() != "") { 

                        if (Common.Nz(L.Cells["ID"].Value,"").ToString() != "") 
                        {
                            LesUser += "," + L.Cells["ID"].Value;
                            if (Common.Nz(L.Cells["MOD"].Value, "").ToString() != "")
                            {
                                LesModifs += "update extract.vueuser set CritSql='" + Common.Txt2Sql(Common.Nz(L.Cells["CritSql"].Value, "").ToString()) + "', UserLogin='" + L.Cells["User"].Value.ToString() + "' where VueUser_Id=" + L.Cells["ID"].Value + ";";
                            }
                        }
                        else {
                            LesNouveaux += "insert into extract.vueuser (vue_id,UserLogin,CritSql) Values (" + Lid + ",'" + L.Cells["User"].Value.ToString() + "','" + Common.Txt2Sql(Common.Nz(L.Cells["CritSql"].Value,"").ToString()) + "');";
                        }


                    }
                }
                Common.SqlDo("Delete from extract.vueuser where VueUser_Id not in (" + LesUser + ") and vue_id=" + Lid, ref LaConnect);
                if (LesModifs != "") { Common.SqlDo(LesModifs, ref LaConnect); }
                if (LesNouveaux != "") { Common.SqlDo(LesNouveaux, ref LaConnect); }

                //MAZ des champs
                Common.FormVide(this);
                this.gUser.Rows.Clear();
                Common.ListFill("Select vue.vue_id,vue.nom from extract.vue order by vue.nom", this.lVue, ref LaConnect);
            }
        }

        private void gUser_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            LaCelVal =Common.Nz(this.gUser.Rows[e.RowIndex].Cells[e.ColumnIndex].Value,"").ToString();
        }

        private void gUser_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (Common.Nz(this.gUser.Rows[e.RowIndex].Cells[e.ColumnIndex].Value,"").ToString() != LaCelVal)
            {
                this.gUser.Rows[e.RowIndex].Cells["MOD"].Value="x";
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Common.SqlDo("Delete from extract.vueuser where vue_id=" + ((Common.ListItem)(this.lVue.SelectedItem)).Val, ref LaConnect);
            Common.SqlDo("Delete from extract.vue where vue_id=" + ((Common.ListItem)(this.lVue.SelectedItem)).Val, ref LaConnect);
            Common.FormVide(this);
            Common.ListFill("Select vue.vue_id,vue.nom from extract.vue order by vue.nom", this.lVue, ref LaConnect);
        }
    }
}
