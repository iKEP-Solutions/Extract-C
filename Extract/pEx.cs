using System;
using System.Data.OleDb;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Extract
{
    public partial class pEx : UserControl

   {

     

        private OleDbConnection LaConnect = new OleDbConnection();

        private  void AfficheVue()
        {
            Common.ListFill("Select vue.vue_id,vue.Nom from extract.vue inner join extract.VueUser on vue.Vue_id= VueUser.Vue_Id where userlogin='" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "' order by Nom", lVue, ref LaConnect);
        }


        public pEx()
        {
            InitializeComponent();
        }

        private bool Initialise()
        {
            Microsoft.Office.Interop.Excel.Application APP = Globals.CompoExtract.Application;
            bool Init = false;
            try
            {
                lVue.Items.Clear();
                APP.StatusBar = "Vues...";
                tInit.Text = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                Properties.Settings.Default.Reload();
                Common.ConnexionInit(Properties.Settings.Default.VueConStr, ref LaConnect);
                if (Common.ConnexionTest(LaConnect.ConnectionString))
                {
                    Init = true;
                    AfficheVue();
                } else
                {
                    Init = false;
                }
            }
            catch
            {
                Init = false;
                tInit.Text = "Non Connecté";
            }
            i_info.Enabled = Init;
            APP.StatusBar = "";
            return Init;
        }

        private void pEx_Load(object sender, EventArgs e)
        {
           Initialise();
        }

        private void tInit_Click(object sender, EventArgs e) 
        {
            Form frm = new pParam();
            if (frm.ShowDialog() == DialogResult.OK) { Initialise();}
        }

        private void lVue_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void lVue_MouseDoubleClick(object sender, EventArgs e)
        {
            string sSQL = "";
            string ExtractConStr = "";
            OleDbConnection ExtractConSql = new OleDbConnection();
            
            if (lVue.SelectedIndex < 0) { return; }

            try
            {
                Microsoft.Office.Interop.Excel.Application APP = Globals.CompoExtract.Application;
                APP.Calculation = XlCalculation.xlCalculationManual;

                OleDbDataReader LeRs;
                Common.ListItem LItem = (Common.ListItem)lVue.SelectedItem;

                sSQL = "Select CmdSql, critsql,connectionstring from extract.vue inner join extract.VueUser on vue.Vue_id= VueUser.Vue_Id Inner join extract.ConnectionString "
                     + " on vue.ConStr_id = ConnectionString.ConStr_id where userlogin='" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "' and vue.vue_id=" + LItem.Val;
                LeRs = Common.SqlLit(sSQL, ref LaConnect);

                if (LeRs.Read())
                {
                    //on créer la requete
                    sSQL = LeRs.GetString(0) + LeRs.GetString(1);  //+ " " + LeRs.GetString(1);
                    ExtractConStr = LeRs.GetString(2);

                }

//                int curWS = APP.ActiveSheet.index;

                //            Dim curWS As Integer = APP.ActiveSheet.index
                //            Dim parWS As Integer = 0
                //            Dim curCol = APP.ActiveCell.Column
                //            Dim curlig = APP.ActiveCell.Row
                if (APP.ActiveCell.ListObject == null) {
                    QueryTable Q =  APP.ActiveSheet.ListObjects.Add(SourceType: 0, Source: "OLEDB;" + ExtractConStr, Destination: APP.ActiveCell).QueryTable;
                    Q.CommandType = XlCmdType.xlCmdSql;
                    Q.CommandText = sSQL;
                    Q.RowNumbers = false;
                    Q.FillAdjacentFormulas = false;
                    Q.PreserveFormatting = true;
                    Q.RefreshOnFileOpen = false;
                    Q.BackgroundQuery = false;
                    Q.SavePassword = false;
                    Q.SaveData = true;
                    Q.AdjustColumnWidth = true;
                    Q.RefreshPeriod = 0;
                    Q.PreserveColumnInfo = true;
                    Q.Refresh(BackgroundQuery: false);
                } else
                {
                    MessageBox.Show("Impossible de mettre à jour ce tableau. Veuillez extraire dans un autre onglet !");
                }
                APP.Calculation = XlCalculation.xlCalculationAutomatic;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            lVue_MouseDoubleClick(sender, e);
        }

        private void Param(object sender, EventArgs e)
        {
            Form passe = new fPasse();
            if (passe.ShowDialog()== DialogResult.OK) {
                Form leP = new fParamSQL();
                leP.ShowDialog();
                AfficheVue();
            }


        }

    }
}
