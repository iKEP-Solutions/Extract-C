using System;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Extract
{
    public partial class fConnect : Form
    {
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

        public fConnect()
        {
            InitializeComponent();
        }

        private void AfficheConnect()
        {
            Common.ListFill("Select constr_id, Nom from extract.ConnectionString", this.lCon, ref LaConnect);

        }

        private void fConnect_Load(object sender, EventArgs e)
        {
            Initialise();
            AfficheConnect();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Common.FormVide(this);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (Common.ConnexionTest(this.tConStr.Text)  )
            {
                MessageBox.Show("Connexion OK");
            }
            else
            {
                MessageBox.Show("Echec Connexion");
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Supprimer la connexion" + lCon.SelectedItem + "?","Attention", MessageBoxButtons.OKCancel) == DialogResult.OK) { 
                Common.SqlDo("Delete from extract.ConnectionString where ConStr_Id=" + Common.ListGetVal(this.lCon), ref LaConnect);
                Common.FormVide(this);
                AfficheConnect();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
           // string laVal = Common.ListGetVal(this.lCon);
            Common.FormEnreg(this, "extract.ConnectionString", ref LaConnect);
            AfficheConnect();
            Common.FormVide(this);
            //Common.ListSetVal(laVal, this.lCon);
        }

        private void dbl(object sender, EventArgs e)
        {

            Common.FormRempli(this, "SELECT constr_id,nom,ConnectionString FROM Extract.ConnectionString where ConStr_Id=" + Common.ListGetVal(this.lCon), ref LaConnect);


        }
    }
}
