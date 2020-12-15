using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Extract
{
    public partial class pParam : Form
    {

        public pParam()
        {
            InitializeComponent();
        }

        private void pParam_Load(object sender, EventArgs e)
        {
            Extract.Properties.Settings.Default.Reload();
            pGrid1.SelectedObject = Extract.Properties.Settings.Default;
            Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = "Paramètres " + version.ToString();

        }

        private void Button1_Click(object sender, EventArgs e)
        {
           Extract.Properties.Settings.Default.Reload();
            if (Common.ConnexionTest(Extract.Properties.Settings.Default.VueConStr2)) {
                MessageBox.Show("Connexion OK");
            } else {
                MessageBox.Show("Erreur Connection");
            }
        }

        private void bAnnul_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void bOK_Click(object sender, EventArgs e)
        {
            Extract.Properties.Settings.Default.Save();
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
