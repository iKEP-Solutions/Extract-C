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
    public partial class fPasse : Form
    {
        public fPasse()
        {
            InitializeComponent();
        }

        private void bOK_Click(object sender, EventArgs e)
        {
            if (this.tPasse.Text =="!KEP") {
                DialogResult = DialogResult.OK;
                this.Close();
            } else
            {
                DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        private void bAnnul_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void fPasse_Load(object sender, EventArgs e)
        {
            this.tPasse.Text = "";
        }
    }
}
