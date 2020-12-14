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
    public partial class fCrit : Form
    {
        public string laReq;
        private int nbCtrl = 0;

        public fCrit()
        {
            InitializeComponent();
        }

        private void fCrit_Load(object sender, EventArgs e)
        {
            foreach (string s in laReq.Split('?'))
            {
                int p = s.IndexOf(']');
                if (p > 0) { 


                    if (s.Substring(0, 2) == "D[")
                    {
                        Label aL = new Label();
                        aL.Text = s.Substring(2, p - 2);
                        aL.Location = new Point(12, nbCtrl * 27 + 12);
                        this.Controls.Add(aL);

                        DateTimePicker aD = new DateTimePicker();
                        aD.Width = 130;
                        aD.Format = DateTimePickerFormat.Short;
                        aD.Location = new Point(112, nbCtrl * 27 + 12);
                        aD.Tag = s.Substring(0, p + 1);
                        this.Controls.Add(aD);

                    }

                    if (s.Substring(0, 2) == "T[")
                    {
                        Label aL = new Label();
                        aL.Text = s.Substring(2, p - 2);
                        aL.Location = new Point(12, nbCtrl * 27 + 12);
                        this.Controls.Add(aL);

                        TextBox aT = new TextBox();
                        aT.Width = 130;
                        aT.Location = new Point(112, nbCtrl * 27 + 12);
                        aT.Tag = s.Substring(0, p+1);
                        this.Controls.Add(aT);
                    }

                    if (s.Substring(0, 2) == "B[" && p > 0)
                    {
    
                    }

                    nbCtrl += 1;
                }

            }
            this.Height += nbCtrl * 27;
        }

        private void bAnnul_Click(object sender, EventArgs e)
        {
            laReq = "";
            DialogResult = DialogResult.Cancel;
        }

        private void bOK_Click(object sender, EventArgs e)
        {
            foreach (Control c in this.Controls)
            {
                    if (c.GetType().ToString().Contains("TextBox"))  { laReq = laReq.Replace("?" + c.Tag, c.Text); }

                    if (c.GetType().ToString().Contains("DateTimePicker")) { laReq = laReq.Replace("?" + c.Tag, "'" + ((DateTimePicker)(c)).Value.ToString("dd/MM/yyyy")   + "'"); }
            }
            DialogResult = DialogResult.OK;
        }
    }
}
