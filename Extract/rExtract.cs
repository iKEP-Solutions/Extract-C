using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Data.OleDb;

namespace Extract
{
    public partial class rExtract
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }



        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            bool PaneTrouve = false;
            string PaneName = "Extract";
            Microsoft.Office.Tools.CustomTaskPane ctp;

            foreach (Microsoft.Office.Tools.CustomTaskPane Pane in Globals.CompoExtract.CustomTaskPanes)
            {
                try
                {

                    if (object.ReferenceEquals(Pane.Window, Globals.CompoExtract.Application.ActiveWindow) && Pane.Title == PaneName)
                    {
                        Pane.Visible = !Pane.Visible;
                        PaneTrouve = true;
                    }
                } catch                 
                {
                }
            }

            if (PaneTrouve == false)
            {
                pEx MypEX = new pEx();
                ctp = Globals.CompoExtract.CustomTaskPanes.Add(MypEX, PaneName);
                ctp.Visible = true;
            }
                       
        }

    }
}
