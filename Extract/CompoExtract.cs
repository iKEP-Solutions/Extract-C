using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace Extract
{
    public partial class CompoExtract
    {
        private void CompoExtract_Startup(object sender, System.EventArgs e)
        {
        }

        private void CompoExtract_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Code généré par VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(CompoExtract_Startup);
            this.Shutdown += new System.EventHandler(CompoExtract_Shutdown);
        }
        
        #endregion
    }
}
