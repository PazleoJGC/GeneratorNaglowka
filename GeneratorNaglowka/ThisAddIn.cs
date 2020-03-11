using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace GeneratorNaglowka
{
    public partial class ThisAddIn
    {
        private Kontrolka kontrolka1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            kontrolka1 = new Kontrolka();
            myCustomTaskPane = this.CustomTaskPanes.Add(kontrolka1, "Generator Nagłówka");
            myCustomTaskPane.VisibleChanged += new EventHandler(myCustomTaskPane_VisibleChanged);
            myCustomTaskPane.Width = 290;


        }

        private void myCustomTaskPane_VisibleChanged(object sender, System.EventArgs e)
        {
            Globals.Ribbons.Ribbon1.toggleButton1.Checked = myCustomTaskPane.Visible;
        }

        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get
            {
                return myCustomTaskPane;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }      

        #region Kod wygenerowany przez program VSTO

        /// <summary>
        /// Wymagana metoda obsługi projektanta — nie należy modyfikować 
        /// zawartość tej metody z edytorem kodu.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
