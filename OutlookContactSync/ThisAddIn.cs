
// using System;
// using System.Collections.Generic;
// using System.Linq;
// using System.Text;
// using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;


namespace OutlookContactSync
{


    public partial class ThisAddIn
    {

        Outlook.Inspectors inspectors;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            // http://stackoverflow.com/questions/24532211/outlook-add-in-outlook-shutdown-event
            ((Outlook.ApplicationEvents_11_Event)Application).Quit
+= new Outlook.ApplicationEvents_11_QuitEventHandler(OnQuit);



            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
            new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(OnInspect);

            // Contacts
            SyncContacts();

        } // End Sub ThisAddIn_Startup


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        { } // End Sub ThisAddIn_Shutdown


        // Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        // End VSTO gen. Code
        
        
    } // End Class ThisAddIn


} // End Namespace OutlookContactSync
