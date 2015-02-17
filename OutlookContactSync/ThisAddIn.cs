
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;


namespace OutlookContactSync
{


    public partial class ThisAddIn
    {

        public void Deploy()
        { 
            // http://blogs.msdn.com/b/mcsuksoldev/archive/2010/10/01/building-and-deploying-an-outlook-2010-add-in-part-2-of-2.aspx
            // https://msdn.microsoft.com/en-us/library/cc136646(v=office.12).aspx
            // http://blogs.msdn.com/b/mcsuksoldev/archive/2010/07/12/building-and-deploying-an-outlook-2010-add-in-part-1-of-2.aspx


            // ◦Microsoft .NET Framework 4 Client Profile (x86 and x64) 
            // ◦Microsoft Visual Studio 2010 Tools for Office Runtine (x86 and x64) 
            // (VSTO)

            // To get our add-in working, we will need to add registry keys. 
            // HKCU\Software\Microsoft\Office\Outlook\Addins.

            
            // Windows Registry Editor Version 5.00

            // [HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookContactSync]
            // "Description"="OutlookContactSync"
            // "FriendlyName"="OutlookContactSync"
            // "LoadBehavior"=dword:00000003
            // "Manifest"="file:///d:/stefan.steiger/documents/visual studio 2013/Projects/OutlookContactSync/OutlookContactSync/bin/Debug/OutlookContactSync.vsto|vstolocal"



        }



        Outlook.Inspectors inspectors;

        void ThisAddIn_Quit()
        {
            // System.Threading.Thread.Sleep(10 * 1000);
            System.Windows.Forms.MessageBox.Show("bye bye problem, I found the solution!!");
        }


        public void OnWrite(ref bool Cancel)
        {
            MsgBox("OnWrite");
        }


        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {

            Outlook.ContactItem contactitem = Inspector.CurrentItem as Outlook.ContactItem;
            if (contactitem != null)
            {
                MsgBox("address");
                // contactitem.BeforeAttachmentSave += contactitem_BeforeAttachmentSave;
                // contactitem.BeforeAutoSave += BeforeAutoSave;
                contactitem.Write += OnWrite;
            }

            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "This text was added by using code";
                    mailItem.Body = "This text was added by using code";
                }

            }
        }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            // http://stackoverflow.com/questions/24532211/outlook-add-in-outlook-shutdown-event
            ((Outlook.ApplicationEvents_11_Event)Application).Quit
+= new Outlook.ApplicationEvents_11_QuitEventHandler(ThisAddIn_Quit);



            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
            new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);


            Outlook.ContactItem ca = null;

            // http://www.gregthatcher.com/scripts/vba/outlook/getlistofcontacts.aspx
            Outlook.MAPIFolder fo = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);


            
            

            Outlook.Recipient x = Application.Session.CurrentUser;
            Outlook.ExchangeUser exchangeUser = x.AddressEntry.GetExchangeUser();
            string userMail = null;
            if(exchangeUser != null)
                userMail = exchangeUser.PrimarySmtpAddress;
            else
                userMail = System.DirectoryServices.AccountManagement.UserPrincipal.Current.EmailAddress;

            if(string.IsNullOrEmpty(userMail))
                // userMail = this.Application.ActiveExplorer().Session.CurrentUser.Address;
                userMail = Application.Session.CurrentUser.Address;
            
            MsgBox(userMail);


            // string lol = x.Address;
            // string lol = x.AddressEntry.Address;

            // x.AddressEntry.ID
            // string name = x.AddressEntry.Name;
            // string name2 = x.Name;
            

            
            Outlook.Items contacts = fo.Items.Restrict("[MessageClass]='IPM.Contact'");

            // Find existinc contact
             // Outlook.ContactItem existingContact = (Outlook.ContactItem)contacts.Find("[Email1Address] = '" + dr["EmailID"] + "'");


            foreach (Outlook.ContactItem contact in contacts)
            {
                System.Console.WriteLine(contact.EntryID.Length);
                
            }


            
            

            // ca.EntryID

        }



        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        public static void MsgBox(object obj)
        {

            if (obj != null)
                System.Windows.Forms.MessageBox.Show(obj.ToString());
            else
                System.Windows.Forms.MessageBox.Show("obj IS NULL");
        }



        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }


}
