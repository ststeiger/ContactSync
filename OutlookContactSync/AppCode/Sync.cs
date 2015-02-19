
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;


namespace OutlookContactSync
{


    public partial class ThisAddIn
    {


        public void SyncContacts()
        {
            string userMail = GetPrimaryMailAddress();

            // http://www.gregthatcher.com/scripts/vba/outlook/getlistofcontacts.aspx
            Outlook.MAPIFolder fo = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);

            // Outlook.ContactItem ca = null;
            // ca.EntryID

            // https://msdn.microsoft.com/en-us/library/office/ee692172(v=office.14).aspx
            // Table 1. Ribbon IDs and Message Class
            Outlook.Items contacts = fo.Items.Restrict("[MessageClass]='IPM.Contact'");

            // Find existinc contact
            // Outlook.ContactItem existingContact = (Outlook.ContactItem)contacts.Find("[Email1Address] = '" + dr["EmailID"] + "'");

            foreach (Outlook.ContactItem contact in contacts)
            {
                System.Console.WriteLine(contact.EntryID.Length);
            } // Next contact

            
        } // End Sub SyncContacts 


    } // End partial Class ThisAddIn 


} // End Namespace OutlookContactSync 
