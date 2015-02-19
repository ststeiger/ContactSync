
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;


namespace OutlookContactSync
{


    public partial class ThisAddIn
    {


        public void OnWrite(ref bool Cancel)
        {
            MsgBox("OnWrite");
        } // End Sub OnWrite


        void OnInspect(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.ContactItem contactitem = Inspector.CurrentItem as Outlook.ContactItem;
            if (contactitem != null)
            {
                MsgBox("Item is an address !");
                // contactitem.BeforeAttachmentSave += contactitem_BeforeAttachmentSave;
                // contactitem.BeforeAutoSave += BeforeAutoSave;
                contactitem.Write += OnWrite;
            } // End if (contactitem != null)

            // return;

            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "OMG";
                    // mailItem.Body = "This text was added by using code";
                    // mailItem.HTMLBody = "This <del>text</del> was added by using code";
                    mailItem.HTMLBody = @"

<br /><br /><br />
Freundliche Gr&uuml;sse / Bien cordialement / Cordiali saluti / Kind regards
<br /><br />
Stefan Steiger<br />
<img src=""http://i.stack.imgur.com/nwphb.gif"" alt=""Logo COR"" />
<br />
Fabrikstrasse 1<br />
CH-8586 Erlen/TG<br />
Schweiz/Suisse/Svizzera/Switzerland<br /><br />
";
                } // End if (mailItem.EntryID == null)

            } // End if (mailItem != null)

        } // End Sub Inspectors_NewInspector 


    } // End partial Class ThisAddIn 


} // End Namespace OutlookContactSync 
