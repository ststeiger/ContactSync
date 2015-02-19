
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;


namespace OutlookContactSync
{


    public partial class ThisAddIn
    {


        public string GetPrimaryMailAddress()
        {
            string userMail = null;
            Outlook.Recipient cu = Application.Session.CurrentUser;
            if (cu != null)
            {
                userMail = cu.Address;
                // userMail = cu.AddressEntry.Address;
                // userMail = cu.AddressEntry.ID;
                // userMail = cu.AddressEntry.Name;
                // userMail = cu.Name;
                
                if (!Utilities.IsValidEmail(userMail))
                {
                    // userMail = this.Application.ActiveExplorer().Session.CurrentUser.Address;
                    Outlook.ExchangeUser exchangeUser = cu.AddressEntry.GetExchangeUser();

                    if (exchangeUser != null)
                        userMail = exchangeUser.PrimarySmtpAddress;
                } // End if (!Utilities.IsValidEmail(userMail))

            } // End if (cu != null)

            // if (string.IsNullOrEmpty(userMail)) userMail = System.DirectoryServices.AccountManagement.UserPrincipal.Current.EmailAddress;

            if (string.IsNullOrEmpty(userMail))
                userMail = string.Format("{0}{1}@{2}.local"
                    , string.IsNullOrEmpty(System.Environment.UserDomainName) ? ""
                      : System.Environment.UserDomainName + @"\"
                    , System.Environment.UserName
                    , System.Environment.MachineName
                );


            // userMail = "user@127.0.0.1";

            if (Utilities.IsValidEmail(userMail))
                System.Console.WriteLine(userMail);
            else
                MsgBox("Not a mail address !");

            // MsgBox(userMail);
            return userMail;
        } // End Function GetPrimaryMailAddress


        public static void MsgBox(object obj)
        {
            if (obj != null)
                System.Windows.Forms.MessageBox.Show(obj.ToString());
            else
                System.Windows.Forms.MessageBox.Show("obj IS NULL");
        } // End Sub MsgBox


    } // End partial Class ThisAddIn 


} // End Namespace OutlookContactSync 
