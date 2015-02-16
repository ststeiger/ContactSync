using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


using Google.Contacts;

using Google.GData.Contacts; // Google.GData.Contacts.dll
using Google.GData.Client; // Google.GData.Client.dll
using Google.GData.Extensions; // Google.GData.Extensions.dll




// https://developers.google.com/google-apps/contacts/v3/
// https://msdn.microsoft.com/en-us/library/bb410039(v=office.12).aspx

// http://www.dotnetspark.com/kb/1878-import-gmail-contacts-into-asp-net-gridview.aspx
// http://stackoverflow.com/questions/14064635/not-able-to-add-contact-in-google-contacts-using-authsub-in-asp-net
// http://csharpdotnetfreak.blogspot.com/2011/01/import-gmail-contacts-in-aspnet.html
namespace Google.Contacts
{


    public class API
    {

        // GAPI: Failed to authenticate user. 
        // Error=BadAuthenticationUrl=https://www.google.com/accounts/ServiceLogin?service=analytics#Email=mail&Info=WebLoginRequired
        // You need to approve your account 
        // before you start fetching data for the first time at this url 
        // https://accounts.google.com/DisplayUnlockCaptcha

        // Allow insecure devices
        // I had a similar problem. 
        // On my account, these failed attempts came up in the account activity pane.
        // https://security.google.com/settings/security/activity

        // Execution of request returned unexpected result: 
        // http://www.google.com/m8/feeds/contacts/default/fullMovedPermanently
        // Install-Package Google.GData.Contacts from your nuget pm your issue will be resolved


        // https://developers.google.com/google-apps/contacts/v3/#running_the_sample_code

        public static System.Data.DataTable GetGmailContacts(string App_Name, string Uname, string UPassword)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("EmailID", typeof(string));

            // DataColumn C2 = new DataColumn();
            // //C2.DataType = Type.GetType("System.String");
            // C2.DataType = typeof(string);
            // C2.ColumnName = "EmailID";
            // dt.Columns.Add(C2);
            
            RequestSettings rsLoginInfo = new RequestSettings(App_Name, Uname, UPassword);
            rsLoginInfo.AutoPaging = true;


            ContactsRequest contactRequest = new ContactsRequest(rsLoginInfo);
            Feed<Google.Contacts.Contact> f = contactRequest.GetContacts();

            foreach (Google.Contacts.Contact t in f.Entries)
            {
                foreach (Google.GData.Extensions.EMail email in t.Emails)
                {
                    System.Data.DataRow dr = dt.NewRow();
                    dr["EmailID"] = email.Address.ToString();
                    dt.Rows.Add(dr);
                } // Next email 

            } // Next t

            return dt;
        } // End Function GetGmailContacts 


    }


}