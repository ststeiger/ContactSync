
using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;


namespace GoogleContacts
{


    public partial class _Default : System.Web.UI.Page
    {


        protected void Page_Load(object sender, EventArgs e)
        {

        }



        protected void btnOK_Click(object sender, EventArgs e)
        {
            string appName = "My Web-Application";
            string uname = this.txtUserName.Text;
            string pw = this.txtPassword.Text;

            this.gvData.DataSource = Google.Contacts.API.GetGmailContacts(appName, uname, pw);
            this.gvData.DataBind();
        }


    }
}