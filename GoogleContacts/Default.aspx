<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="GoogleContacts._Default" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Google-Contacts Example</title>

    <style type="text/css">

        label
        {
            display: inline-block;
            width: 2cm;
            margin: 0px;
            padding: 0px;
            padding-right: 0.25cm;   
        }

        input[type="text"],input[type="password"] 
        {
            display: inline-block;
            width: 8cm;
            margin: 0px;
            padding: 0px;
            padding-left: 0.25cm;   
            padding-right: 0.25cm;   
        }

    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <label for="txtUserName">Username</label>
        <asp:TextBox ID="txtUserName" ClientIDMode="Static" placeholder="username@gmail.com" runat="server" />
        <br /><br />
        <label for="txtPassword">Password</label>
        <asp:TextBox ID="txtPassword" ClientIDMode="Static" placeholder="TopSecret123"  runat="server" TextMode="Password"  /><br /><br />
        <asp:Button ID="btnOK" runat="server" Text="OK" OnClick="btnOK_Click" />
        <br /><br />
        <asp:GridView ID="gvData" runat="server"></asp:GridView>
    </div>
    </form>
</body>
</html>
