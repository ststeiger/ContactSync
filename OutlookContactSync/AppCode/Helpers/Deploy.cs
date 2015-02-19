
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

        } // End Sub Deploy


    } // End partial Class ThisAddIn 


} // End Namespace OutlookContactSync 
