
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;


namespace OutlookContactSync
{


    public partial class ThisAddIn
    {


        void OnQuit()
        {
            // System.Threading.Thread.Sleep(10 * 1000);
            // MsgBox("Bye bye problem, I found the solution!!");
            System.Console.WriteLine("Bye bye problem, I found the solution!!");
        } // End Sub ThisAddIn_Quit


    } // End partial Class ThisAddIn 


} // End Namespace OutlookContactSync 
