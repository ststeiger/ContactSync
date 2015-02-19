
namespace OutlookContactSync
{


    public class Utilities
    {
        // bool invalid = false;


        public static bool IsValidEmail(string email)
        {
            bool ret = false;

            try
            {
                System.Net.Mail.MailAddress addr = new System.Net.Mail.MailAddress(email);
                ret = addr.Address == email;
                addr = null;
                return ret;
            }
            catch
            { }

            return ret;
        }


        // Returns wrong result for
        // strIn = "user@127.0.0.1";
        // https://msdn.microsoft.com/en-us/library/01escwtf(v=vs.110).aspx
        public static bool TooSimpleIsValidEmail(string strIn)
        {
            bool invalid = false;
            if (string.IsNullOrEmpty(strIn))
                return false;

            // Use IdnMapping class to convert Unicode domain names. 
            try
            {
                // strIn = System.Text.RegularExpressions.Regex.Replace(strIn, @"(@)(.+)$", this.DomainMapper,
                //                       System.Text.RegularExpressions.RegexOptions.None, System.TimeSpan.FromMilliseconds(200));

                // strIn = System.Text.RegularExpressions.Regex.Replace(strIn, @"(@)(.+)$", this.DomainMapper, System.Text.RegularExpressions.RegexOptions.None);

                strIn = System.Text.RegularExpressions.Regex.Replace(strIn, @"(@)(.+)$", delegate(System.Text.RegularExpressions.Match match)
                {
                    // IdnMapping class with default property values.
                    System.Globalization.IdnMapping idn = new System.Globalization.IdnMapping();

                    string domainName = match.Groups[2].Value;
                    try
                    {
                        domainName = idn.GetAscii(domainName);
                    }
                    catch (System.ArgumentException)
                    {
                        invalid = true;
                    }
                    return match.Groups[1].Value + domainName;
                }
                , System.Text.RegularExpressions.RegexOptions.None);


            }
            //catch (System.Text.RegularExpressions.RegexMatchTimeoutException)
            catch (System.TimeoutException)
            {
                return false;
            }

            if (invalid)
                return false;

            // Return true if strIn is in valid e-mail format. 
            try
            {
                return System.Text.RegularExpressions.Regex.IsMatch(strIn,
                      @"^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                      @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$",
                      System.Text.RegularExpressions.RegexOptions.IgnoreCase); //, System.TimeSpan.FromMilliseconds(250));
            }
            //catch (System.Text.RegularExpressions.RegexMatchTimeoutException)
            catch (System.TimeoutException)
            {
                return false;
            }
        } // End Function TooSimpleIsValidEmail


        //private string DomainMapper(System.Text.RegularExpressions.Match match)
        //{
        //    // IdnMapping class with default property values.
        //    System.Globalization.IdnMapping idn = new System.Globalization.IdnMapping();

        //    string domainName = match.Groups[2].Value;
        //    try
        //    {
        //        domainName = idn.GetAscii(domainName);
        //    }
        //    catch (System.ArgumentException)
        //    {
        //        invalid = true;
        //    }
        //    return match.Groups[1].Value + domainName;
        //}


    } // End Class Utilities 


} // End Namespace OutlookContactSync 
