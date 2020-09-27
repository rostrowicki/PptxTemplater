using System;
using System.Text.RegularExpressions;

namespace PptxTemplater
{
    /// <summary>
    /// Text manipulation features.
    /// </summary>
    public static class TextTransformationHelper
    {
        /// <summary>
        /// Converts HTML to the plain text by removing XML tags and replacing literals with ascii characters.
        /// </summary>
        /// <param name="html">Valid HTML</param>
        /// <returns>Plain text string.</returns>
        public static string HtmlToPlainTxt(string html)
        {
            string plainText = String.Empty;
            if (html != null)
            {
                plainText = Regex.Replace(html, "<[^>]*>", "");
                plainText = plainText.Replace("&amp;", "&");
                plainText = plainText.Replace("&nbsp;", " ");
                plainText = plainText.Replace("&lt;", "<");
                plainText = plainText.Replace("&le;", "≤");
                plainText = plainText.Replace("&gt;", ">");
                plainText = plainText.Replace("&ge;", "≥");
                plainText = plainText.Replace("&quot;", "\"");
                plainText = plainText.Replace("&apos;", "'");
                plainText = plainText.Replace("&euro;", "€");

                // remove multiple \r\n
                var br = "[[[BR]]]";
                plainText = Regex.Replace(plainText, @"\r\n?|\n", br);
                while (plainText.IndexOf(br + " " + br) >= 0)
                {
                    plainText = plainText.Replace((br + " " + br), br);
                }
                while (plainText.IndexOf(br + br) >= 0)
                {
                    plainText = plainText.Replace((br + br), br);
                }
                if (plainText.Length > br.Length)
                {
                    while (plainText.StartsWith(br))
                    {

                        plainText = plainText.Substring(br.Length + 1);
                    }
                }

                plainText = plainText.Replace(br, System.Environment.NewLine);

                // trim each line individually               
                var finalText = new System.Text.StringBuilder();
                var textlines = plainText.Split(Environment.NewLine.ToCharArray());
                foreach (var line in textlines)
                {
                    finalText.Append(line.Trim() + System.Environment.NewLine);
                }
                plainText = finalText.ToString();
            }

            // remove excessing new lines at the begining and the end of the text
            if (plainText.Length > 4)
            {
                while (plainText.StartsWith("\r\n"))
                {
                    plainText = plainText.Substring(2);
                }
            }
            if (plainText.Length > 4)
            {
                while (plainText.EndsWith("\r\n"))
                {
                    plainText = plainText.Substring(0, plainText.Length - 2);
                }
            }

            return plainText;
        }

    }
}
