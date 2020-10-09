namespace PptxTemplater
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using MariGold.OpenXHTML;

    using A = DocumentFormat.OpenXml.Drawing;

    /// <summary>
    /// Represents a paragraph inside a PowerPoint file.
    /// </summary>
    /// <remarks>
    /// Could not simply be named Paragraph, conflicts with DocumentFormat.OpenXml.Drawing.Paragraph.
    ///
    /// Structure of a paragraph:
    /// a:p (Paragraph)
    ///  a:r (Run)
    ///   a:t (Text)
    ///
    /// <![CDATA[
    /// <a:p>
    ///  <a:r>
    ///   <a:rPr lang="en-US" dirty="0" smtClean="0"/>
    ///   <a:t>
    ///    Hello this is a tag: {{hello}}
    ///   </a:t>
    ///  </a:r>
    ///  <a:endParaRPr lang="fr-FR" dirty="0"/>
    /// </a:p>
    ///
    /// <a:p>
    ///  <a:r>
    ///   <a:rPr lang="en-US" dirty="0" smtClean="0"/>
    ///   <a:t>
    ///    Another tag: {{bonjour
    ///   </a:t>
    ///  </a:r>
    ///  <a:r>
    ///   <a:rPr lang="en-US" dirty="0" smtClean="0"/>
    ///   <a:t>
    ///    }} le monde !
    ///   </a:t>
    ///  </a:r>
    ///  <a:endParaRPr lang="en-US" dirty="0"/>
    /// </a:p>
    /// ]]>
    /// </remarks>
    internal static class PptxParagraph
    {
        /// <summary>
        /// Replaces a tag inside a paragraph (a:p).
        /// </summary>
        /// <param name="p">The paragraph (a:p).</param>
        /// <param name="tag">The tag to replace by newText, if null or empty do nothing; tag is a regex string.</param>
        /// <param name="newText">The new text to replace the tag with, if null replaced by empty string.</param>
        /// <returns><c>true</c> if a tag has been found and replaced, <c>false</c> otherwise.</returns>
        internal static bool ReplaceTag(A.Paragraph p, string tag, string newText)
        {
            bool replaced = false;

            if (string.IsNullOrEmpty(tag))
            {
                return replaced;
            }

            if (newText == null)
            {
                newText = string.Empty;
            }
            newText = RemoveInvalidXMLChars(newText);

            while (true)
            {
                // Search for the tag
                Match match = Regex.Match(GetTexts(p), tag);
                if (!match.Success)
                {
                    break;
                }

                replaced = true;

                List<TextIndex> texts = GetTextIndexList(p);

                for (int i = 0; i < texts.Count; i++)
                {
                    TextIndex text = texts[i];
                    if (match.Index >= text.StartIndex && match.Index <= text.EndIndex)
                    {
                        // Got the right A.Text

                        int index = match.Index - text.StartIndex;
                        int done = 0;

                        for (; i < texts.Count; i++)
                        {
                            TextIndex currentText = texts[i];
                            List<char> currentTextChars = new List<char>(currentText.Text.Text.ToCharArray());

                            for (int k = index; k < currentTextChars.Count; k++, done++)
                            {
                                if (done < newText.Length)
                                {
                                    if (done >= tag.Length - 1)
                                    {
                                        // Case if newText is longer than the tag
                                        // Insert characters
                                        int remains = newText.Length - done;
                                        currentTextChars.RemoveAt(k);
                                        currentTextChars.InsertRange(k, newText.Substring(done, remains));
                                        done += remains;
                                        break;
                                    }
                                    else
                                    {
                                        currentTextChars[k] = newText[done];
                                    }
                                }
                                else
                                {
                                    if (done < tag.Length)
                                    {
                                        // Case if newText is shorter than the tag
                                        // Erase characters
                                        int remains = tag.Length - done;
                                        if (remains > currentTextChars.Count - k)
                                        {
                                            remains = currentTextChars.Count - k;
                                        }
                                        currentTextChars.RemoveRange(k, remains);
                                        done += remains;
                                        break;
                                    }
                                    else
                                    {
                                        // Regular case, nothing to do
                                        //currentTextChars[k] = currentTextChars[k];
                                    }
                                }
                            }

                            currentText.Text.Text = new string(currentTextChars.ToArray());
                            index = 0;
                        }
                    }
                }
            }

            return replaced;
        }

        /// <summary>
        /// Replaces text tag with hyperlink id. Hyperlink identified by Id hes to be a part of the slide.
        /// </summary>
        /// <param name="p">The paragraph (a:p).</param>
        /// <param name="tag">The tag to replace by newText, if null or empty do nothing; tag is a regex string.</param>
        /// <param name="newText">The new text to replace the tag with, if null replaced by empty string and not visible.</param>
        /// <param name="relationshipId">Hyperlink relationship Id. Relationship has to be existing on slide level.</param>
        /// <param name="fontName">Font name</param>
        /// <param name="fontSize">Font size. E.g. 800 is 8pt (small) font. If value is less than 100 it will be multiplied by 100 to keep up with PPT notation.</param>
        /// <returns></returns>
        internal static bool ReplaceTagWithHyperlink(A.Paragraph p, string tag, string newText, string relationshipId, string fontName = "Calibri", int fontSize = 800)
        {

            bool replaced = false;

            if (string.IsNullOrEmpty(tag))
            {
                return replaced;
            }

            if (newText == null)
            {
                newText = string.Empty;
            }
            newText = RemoveInvalidXMLChars(newText);

            while (true)
            {
                // Search for the tag
                Match match = Regex.Match(GetTexts(p), tag);
                if (!match.Success)
                {
                    break;
                }

                p.RemoveAllChildren(); // remove exisitng children then add new

                A.HyperlinkOnClick link = new A.HyperlinkOnClick() { Id = relationshipId };
                A.Run r = new A.Run();
                r.RunProperties = new A.RunProperties();
                A.Text at = new A.Text(newText);
                r.RunProperties.AppendChild(link);
                r.AppendChild(at);
                p.Append(r);

                var run = p.Descendants<A.Run>();
                foreach (var item in run)
                {
                    item.RunProperties.RemoveAllChildren<A.LatinFont>();
                    var latinFont = new A.LatinFont();
                    latinFont.Typeface = fontName;
                    item.RunProperties.AppendChild(latinFont);
                    item.RunProperties.FontSize = (fontSize > 99) ? fontSize : fontSize * 100; // e.g. translate value 8 (Power Point UI font size) to 800 for API
                }

                replaced = true;
            }

            return replaced;
        }

        /// <summary>
        /// Replaces a tag inside a paragraph (a:p) with parsed HTML
        /// </summary>
        /// <param name="p">The paragraph (a:p).</param>
        /// <param name="tag">The tag to replace by newText, if null or empty do nothing; tag is a regex string.</param>
        /// <param name="newText">The new text to replace the tag with, if null replaced by empty string and not visible.</param>
        /// <param name="fontName">Font name</param>
        /// <param name="fontSize">Font size. E.g. 800 is 8pt (small) font. If value is less than 100 it will be multiplied by 100 to keep up with PPT notation.</param>
        /// <param name="hyperlinks">URL Relationships dictionary. Relationship has to be defined on slide level.</param>
        /// <returns><c>true</c> if a tag has been found and replaced, <c>false</c> otherwise.</returns>
        internal static bool ReplaceTagWithHtml(A.Paragraph p, string tag, string newText, string fontName = null, int fontSize = 0, IDictionary<string, string> hyperlinks = null)
        {
            newText = CorrectUnhandledHtmlTags(newText); // e.g. deal with ul/li html tags
            bool isFirstLine = true; // avoiding unintentional empty line at the begining of the text
            bool replaced = false;
            string[] closingTags = new string[] { "div", "p" }; // tags that force line break in PPTX paragraph

            if (string.IsNullOrEmpty(tag))
            {
                return replaced;
            }

            if (newText == null)
            {
                newText = string.Empty;
            }
            newText = RemoveInvalidXMLChars(newText);

            while (true)
            {
                // Search for the tag
                Match match = Regex.Match(GetTexts(p), tag);
                if (!match.Success)
                {
                    break;
                }

                p.RemoveAllChildren(); // // remove exisitng children then add new

                HtmlParser hp = new HtmlParser(newText);
                MariGold.HtmlParser.IHtmlNode nodes = null;

                try
                {
                    nodes = hp.FindBodyOrFirstElement();
                }
                catch
                {
                    Console.WriteLine(String.Format("WARNING: HTML is empty or HTML schema has errors. Parsed HTML[]: [{0}]", newText));
                }

                while (nodes != null)
                {
                    foreach (var item in nodes.Children)
                    {
                        bool skipLineBreak = false;
                        A.Run r = new A.Run();
                        r.RunProperties = new A.RunProperties();

                        if (item.Html.Contains("<b>") || item.Html.Contains("<strong>")) { r.RunProperties.Bold = new DocumentFormat.OpenXml.BooleanValue(true); }
                        if (item.Html.Contains("<i>") || item.Html.Contains("<em>")) { r.RunProperties.Italic = new DocumentFormat.OpenXml.BooleanValue(true); }
                        if (item.Html.Contains("<u>") || item.Html.Contains("text-decoration: underline")) { r.RunProperties.Underline = new DocumentFormat.OpenXml.EnumValue<A.TextUnderlineValues>(A.TextUnderlineValues.Dash); }
                        if (item.Html.Contains("<a"))
                        {
                            string uriId = null;
                            try
                            {
                                string url = PptxSlide.ParseHttpUrls(item.Html).First().Value;
                                uriId = hyperlinks.Where(q => q.Value == url).FirstOrDefault().Key;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("URL is no available");
                            }
                            if (uriId != null)
                            {
                                A.HyperlinkOnClick link = new A.HyperlinkOnClick() { Id = uriId };
                                r.RunProperties.AppendChild(link);
                            }
                        }

                        A.Text at = new A.Text(PptxTemplater.TextTransformationHelper.HtmlToPlainTxt(item.InnerHtml) + " "); // clear not interpreted html tags
                        if (at.InnerText.Trim() == "" && isFirstLine)
                        {
                            at = new A.Text(); // avoid excessive new lines
                            isFirstLine = false;
                        }
                        r.AppendChild(at);
                        p.Append(r);

                        // LINE BREAK -- if outer tag is div add line break
                        if (closingTags.Contains(item.Parent.Tag) && skipLineBreak == false)
                        {
                            p.Append(new A.Break());
                        }
                    }

                    // remove parsed html part
                    newText = newText.Substring(nodes.Html.Length);
                    if (newText.Trim() == "")
                    {
                        break;
                    }
                    hp = new HtmlParser(newText);
                    nodes = hp.FindBodyOrFirstElement();
                }

                var run = p.Descendants<A.Run>();
                foreach (var item in run)
                {
                    if (fontName != null)
                    {
                        item.RunProperties.RemoveAllChildren<A.LatinFont>();
                        var latinFont = new A.LatinFont();
                        latinFont.Typeface = fontName;
                        item.RunProperties.AppendChild(latinFont);
                    }
                    if (fontSize > 0)
                    {
                        item.RunProperties.FontSize = (fontSize > 99) ? fontSize : fontSize * 100; // e.g. translate value 8 (Power Point UI font size) to 800 for API
                    }
                }

                replaced = true;
            }

            return replaced;
        }

        /// <summary>
        /// Removes characters that are invalid for XML encoding.
        /// </summary>
        /// <param name="input">Text to be encoded.</param>
        /// <returns>Text with invalid XML characters removed.</returns>
        /// <remarks>
        /// <see href="http://stackoverflow.com/questions/20762/how-do-you-remove-invalid-hexadecimal-characters-from-an-xml-based-data-source-p">How do you remove invalid hexadecimal characters from an XML-based data source</see>
        /// </remarks>
        private static string RemoveInvalidXMLChars(string input)
        {
            return new string(input.Where(value =>
                                (value >= 0x0020 && value <= 0xD7FF) ||
                                (value >= 0xE000 && value <= 0xFFFD) ||
                                value == 0x0009 ||
                                value == 0x000A ||
                                value == 0x000D).ToArray());
        }

        /// <summary>
        /// Changing unhandled HTML tags to suitable alternatives
        /// </summary>
        /// <param name="newText"></param>
        /// <returns></returns>
        private static string CorrectUnhandledHtmlTags(string newText)
        {
            newText = newText.Replace(@"<ul>", "<div>");
            newText = newText.Replace(@"</ul>", "</div>");
            newText = newText.Replace(@"<li>", "• ");
            newText = newText.Replace(@"</li>", Environment.NewLine);
            return newText?.Trim();
        }

        /// <summary>
        /// Returns all the texts found inside a given paragraph.
        /// </summary>
        /// <remarks>
        /// If all A.Text in the given paragraph are empty, returns an empty string.
        /// </remarks>
        internal static string GetTexts(A.Paragraph p)
        {
            StringBuilder concat = new StringBuilder();
            foreach (A.Text t in p.Descendants<A.Text>())
            {
                concat.Append(t.Text);
            }
            return concat.ToString();
        }

        /// <summary>
        /// Associates a A.Text with start and end index matching a paragraph full string (= the concatenation of all A.Text of a paragraph).
        /// </summary>
        private class TextIndex
        {
            public A.Text Text { get; private set; }
            public int StartIndex { get; private set; }
            public int EndIndex { get { return StartIndex + Text.Text.Length; } }

            public TextIndex(A.Text t, int startIndex)
            {
                this.Text = t;
                this.StartIndex = startIndex;
            }
        }

        /// <summary>
        /// Gets all the TextIndex for a given paragraph.
        /// </summary>
        private static List<TextIndex> GetTextIndexList(A.Paragraph p)
        {
            List<TextIndex> texts = new List<TextIndex>();

            StringBuilder concat = new StringBuilder();
            foreach (A.Text t in p.Descendants<A.Text>())
            {
                int startIndex = concat.Length;
                texts.Add(new TextIndex(t, startIndex));
                concat.Append(t.Text);
            }

            return texts;
        }
    }
}
