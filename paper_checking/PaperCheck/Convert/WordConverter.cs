using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace paper_checking.PaperCheck.Convert
{
    class WordConverter : ConvertCore
    {
        public override string ConvertToString(string path, string blockText)
        {
            Spire.Doc.Document doc = null;
            string text = null;
            try
            {
                doc = new Spire.Doc.Document();
                doc.LoadFromFile(path);
                try
                {
                    doc.Sections[0].HeadersFooters.Header.ChildObjects.Clear();
                    doc.Sections[0].HeadersFooters.Footer.ChildObjects.Clear();
                }
                catch
                { }
                text = doc.GetText();
                text = text.Replace("#", "").Replace('\r', '#').Replace('\n', '#');
                text = Regex.Replace(text, @"[^\u4e00-\u9fa5\《\》\（\）\——\；\，\。\“\”\！\#]", "");
                text = new Regex("[#]+").Replace(text, "@@").Trim();
                text = TextFormat(text, blockText);
            }
            catch (Exception e)
            {
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                }
            }
            return text;
        }
    }
}
