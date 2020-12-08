using System;
using System.Text.RegularExpressions;

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
                // 这里使用了Spire Doc免费版，免费版有篇幅限制。在加载或操作Word文档时，要求Word文档不超过500个段落，25个表格。如您有更高的需求，请自行购买、升级使用付费版。
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
