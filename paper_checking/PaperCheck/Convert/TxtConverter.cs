using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace paper_checking.PaperCheck.Convert
{
    class TxtConverter : ConvertCore
    {
        public override string ConvertToString(string path)
        {
            string text = File.ReadAllText(path, Encoding.GetEncoding("GBK"));
            text = text.Replace("#", "").Replace('\r', '#').Replace('\n', '#');
            text = Regex.Replace(text, @"[^\u4e00-\u9fa5\《\》\（\）\——\；\，\。\“\”\！\#]", "");
            text = new Regex("[#]+").Replace(text, "@@").Trim();
            text = TextFormat(text);
            return text;
        }
    }
}
