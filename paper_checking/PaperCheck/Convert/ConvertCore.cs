using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace paper_checking.PaperCheck.Convert
{
    public abstract class ConvertCore
    {
        /*
         * 获取文件的文本
         */
        public abstract string ConvertToString(string path, string blockText);

        /*
         * 去除参考文献后的部分和目录前的部分
         */
        public string TextFormat(string text, string blockText)
        {
            if (text.IndexOf("参考文献") >= 0 && 1.0 * text.IndexOf("参考文献") / text.Length < 0.2)
                text = text.Substring(text.IndexOf("参考文献") + 4);
            if (text.IndexOf("引言") >= 0 && 1.0 * text.IndexOf("引言") / text.Length < 0.1)
                text = text.Substring(text.IndexOf("引言") + 2);
            else if (text.IndexOf("绪论") >= 0 && 1.0 * text.IndexOf("绪论") / text.Length < 0.1)
                text = text.Substring(text.IndexOf("绪论") + 2);
            else if (text.IndexOf("序言") >= 0 && 1.0 * text.IndexOf("序言") / text.Length < 0.1)
                text = text.Substring(text.IndexOf("序言") + 2);

            int cntckwx = 0;
            while (text.LastIndexOf("参考文献") > 0 && 1.0 * text.LastIndexOf("参考文献") / text.Length > 0.85 && cntckwx < 3)
            {
                text = text.Substring(0, text.LastIndexOf("参考文献"));
                cntckwx++;
            }
            if (blockText != null && !blockText.Trim().Equals(string.Empty))
            {
                foreach (string keyword in blockText.Split(' '))
                {
                    if(keyword.Length > 0){
                        text = text.Replace(keyword, string.Empty);
                    }
                }
            }
            text = text.Replace("（）", string.Empty).Replace("“”", string.Empty).Replace("，，", "，");
            text = text.Trim("@".ToCharArray());
            return text;
        }

    }
}
