using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace paper_checking.PaperCheck.Convert
{
    public class ConverterFactory
    {

        private static readonly WordConverter _wordConverter = new WordConverter();
        private static readonly PdfConverter _pdfConverter = new PdfConverter();
        private static readonly TxtConverter _txtConverter = new TxtConverter();

        public ConverterFactory()
        {
        }

        public ConvertCore GetConverter(string type, RunningEnv runningEnv)
        {
            type = type.ToLower();
            //判断支持该格式
            if ("doc".Equals(type) && !runningEnv.SettingData.SuportDoc ||
                "docx".Equals(type) && !runningEnv.SettingData.SuportDocx ||
                "pdf".Equals(type) && !runningEnv.SettingData.SuportPdf ||
                "txt".Equals(type) && !runningEnv.SettingData.SuportTxt)
            {
                return null;
            }
            //获取对应的转换器
            switch (type)
            {
                case "doc":
                case "docx":
                    return _wordConverter;
                case "pdf":
                    return _pdfConverter;
                case "txt":
                    return _txtConverter;
                default:
                    return null;
            }
        }

    }
}
