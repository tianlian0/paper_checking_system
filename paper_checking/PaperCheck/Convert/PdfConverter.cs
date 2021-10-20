using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.text;
using System.Text.RegularExpressions;

namespace paper_checking.PaperCheck.Convert
{
    class PdfConverter : ConvertCore
    {
        public PdfConverter()
        {
            try
            {
                org.apache.pdfbox.pdmodel.font.FontMapper fontMapper = org.apache.pdfbox.pdmodel.font.FontMappers.instance();
                java.lang.Class clazz = java.lang.Class.forName("org.apache.pdfbox.pdmodel.font.FontMapperImpl");
                java.lang.reflect.Method method = clazz.getMethod("addSubstitute", java.lang.Class.forName("java.lang.String.class"), java.lang.Class.forName("java.lang.String.class"));
                method.setAccessible(true);
                method.invoke(fontMapper, "STSong-Light", "STFangsong");
                method.invoke(fontMapper, "SimSun", "STFangsong");
                method.invoke(fontMapper, "AdobeKaitiStd-Regular", "STFangsong");
                method.invoke(fontMapper, "ArialMT", "Arial");
                method.invoke(fontMapper, "cid0ct", "STFangsong");
                method.invoke(fontMapper, "CMEX", "STFangsong");
                method.invoke(fontMapper, "TCBLZV", "STFangsong");
                method.invoke(fontMapper, "LCIRCLE10", "STFangsong");
            }
            catch (System.Exception)
            {
            }
        }

        public override string ConvertToString(string path, string blockText)
        {
            PDDocument pdffile = null;
            string text = null;
            try
            {
                pdffile = org.apache.pdfbox.Loader.loadPDF(new java.io.File(path));
                PDFTextStripper pdfStripper = new PDFTextStripper();
                text = pdfStripper.getText(pdffile);
                text = Regex.Replace(text, @"[^\u4e00-\u9fa5\《\》\（\）\——\；\，\。\“\”\！]", "");
                text = Regex.Replace(text, @"\s", string.Empty);
                text = TextFormat(text, blockText);
            }
            finally
            {
                if (pdffile != null)
                {
                    pdffile.close();
                }
            }
            return text;
        }
    }
}
