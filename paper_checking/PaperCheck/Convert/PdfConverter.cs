using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.text;
using System.Text.RegularExpressions;

namespace paper_checking.PaperCheck.Convert
{
    class PdfConverter : ConvertCore
    {
        public override string ConvertToString(string path, string blockText)
        {
            PDDocument pdffile = null;
            string text = null;
            try
            {
                pdffile = PDDocument.load(new java.io.File(path));
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
