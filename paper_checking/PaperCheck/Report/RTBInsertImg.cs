using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Specialized;
using System.Drawing.Imaging;
using System.IO;

namespace RTF_Operation
{

    public class RTB_InsertImg
    {
        public const int
        EmfToWmfBitsFlagsDefault = 0x00000000,
        EmfToWmfBitsFlagsEmbedEmf = 0x00000001,
        EmfToWmfBitsFlagsIncludePlaceable = 0x00000002,
        EmfToWmfBitsFlagsNoXORClip = 0x00000004;

        [DllImport("gdiplus.dll")]
        public static extern uint GdipEmfToWmfBits(IntPtr _hEmf, uint _bufferSize, byte[] _buffer, int _mappingMode, int _flags);

        private struct RtfFontFamilyDef
        {
            public const string Unknown = @"\fnil";
            public const string Roman = @"\froman";
            public const string Swiss = @"\fswiss";
            public const string Modern = @"\fmodern";
            public const string Script = @"\fscript";
            public const string Decor = @"\fdecor";
            public const string Technical = @"\ftech";
            public const string BiDirect = @"\fbidi";
        }

        private const int MM_ISOTROPIC = 7;
        private const int MM_ANISOTROPIC = 8;
        private const int HMM_PER_INCH = 2540;
        private const int TWIPS_PER_INCH = 1440;

        private const string FF_UNKNOWN = "UNKNOWN";

        private const string RTF_HEADER = @"{\rtf1\ansi\ansicpg1252\deff0\deflang1033";
        private const string RTF_DOCUMENT_PRE = @"\viewkind4\uc1\pard\cf1\f0\fs20";
        private const string RTF_DOCUMENT_POST = @"\cf0\fs17}";
        private const string RTF_IMAGE_POST = @"}";

        private static HybridDictionary rtfFontFamily;

        static RTB_InsertImg()
        {
            rtfFontFamily = new HybridDictionary();
            rtfFontFamily.Add(FontFamily.GenericMonospace.Name, RtfFontFamilyDef.Modern);
            rtfFontFamily.Add(FontFamily.GenericSansSerif.Name, RtfFontFamilyDef.Swiss);
            rtfFontFamily.Add(FontFamily.GenericSerif.Name, RtfFontFamilyDef.Roman);
            rtfFontFamily.Add(FF_UNKNOWN, RtfFontFamilyDef.Unknown);
        }

        private static string GetFontTable(Font _font)
        {
            StringBuilder _fontTable = new StringBuilder();
            _fontTable.Append(@"{\fonttbl{\f0");
            _fontTable.Append(@"\");
            if (rtfFontFamily.Contains(_font.FontFamily.Name))
                _fontTable.Append(rtfFontFamily[_font.FontFamily.Name]);
            else
                _fontTable.Append(rtfFontFamily[FF_UNKNOWN]);
            _fontTable.Append(@"\fcharset0 ");
            _fontTable.Append(_font.Name);
            _fontTable.Append(@";}}");
            return _fontTable.ToString();
        }

        /// 
              /// 在RichTextBox当前光标处插入一副图像。 
              /// 
              /// 多格式文本框控件 
              /// 插入的图像 
        public static void InsertImage(RichTextBox rtb, Image image)
        {
            StringBuilder _rtf = new StringBuilder();
            _rtf.Append(RTF_HEADER);
            _rtf.Append(GetFontTable(rtb.Font));
            _rtf.Append(GetImagePrefix(rtb, image));
            _rtf.Append(GetRtfImage(rtb, image));
            _rtf.Append(RTF_IMAGE_POST);
            rtb.SelectedRtf = _rtf.ToString();
        }

        private static string GetImagePrefix(RichTextBox rtb, Image _image)
        {
            float xDpi;
            float yDpi;
            using (Graphics _graphics = rtb.CreateGraphics())
            {
                xDpi = _graphics.DpiX;
                yDpi = _graphics.DpiY;
            }

            StringBuilder _rtf = new StringBuilder();
            int picw = (int)Math.Round((_image.Width / xDpi) * HMM_PER_INCH);
            int pich = (int)Math.Round((_image.Height / yDpi) * HMM_PER_INCH);
            int picwgoal = (int)Math.Round((_image.Width / xDpi) * TWIPS_PER_INCH);
            int pichgoal = (int)Math.Round((_image.Height / yDpi) * TWIPS_PER_INCH);
            _rtf.Append(@"{\pict\wmetafile8");
            _rtf.Append(@"\picw");
            _rtf.Append(picw);
            _rtf.Append(@"\pich");
            _rtf.Append(pich);
            _rtf.Append(@"\picwgoal");
            _rtf.Append(picwgoal);
            _rtf.Append(@"\pichgoal");
            _rtf.Append(pichgoal);
            _rtf.Append(" ");
            return _rtf.ToString();
        }

        private static string GetRtfImage(RichTextBox rtb, Image _image)
        {
            StringBuilder _rtf = null;
            MemoryStream _stream = null;
            Graphics _graphics = null;
            Metafile _metaFile = null;
            IntPtr _hdc;
            try
            {
                _rtf = new StringBuilder();
                _stream = new MemoryStream();
                using (_graphics = rtb.CreateGraphics())
                {
                    _hdc = _graphics.GetHdc();
                    _metaFile = new Metafile(_stream, _hdc);
                    _graphics.ReleaseHdc(_hdc);
                }
                using (_graphics = Graphics.FromImage(_metaFile))
                {
                    _graphics.DrawImage(_image, new Rectangle(0, 0, _image.Width, _image.Height));
                }
                IntPtr _hEmf = _metaFile.GetHenhmetafile();
                uint _bufferSize = GdipEmfToWmfBits(_hEmf, 0, null, MM_ANISOTROPIC, EmfToWmfBitsFlagsDefault);
                byte[] _buffer = new byte[_bufferSize];
                uint _convertedSize = GdipEmfToWmfBits(_hEmf, _bufferSize, _buffer, MM_ANISOTROPIC, EmfToWmfBitsFlagsDefault);
                for (int i = 0; i < _buffer.Length; ++i)
                {
                    _rtf.Append(String.Format("{0:X2}", _buffer[i]));
                }

                return _rtf.ToString();
            }
            finally
            {
                if (_graphics != null)
                    _graphics.Dispose();
                if (_metaFile != null)
                    _metaFile.Dispose();
                if (_stream != null)
                    _stream.Close();
            }
        }
    }
}