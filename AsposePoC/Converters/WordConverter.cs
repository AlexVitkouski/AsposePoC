using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AsposePoC.Converters;
using Word = Aspose.Words;

namespace AsposePoC.Converters
{
    class WordConverter : IConverter
    {
        public Image Convert(string filePath)
        {
            var document = new Word.Document(filePath);
            Image img;
            using (MemoryStream ms = new MemoryStream())
            {
                document.Save(ms, Word.SaveFormat.Png);
                img = new Bitmap(ms);
            }
            return img;
        }
    }
}
