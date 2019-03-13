using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;

namespace AsposePoC.Converters
{
    class PowerPointConverter : IConverter
    {
        public Image Convert(string filePath)
        {
            Image img;
            using (var presentation = new Presentation(filePath))
            {
                var slide = presentation.Slides[0];
                img = slide.GetThumbnail(1f, 1f);
            }
            return img;
        }
    }
}
