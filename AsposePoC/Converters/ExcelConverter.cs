using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposePoC.Converters
{
    class ExcelConverter : IConverter
    {
        public Image Convert(string filePath)
        {
            Image img;
            using (var workbook = new Workbook(filePath))
            {
                using (var worksheet = workbook.Worksheets[0])
                {

                    var conversionOptions = new ImageOrPrintOptions()
                    {
                        OnePagePerSheet = true,
                        ImageFormat = ImageFormat.Png,
                    };
                    SheetRender renderer = new SheetRender(worksheet, conversionOptions);
                    using (MemoryStream ms = new MemoryStream())
                    {
                        renderer.ToImage(0, ms);
                        img = new Bitmap(ms);
                    }
                }
            }
            return img;
        }

    }
}
