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
                using (Worksheet worksheet = workbook.Worksheets[0])
                {

                    var conversionOptions = new ImageOrPrintOptions()
                    {
                        OnePagePerSheet = false,
                        ImageFormat = ImageFormat.Png,
                    };
                    
                    using (MemoryStream ms = new MemoryStream())
                    {
                        RenderPreview(worksheet, ms, conversionOptions);
                        if (ms.Length == 0)
                        {
                            FixEmptyExcel(worksheet);
                            ms.Seek(0, 0);
                            RenderPreview(worksheet, ms, conversionOptions);
                        }
                        img = new Bitmap(ms);
                    }
                }
            }
            return img;
        }


        private void RenderPreview(Worksheet worksheet, MemoryStream ms, ImageOrPrintOptions conversionOptions)
        {
            SheetRender renderer = new SheetRender(worksheet, conversionOptions);
            renderer.ToImage(0, ms);
        }

        private void FixEmptyExcel(Worksheet worksheet)
        {
            if (worksheet.Cells.MaxDataRow == -1 && worksheet.Cells.MaxDataColumn == -1)
            {
                worksheet.Cells[0, 0].Value = " ";
            }
        }


    }
}
