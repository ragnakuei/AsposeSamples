using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeExcelSamples
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertPicture();
        }

        private static string _tempFile    = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "xlsx", "Sample01.xlsx");
        private static string _pictureFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pics", "1.PNG");
        private static string _outputFile  = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "xlsx", "AfterSample01.xlsx");

        private static void InsertPicture()
        {
            // 副檔名一定要正確，否則會判斷失敗

            if (File.Exists(_outputFile))
            {
                File.Delete(_outputFile);
            }

            var workBook = new Workbook(_tempFile);

            var sheet = workBook.Worksheets[0];

            var pictureIndex = sheet.Pictures.Add(3, 1, _pictureFile);

            Picture picture = sheet.Pictures[pictureIndex];
            picture.Placement   = PlacementType.MoveAndSize;
            picture.UpperDeltaX = 400;  // 只能給定 0~1024
            // picture.UpperDeltaY = 200;  // 只能給定 0~1024

            workBook.Save(_outputFile);
        }
    }
}
