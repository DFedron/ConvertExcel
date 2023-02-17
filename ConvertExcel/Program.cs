using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;    
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing;
using System.Drawing;
using OfficeOpenXml.Drawing.Chart.Style;


namespace ConvertExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var path = "C:\\Users\\wangh\\Desktop\\Access.xlsx";
            ReadExcel.Instance.ReadFolder(path);
            //WriteExcel.Instance.WriteToExcel(existingPath, "test");
        }
    }
}
