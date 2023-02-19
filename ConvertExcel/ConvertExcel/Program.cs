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
            var folderPath = @"D:\Work\ConvertExcel\Test";

            ReadExcel.Instance.ReadFolder(folderPath);
            WriteExcel.Instance.WriteToFolder(folderPath);
            ErrorMsgMgr.Instance.AutoPrintErrorOrSucces();
        }
    }
}
