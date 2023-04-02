using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;    
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing;
using System.Drawing;
using System.Threading.Tasks;
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
        // static async Task Main(string[] args)
        // {
        //     Console.WriteLine("Starting...");
        //
        //     await Task.Delay(1000); // 模拟耗时操作1s
        //
        //     Console.WriteLine("Waiting...");
        //
        //     await Task.Delay(2000); // 模拟耗时操作2s
        //
        //     Console.WriteLine("Done!");
        // }
        
    }
}
