using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace ConvertExcel
{
    public class WriteExcel
    {
        private static readonly WriteExcel instance = new WriteExcel();

        static WriteExcel()
        {
        }

        private WriteExcel()
        {
        }

        public static WriteExcel Instance
        {
            get { return instance; }
        }

        private List<string> m_ErrorMsg = new List<string>();

        private Dictionary<string, string> TypeCastDic = new Dictionary<string, string>
        {
            { "int", "num" },
            { "string", "string" },
            { "bool", "bool" },
            { "list", "array" },
            { "array", "array" },
        };

        public void WriteToExcel(string path, ExcelBook excelBook)
        {
            ExcelPackage excel = null;
            Stream stream = null;
            try
            {
                excel = new ExcelPackage();
                foreach (var sheet in excelBook.GetSheets())
                {
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add(sheet.GetSheetName());
                    int col = 1;
                    foreach (var column in sheet.GetColumns())
                    {
                        worksheet.SetValue(1, col, column.GetFieldName());
                        worksheet.SetValue(2, col, column.GetAliasName());
                        worksheet.SetValue(3, col, column.GetDataType());
                        int row = 4;
                        foreach (var rowContent in column.GetColumnContent())
                        {
                            worksheet.SetValue(row, col, rowContent);
                            row++;
                        }
                        ++col  ;
                    }
                }

                //excel.Save();
                stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
                excel.SaveAs(stream);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            finally
            {
                if (excel != null)
                    excel.Dispose();
                if (stream != null)
                    stream.Dispose();
            }
        }

        public void PrintErrorMsg(string msg)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(msg);
            Console.ResetColor();
        }

        public void PrintSuccessMsg()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("写入成功");
            Console.ResetColor();
        }

        private void AddErrorMsg(string msg)
        {
            if (!m_ErrorMsg.Contains(msg))
            {
                m_ErrorMsg.Add(msg);
            }
        }
    }
}