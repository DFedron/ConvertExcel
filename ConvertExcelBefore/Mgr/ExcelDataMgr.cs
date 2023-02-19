using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace ConvertExcel
{
    public sealed class ExcelDataMgr : Singleton<ExcelDataMgr>
    {
    //     private static readonly ExcelDataMgr instance = new ExcelDataMgr();
    //
    //     static ExcelDataMgr()
    //     {
    //     }
    //
    //     private ExcelDataMgr()
    //     {
    //     }
    //
    //     public static ExcelDataMgr Instance
    //     {
    //         get { return instance; }
    //     }

        private Dictionary<string, ExcelBook> m_ExcelBooksDic = new Dictionary<string, ExcelBook>();

        public void AddExcelBook(string excelName, ExcelBook excelBook)
        {
            if (m_ExcelBooksDic.ContainsKey(excelName))
            {
                m_ExcelBooksDic[excelName] = excelBook;
            }
            else
            {
                m_ExcelBooksDic.Add(excelName, excelBook);
            }
            
        }
        
        public ExcelBook GetExcelBookByExcelName(string excelName)
        {
            if (m_ExcelBooksDic.ContainsKey(excelName))
            {
                return m_ExcelBooksDic[excelName];
            }

            return null;
        }

        public Dictionary<string, ExcelBook> GetExcelBooks()
        {
            return m_ExcelBooksDic;
        }
    }
}