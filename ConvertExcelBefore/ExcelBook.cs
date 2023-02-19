using System.Collections.Generic;

namespace ConvertExcel
{
    public class ExcelBook
    {
        private string m_ExcelName;     
        private List<ExcelSheet> m_Sheets;
        
        public ExcelBook(string excelName, List<ExcelSheet> sheets)
        {
            m_ExcelName = excelName;
            m_Sheets = sheets;
        }


        public List<ExcelSheet> GetSheets()
        {
            return m_Sheets;
        }

        public string GetExcelName()
        {
            var name = m_ExcelName.Substring(0, m_ExcelName.Length - 5);
            return name;
        }
        
        public string GetLowerExcelName()
        {
            var name = m_ExcelName.Substring(0, m_ExcelName.Length - 5);
            return name.ToLower();
        }

    }
}