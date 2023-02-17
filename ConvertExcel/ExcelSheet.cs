using System.Collections.Generic;

namespace ConvertExcel
{
    public class ExcelSheet
    {
        private string m_SheetName;
        private List<ExcelColumn> m_Columns;

        public ExcelSheet(string sheetName, List<ExcelColumn> sheetColumns)
        {
            m_SheetName = sheetName;
            m_Columns = sheetColumns;
        }

        public string GetSheetName()
        {
            return m_SheetName;
        }

        public List<ExcelColumn> GetColumns()
        {
            return m_Columns;
        }
        
    }
}