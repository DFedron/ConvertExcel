using System.Collections.Generic;

namespace ConvertExcel
{
    public class BaseExcelSheet
    {
        private string m_SheetName;
        private List<BaseExcelColumn> m_Columns;

        public BaseExcelSheet(string sheetName, List<BaseExcelColumn> sheetColumns)
        {
            m_SheetName = sheetName;
            m_Columns = sheetColumns;
        }

        public string GetSheetName()
        {
            return m_SheetName;
        }

        public List<BaseExcelColumn> GetColumns()
        {
            return m_Columns;
        }
        
        public BaseExcelColumn GetColumnByIdx(int idx)
        {
            if (m_Columns.Count > idx)
                return m_Columns[idx];
            return null;
        }
        
    }
}