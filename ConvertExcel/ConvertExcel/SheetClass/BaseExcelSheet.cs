using System.Collections.Generic;

namespace ConvertExcel
{
 
    public class BaseExcelSheet
    {
        private string m_SheetName;
        private List<BaseExcelColumn> m_Columns;
        private ExcelSheetType m_ExcelSheetType;

        public BaseExcelSheet(string sheetName, List<BaseExcelColumn> sheetColumns, ExcelSheetType type)
        {
            m_SheetName = sheetName;
            m_Columns = sheetColumns;
            m_ExcelSheetType = type;
        }

        public string GetSheetName()
        {
            return m_SheetName;
        }

        public List<BaseExcelColumn> GetColumns()
        {
            return m_Columns;
        }
        
        public BaseExcelColumn GetPrimeColumn()
        {
            if (m_Columns.Count > 0)
                return m_Columns[0];
            return null;
        }
        
        
        
    }
}