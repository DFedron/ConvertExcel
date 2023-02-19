using System.Collections.Generic;

namespace ConvertExcel
{

    public class BaseExcelColumn
    {
        private List<string> m_ColumnContent;
        private int m_ColumnIndex;
        private ColumnType m_ColumnType;

        public BaseExcelColumn(List<string> columnContent, int colunmIndex, ColumnType type)
        {
            m_ColumnContent = columnContent;
            m_ColumnIndex = colunmIndex;
            m_ColumnType = type;
        }

        public List<string> GetColumnContent()
        {
            return m_ColumnContent;
        }
        
        public ColumnType GetColumnType()
        {
            return m_ColumnType;
        }

        public bool IfHasColumnContentByIdx(int idx)
        {
            if (m_ColumnContent.Count > idx)
            {
                if (m_ColumnContent[idx].Length > 0)
                    return true;
            }

            return false;
        }

        public int GetColmnIndex()
        {
            return m_ColumnIndex;
        }
    }
}