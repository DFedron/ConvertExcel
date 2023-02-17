using System.Collections.Generic;

namespace ConvertExcel
{
    public class ExcelColumn
    {
        private bool m_IsStructColumn;
        private string m_FieldName;     
        private string m_AliasName;     
        private string m_DataType;
        private List<string> m_ColumnContent;
        private int m_ColumnIndex;
        private int m_ColumnStructStartIndex;
        private int m_ColumnStructEndIndex;
        private List<ExcelColumn> m_StructColumns;
        
        public ExcelColumn(string fieldName, string aliasName,string dataType, List<string> sheetColumn, int colunmIndex)
        {
            m_DataType = dataType;
            m_IsStructColumn = false;
            m_FieldName = fieldName;
            m_AliasName = aliasName;
            m_ColumnContent = sheetColumn;
            m_ColumnIndex = colunmIndex;
        }
        
        public ExcelColumn(string fieldName, List<ExcelColumn> structColumns, int columnStructStartIndex, int columnStructEndIndex)
        {
            m_IsStructColumn = true;
            m_FieldName = fieldName;
            m_StructColumns = structColumns;
            m_ColumnStructStartIndex = columnStructStartIndex;
            m_ColumnStructEndIndex = columnStructEndIndex;
        }

        public bool IsStruct()
        {
            return m_IsStructColumn;
        }

        public string GetFieldName()
        {
            return m_FieldName;
        }
        public string GetAliasName()
        {
            return m_AliasName;
        }
        public string GetDataType()
        {
            return m_DataType;
        }

        public List<string> GetColumnContent()
        {
            return m_ColumnContent;
        }

        public int GetColmnIndex()
        {
            return m_ColumnIndex;
        }
    }
}