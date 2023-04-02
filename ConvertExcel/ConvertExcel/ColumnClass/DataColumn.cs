using System.Collections.Generic;

namespace ConvertExcel
{
    public class DataColumn : BaseExcelColumn
    {
        private string m_FieldName;
        private string m_AliasName;
        private string m_DataType;

        public DataColumn(string fieldName, string aliasName, string dataType, List<string> columnContent,
            int colunmIndex) : base(columnContent, colunmIndex, ColumnType.DataColumn)
        {
            m_DataType = dataType;
            if (fieldName.ToLower() == "tag")
                fieldName = "id";
            m_FieldName = fieldName;
            m_AliasName = aliasName;
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
    }
}