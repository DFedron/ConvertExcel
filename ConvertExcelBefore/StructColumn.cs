using System.Collections.Generic;

namespace ConvertExcel
{
    public class StructColumn : BaseExcelColumn
    {
        private string m_FieldName;
        private string m_AliasName;
        private string m_DataType;
        private List<DataColumn> m_DataColumns;

        public StructColumn(string excelName, string aliasName,List<string> columnContent, int colunmIndex, List<DataColumn> dataColumns) : base(columnContent, colunmIndex, ColumnType.StructColumn)
        {
            m_DataType = "list";
            m_FieldName =  $"{GetLowerExcelName(excelName)}Struct:#{GetLowerExcelName(excelName)}Struct.id";;
            m_AliasName = aliasName;
            m_DataColumns = dataColumns;
        }
        
        public string GetLowerExcelName(string excelName)
        {
            var name = excelName.Substring(0, excelName.Length - 5);
            return name.ToLower();
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

        public List<DataColumn> GetDataColumns()
        {
            return m_DataColumns;
        }
    }
}