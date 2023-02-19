using System.Collections.Generic;

namespace ConvertExcel
{
    public class ExcelStructSheet
    {
        private string m_StructName;
        private Dictionary<string, string> m_StructDic = new Dictionary<string, string>();

        public ExcelStructSheet(string structName, Dictionary<string, string> structDic)
        {
            m_StructName = structName;
            m_StructDic = structDic;
        }

        public string GetFieldName()
        {
            return m_StructName;
        }

        public Dictionary<string, string> GetStructDic()
        {
            return m_StructDic;
        }
    }
}