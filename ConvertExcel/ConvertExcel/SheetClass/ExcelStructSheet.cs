using System.Collections.Generic;

namespace ConvertExcel
{
    public class ExcelStructSheet : BaseExcelSheet
    {
        private string m_StructName;
        private Dictionary<string, Dictionary<string, string>> m_StructDic;

        public ExcelStructSheet(string structName, Dictionary<string, Dictionary<string, string>> structDic,
            List<BaseExcelColumn> sheetColumns) : base(structName, sheetColumns, ExcelSheetType.StructSheet)
        {
            m_StructName = structName;
            m_StructDic = structDic;
        }

        public Dictionary<string, string> GetStructByName(string name)
        {
            if (m_StructDic.ContainsKey(name))
                return m_StructDic[name];
            return null;
        }
    }
}