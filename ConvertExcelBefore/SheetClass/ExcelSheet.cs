using System.Collections.Generic;

namespace ConvertExcel
{
    public class ExcelSheet : BaseExcelSheet
    {
        public ExcelSheet(string sheetName, List<BaseExcelColumn> sheetColumns):base(sheetName, sheetColumns, ExcelSheetType.NormalSheet)
        {

        }
    }
}