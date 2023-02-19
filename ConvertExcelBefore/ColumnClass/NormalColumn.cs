using System.Collections.Generic;

namespace ConvertExcel
{
    public class NormalColumn : BaseExcelColumn
    {
        public NormalColumn(List<string> columnContent,
            int colunmIndex) : base(columnContent, colunmIndex, ColumnType.NormalColumn)
        {
        }
    }
}