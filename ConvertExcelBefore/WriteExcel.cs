using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ConvertExcel
{
    public sealed class WriteExcel : Singleton<WriteExcel>
    {
        private List<string> m_ErrorMsg = new List<string>();

        private Dictionary<string, string> TypeCastDic = new Dictionary<string, string>
        {
            { "int", "number" },
            { "string", "string" },
            { "bool", "bool" },
            { "list", "array" },
            { "array", "array" },
        };

        public void WriteToFolder(string folderPath)
        {
            folderPath = $"{folderPath}\\GeneratedYamato";
            DirectoryInfo folder = new DirectoryInfo(folderPath);
            if (!folder.Exists)
            {
                Directory.CreateDirectory(folderPath);
            }

            foreach (var excelBook in ExcelDataMgr.Instance.GetExcelBooks().Values)
            {
                WriteToExcel(folderPath, excelBook);
            }
        }

        private void WriteToExcel(string path, ExcelBook excelBook)
        {
            ExcelPackage excel = null;
            Stream stream = null;
            try
            {
                excel = new ExcelPackage($"{path}\\{excelBook.GetExcelName()}Yamato.xlsx");
                foreach (var sheet in excelBook.GetSheets())
                {
                    var worksheet =
                        excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == sheet.GetSheetName());
                    //If worksheet "Content" was not found, add it
                    if (worksheet != null)
                    {
                        excel.Workbook.Worksheets.Delete(sheet.GetSheetName());
                    }

                    worksheet = excel.Workbook.Worksheets.Add(sheet.GetSheetName());

                    //ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add(excelBook.GetExcelName());
                    int col = 1;
                    foreach (var column in sheet.GetColumns())
                    {
                        switch (column.GetColumnType())
                        {
                            case ColumnType.DataColumn:
                                AppendDataColumn(worksheet, column as DataColumn, col);
                                break;
                            case ColumnType.NormalColumn:
                                AppendNormalColumn(worksheet, column as NormalColumn, col);
                                break;
                            case ColumnType.StructColumn:
                                var sheetName = $"{excelBook.GetLowerExcelName()}Struct";
                                var workStructSheet =
                                    excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == sheetName);
                                //If worksheet "Content" was not found, add it
                                if (workStructSheet != null)
                                {
                                    excel.Workbook.Worksheets.Delete(sheetName);
                                }

                                workStructSheet = excel.Workbook.Worksheets.Add(sheetName);
                                var structColumn = column as StructColumn;
                                WriteToStructSheet(workStructSheet, structColumn);
                                AppendDataColumn(worksheet,
                                    new DataColumn(structColumn.GetFieldName(), structColumn.GetAliasName(),
                                        structColumn.GetDataType(), structColumn.GetColumnContent(),
                                        structColumn.GetColmnIndex()), col);

                                break;
                            default:
                                break;
                        }

                        ++col;
                    }
                }

                //excel.Save();
                stream = new FileStream($"{path}\\{excelBook.GetExcelName()}Yamato.xlsx", FileMode.Create,
                    FileAccess.Write, FileShare.ReadWrite);
                excel.SaveAs(stream);
            }
            catch (Exception e)
            {
                ErrorMsgMgr.Instance.AddErrorMsg(
                    $"{excelBook.GetExcelName()}写入出错, 检查是否Excel是否在另外的进程中打开\n 详细信息:\n{e}\n\n");
            }
            finally
            {
                if (excel != null)
                    excel.Dispose();
                if (stream != null)
                    stream.Dispose();
            }
        }

        private void AppendNormalColumn(ExcelWorksheet worksheet, NormalColumn column, int col)
        {
            int row = 1;
            foreach (var rowContent in column.GetColumnContent())
            {
                worksheet.SetValue(row, col, rowContent);
                row++;
            }
        }

        private void AppendDataColumn(ExcelWorksheet worksheet, DataColumn column, int col)
        {
            string columnFieldName = column.GetFieldName();
            string columnAliasName = column.GetAliasName();
            string dataType;
            if (!TypeCastDic.TryGetValue(column.GetDataType(), out dataType))
            {
                dataType = "string";
            }

            worksheet.SetValue(1, col, columnFieldName);
            worksheet.SetValue(2, col, columnAliasName);
            worksheet.SetValue(3, col, dataType);

            int row = 4;
            foreach (var rowContent in column.GetColumnContent())
            {
                worksheet.SetValue(row, col, rowContent);
                row++;
            }
        }

        private void WriteToStructSheet(ExcelWorksheet workStructSheet, StructColumn structColumn)
        {
            workStructSheet.SetValue(1, 1, "id");
            workStructSheet.SetValue(2, 1, "ID");
            workStructSheet.SetValue(3, 1, "number");
            int structRow = 4;
            foreach (var structColumnContent in structColumn.GetColumnContent())
            {
                workStructSheet.SetValue(structRow++, 1, structColumnContent);
            }


            int structCol = 2;
            foreach (var dataColumn in structColumn.GetDataColumns())
            {
                workStructSheet.SetValue(1, structCol, dataColumn.GetFieldName());
                workStructSheet.SetValue(2, structCol, dataColumn.GetAliasName());
                if (TypeCastDic.TryGetValue(dataColumn.GetDataType(), out var dataStructType))
                {
                    workStructSheet.SetValue(3, structCol, dataStructType);
                }
                else
                {
                    workStructSheet.SetValue(3, structCol, "string");
                }

                int structSheetRow = 4;
                foreach (var rowContent in dataColumn.GetColumnContent())
                {
                    workStructSheet.SetValue(structSheetRow, structCol, rowContent);
                    structSheetRow++;
                }

                structCol++;
            }
        }
    }
}