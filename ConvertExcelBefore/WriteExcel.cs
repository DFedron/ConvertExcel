using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ConvertExcel
{
    public class WriteExcel
    {
        private static readonly WriteExcel instance = new WriteExcel();

        static WriteExcel()
        {
        }

        private WriteExcel()
        {
        }

        public static WriteExcel Instance
        {
            get { return instance; }
        }

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

            foreach (var excelBook in ReadExcel.Instance.GetExcelDic().Values)
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
                        excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == excelBook.GetLowerExcelName());
                    //If worksheet "Content" was not found, add it
                    if (worksheet != null)
                    {
                        excel.Workbook.Worksheets.Delete(excelBook.GetLowerExcelName());
                    }

                    worksheet = excel.Workbook.Worksheets.Add(excelBook.GetLowerExcelName());

                    //ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add(excelBook.GetExcelName());
                    int col = 1;
                    foreach (var column in sheet.GetColumns())
                    {
                        if (column.IsStruct())
                        {
                            var sheetName = $"{excelBook.GetLowerExcelName()}Struct";
                            var workStructSheet = excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == sheetName);
                            //If worksheet "Content" was not found, add it
                            if (workStructSheet != null)
                            {
                                excel.Workbook.Worksheets.Delete(sheetName);
                            }
                            workStructSheet = excel.Workbook.Worksheets.Add(sheetName);

                            // ExcelWorksheet workStructSheet =
                            //     excel.Workbook.Worksheets.Add(sheetName);
                            foreach (var firstColumn in column.GetStructColumns())
                            {
                                workStructSheet.SetValue(1, 1, "id");
                                workStructSheet.SetValue(2, 1, "ID");
                                workStructSheet.SetValue(3, 1, "number");
                                int countId = 1;
                                for (int i = 1; i <= firstColumn.GetColumnContent().Count; ++i)
                                {
                                    if (firstColumn.IfHasColumnContentByIdx(i - 1))
                                    {
                                        workStructSheet.SetValue(i + 3, 1, countId++);
                                    }
                                }

                                break;
                            }

                            int structCol = 2;
                            foreach (var structColumn in column.GetStructColumns())
                            {
                                workStructSheet.SetValue(1, structCol, structColumn.GetFieldName());
                                workStructSheet.SetValue(2, structCol, structColumn.GetAliasName());
                                if (TypeCastDic.TryGetValue(structColumn.GetDataType(), out var dataStructType))
                                {
                                    workStructSheet.SetValue(3, structCol, dataStructType);
                                }
                                else
                                {
                                    workStructSheet.SetValue(3, structCol, "string");
                                }

                                int structSheetRow = 4;
                                foreach (var rowContent in structColumn.GetColumnContent())
                                {
                                    workStructSheet.SetValue(structSheetRow, structCol, rowContent);
                                    structSheetRow++;
                                }

                                structCol++;
                            }
                        }

                        string columnFieldName = column.GetFieldName();
                        string columnAliasName = column.GetAliasName();
                        string dataType;
                        if (column.IsStruct())
                        {
                            columnFieldName =
                                $"{excelBook.GetLowerExcelName()}Struct:#{excelBook.GetLowerExcelName()}Struct.id";
                            columnAliasName = column.GetFieldName();
                            dataType = "array";
                        }
                        else
                        {
                            if (!TypeCastDic.TryGetValue(column.GetDataType(), out dataType))
                            {
                                dataType = "string";
                            }
                        }

                        worksheet.SetValue(1, col, columnFieldName);
                        worksheet.SetValue(2, col, columnAliasName);
                        worksheet.SetValue(3, col, dataType);

                        int row = 4;
                        int count = 1;
                        if (column.IsStruct())
                        {
                            var fisrtColumn = column.GetStructColumnsByIdx(0);
                            if (fisrtColumn != null)
                            {
                                foreach (var rowContent in fisrtColumn.GetColumnContent())
                                {
                                    if (rowContent.Length > 0)
                                    {
                                        worksheet.SetValue(row, col, count++);
                                    }

                                    row++;
                                }
                            }

                            // var fisrtColumn = sheet.GetColumnByIdx(0);
                            // if (fisrtColumn != null)
                            // {
                            //     foreach (var rowContent in fisrtColumn.GetColumnContent())
                            //     {
                            //         if (rowContent.Length > 0)
                            //         {
                            //             worksheet.SetValue(row, col, count);
                            //         }
                            //
                            //         count++;
                            //         row++;
                            //     }
                            // }
                        }
                        else
                        {
                            foreach (var rowContent in column.GetColumnContent())
                            {
                                worksheet.SetValue(row, col, rowContent);
                                row++;
                            }
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
    }
}