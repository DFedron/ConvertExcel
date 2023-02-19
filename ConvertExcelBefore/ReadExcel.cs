using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace ConvertExcel
{
    public sealed class ReadExcel : Singleton<ReadExcel>
    {
        public void ReadFolder(string folderPath)
        {
            ErrorMsgMgr.Instance.ClearErrorMsg();
            DirectoryInfo folder = new DirectoryInfo(folderPath);
            ReadAllExcelData(folder);
        }

        private void ReadAllExcelData(DirectoryInfo folder)
        {
            if (!folder.Exists)
            {
                ErrorMsgMgr.Instance.AddErrorMsg($"{folder.Name}文件夹不存在");
                return;
            }

            foreach (DirectoryInfo subFolder in folder.GetDirectories())
            {
                if (subFolder.Name.Contains("GeneratedYamato") || subFolder.Name.Contains("ConvertExcel"))
                    continue;
                ReadAllExcelData(subFolder);
            }

            foreach (FileInfo file in folder.GetFiles("*.xlsx"))
            {
                if (file.Name.Contains("~$") || file.Name.Contains("Yamato"))
                {
                    continue;
                }

                ReadOneExcel(file, folder.FullName);
            }
        }

        public void ReadOneExcel(FileInfo file, string folderPath)
        {
            FileStream fileStream;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excel;
            try
            {
                fileStream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                excel = new ExcelPackage(fileStream);
                List<BaseExcelSheet> sheets = new List<BaseExcelSheet>();

                if (excel.Workbook.Worksheets.Count == 0)
                {
                    ErrorMsgMgr.Instance.AddErrorMsg($"{file.Name}表里的sheet数量为0");
                    return;
                }

                int count = 1;
                foreach (var sheet in excel.Workbook.Worksheets)
                {
                    ExcelSheet excelSheet;
                    if (count == 1)
                        excelSheet = ReadDataSheet(sheet, file.Name);
                    else
                        excelSheet = ReadOtherSheet(sheet);
                    ++count;
                    if (excelSheet != null)
                        sheets.Add(excelSheet);
                }

                ExcelDataMgr.Instance.AddExcelBook(file.Name, new ExcelBook(file.Name, sheets));
                excel.Dispose();
                fileStream.Dispose();
            }
            catch (Exception e)
            {
                ErrorMsgMgr.Instance.AddErrorMsg($"读取{file.Name}出错 \n详细信息:\n{e}\n\n");
            }
        }

        private ExcelSheet ReadOtherSheet(ExcelWorksheet otherSheet)
        {
            if (otherSheet == null) return null;
            List<BaseExcelColumn> sheetColumns = new List<BaseExcelColumn>();
            try
            {
                int maxColumnNum = otherSheet.Dimension.End.Column;
                int maxRowNum = otherSheet.Dimension.End.Row;
                for (var col = 1; col <= maxColumnNum; col++)
                {
                    List<string> columnContent = new List<string>();
                    for (int row = 1; row <= maxRowNum; row++)
                    {
                        columnContent.Add(GetValue(otherSheet, row, col));
                    }

                    sheetColumns.Add(new NormalColumn(columnContent, col));
                }
            }
            catch
            {
                return null;
            }


            return new ExcelSheet(otherSheet.Name, sheetColumns);
        }

        private ExcelSheet ReadDataSheet(ExcelWorksheet dataSheet, string excelName)
        {
            if (dataSheet == null) return null;
            List<BaseExcelColumn> sheetColumns = new List<BaseExcelColumn>();
            int maxColumnNum;
            int maxRowNum;
            try
            {
                maxColumnNum = dataSheet.Dimension.End.Column;
                maxRowNum = dataSheet.Dimension.End.Row;
            }
            catch
            {
                return null;
            }

            for (var col = 2; col <= maxColumnNum; col++)
            {
                //一列column需要的数据
                List<string> columnContent = new List<string>();
                string fieldName = "";
                string aliasName = "";
                string dataType = "";

                StructColumn structColumn = null;
                for (var row = 1; row <= maxRowNum; row++)
                {
                    switch (GetValue(dataSheet, row, 1))
                    {
                        case "##comment":
                            aliasName = GetValue(dataSheet, row, col);
                            break;
                        case "##var":
                            if (fieldName.Length == 0)
                                fieldName = GetValue(dataSheet, row, col);
                            if (fieldName.Contains("*"))
                            {
                                structColumn = GetStructColumn(dataSheet, ref col, row, excelName);
                            }

                            break;
                        case "##type":
                            dataType = GetValue(dataSheet, row, col);
                            break;
                        case "##":
                            break;
                        case "":
                            columnContent.Add(GetValue(dataSheet, row, col));
                            break;
                    }

                    if (structColumn != null) break;
                }

                if (structColumn != null)
                {
                    sheetColumns.Add(structColumn);
                }
                else
                {
                    sheetColumns.Add(new DataColumn(fieldName, aliasName, dataType, columnContent, col));
                }
            }

            return new ExcelSheet(dataSheet.Name, sheetColumns);
        }

        private string GetValue(ExcelWorksheet firstWorksheet, int row, int col)
        {
            if (firstWorksheet.Cells[row, col].Value == null)
            {
                return "";
            }
            else
            {
                return firstWorksheet.Cells[row, col].Value.ToString();
            }
        }


        private StructColumn GetStructColumn(ExcelWorksheet curWorksheet, ref int col, int row, string excelName)
        {
            int startCol = col;
            List<DataColumn> structDataColumns = new List<DataColumn>();
            List<string> structColumnContent = new List<string>();
            string structFieldName = curWorksheet.Cells[row, col].Value.ToString();
            bool first = true;
            do
            {
                string fieldName = "";
                string aliasName = "";
                string dataType = "";
                List<string> columnContent = new List<string>();

                for (int structRow = 1; structRow <= curWorksheet.Dimension.End.Row; ++structRow)
                {
                    if (structRow == row) continue;
                    switch (GetValue(curWorksheet, structRow, 1))
                    {
                        case "##comment":
                            aliasName = GetValue(curWorksheet, structRow, col);
                            break;
                        case "##var":
                            fieldName = GetValue(curWorksheet, structRow, col);
                            break;
                        case "##type":
                            dataType = GetValue(curWorksheet, structRow, col);
                            break;
                        case "##":
                            break;
                        case "":
                            columnContent.Add(GetValue(curWorksheet, structRow, col));
                            if (first)
                            {
                                structColumnContent.Add((columnContent.Count).ToString());
                            }

                            break;
                    }
                }

                structDataColumns.Add(new DataColumn(fieldName, aliasName, dataType, columnContent, col));
                col++;
                first = false;
                if (col > curWorksheet.Dimension.End.Column) break;
            } while (GetValue(curWorksheet, row, col).Length == 0 &&
                     GetValue(curWorksheet, row + 1, col).Length > 0);

            col--;
            return new StructColumn(excelName, structFieldName, structColumnContent, startCol, structDataColumns);
        }

        // private Dictionary<string, ExcelStructSheet> GetAllStruct(ExcelWorksheet structSheet)
        // {
        //     Dictionary<string, ExcelStructSheet> allStructDic = new Dictionary<string, ExcelStructSheet>();
        //     int maxColumnNum = structSheet.Dimension.End.Column;
        //     int maxRowNum = structSheet.Dimension.End.Row;
        //
        //     int fullNameIndex = 0;
        //     int fieldNameIndex = 0;
        //     int typeIndex = 0;
        //     for (var col = 1; col <= maxColumnNum; col++)
        //     {
        //         switch (GetValue(structSheet, 1, col))
        //         {
        //             case "full_name":
        //                 fullNameIndex = col;
        //                 break;
        //             case "*fields":
        //                 fieldNameIndex = col;
        //                 typeIndex = col + 1;
        //                 break;
        //             default:
        //                 break;
        //         }
        //
        //         if (fullNameIndex != 0 && fullNameIndex != 0 && typeIndex != 0)
        //             break;
        //     }
        //
        //     for (var row = 1; row <= maxRowNum; row++)
        //     {
        //         string fullName = "";
        //         switch (GetValue(structSheet, row, 1))
        //         {
        //             case "":
        //                 fullName = structSheet.Cells[row, fullNameIndex].Value.ToString();
        //                 Dictionary<string, string> structDic = new Dictionary<string, string>();
        //                 do
        //                 {
        //                     structDic.Add(structSheet.Cells[row, fieldNameIndex].Value.ToString(),
        //                         structSheet.Cells[row, typeIndex].Value.ToString());
        //                     row++;
        //                 } while (structSheet.Cells[row, fullNameIndex].Value.ToString().Length == 0 &&
        //                          structSheet.Cells[row, fieldNameIndex].Value.ToString().Length > 0);
        //
        //                 allStructDic.Add(fullName, new ExcelStructSheet(fullName, structDic));
        //                 break;
        //             default:
        //                 break;
        //         }
        //     }
        //
        //     return allStructDic;
        // }

        // private int GetTypeIndex(ExcelWorksheet curWorksheet)
        // {
        //     for (int col = 1; col <= curWorksheet.Dimension.End.Column; ++col)
        //     {
        //         if (curWorksheet.Cells[1, col].Value.ToString() == "##type")
        //         {
        //             return col;
        //         }
        //     }
        //     return -1;
        // }
    }
}