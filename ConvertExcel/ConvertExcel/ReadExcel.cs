using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ConvertExcel
{
    public class ReadExcel
    {
        private static readonly ReadExcel instance = new ReadExcel();

        static ReadExcel()
        {
        }

        private ReadExcel()
        {
        }

        public static ReadExcel Instance
        {
            get { return instance; }
        }


        private List<string> m_ErrorMsg = new List<string>();
        private Dictionary<string, ExcelBook> m_ExcelBooksDic = new Dictionary<string, ExcelBook>();

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
                if(subFolder.Name.Contains("GeneratedYamato") || subFolder.Name.Contains("ConvertExcel"))
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
            if (file.Name.Contains("itemresource"))
            {
                Console.WriteLine();
            }
            FileStream fileStream;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excel;
            try
            {
                fileStream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                excel = new ExcelPackage(fileStream);
                List<ExcelSheet> sheets = new List<ExcelSheet>();

                if (excel.Workbook.Worksheets.Count == 0)
                {
                    ErrorMsgMgr.Instance.AddErrorMsg($"{file.Name}表里的sheet数量为0");
                    return;
                }

                var firstSheet = ReadFirstSheet(excel.Workbook.Worksheets[0]);
                sheets.Add(firstSheet);
                m_ExcelBooksDic.Add(file.Name, new ExcelBook(file.Name, sheets));
                //WriteExcel.Instance.WriteToExcel($"{folderPath}", m_ExcelBooksDic[file.Name]);
                excel.Dispose();
                fileStream.Dispose();
            }
            catch (Exception e)
            {
                ErrorMsgMgr.Instance.AddErrorMsg($"读取{file.Name}出错 \n详细信息:\n{e}\n\n");
            }
        }

        private ExcelSheet ReadFirstSheet(ExcelWorksheet firstWorksheet)
        {
            List<ExcelColumn> sheetColumns = new List<ExcelColumn>();
            int maxColumnNum = firstWorksheet.Dimension.End.Column;
            int maxRowNum = firstWorksheet.Dimension.End.Row;

            for (var col = 2; col <= maxColumnNum; col++)
            {
                //一列column需要的数据
                List<string> columnContent = new List<string>();
                string fieldName = "";
                string aliasName = "";
                string dataType = "";
                bool isStruct = false;

                ExcelColumn excelStructColumn = null;
                for (var row = 1; row <= maxRowNum; row++)
                {
                    switch (GetValue(firstWorksheet,row, 1))
                    {
                        case "##comment":
                            aliasName = GetValue(firstWorksheet, row, col);
                            break;
                        case "##var":
                            if (fieldName.Length == 0)
                                fieldName = GetValue(firstWorksheet, row, col);
                            if (fieldName.Contains("*"))
                            {
                                isStruct = true;
                                excelStructColumn = GetStructColumn(firstWorksheet, ref col, row);
                            }

                            break;
                        case "##type":
                            dataType = GetValue(firstWorksheet, row, col);
                            break;
                        case "##":
                            break;
                        case "":
                            columnContent.Add(GetValue(firstWorksheet, row, col));
                            break;
                        default:
                            break;
                    }

                    if (isStruct) break;
                }

                if (isStruct && excelStructColumn != null)
                {
                    sheetColumns.Add(excelStructColumn);
                }
                else
                {
                    sheetColumns.Add(new ExcelColumn(fieldName, aliasName, dataType, columnContent, col));
                }
            }

            return new ExcelSheet(firstWorksheet.Name, sheetColumns);
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


        private ExcelColumn GetStructColumn(ExcelWorksheet curWorksheet, ref int col, int row)
        {
            int startIndex = col;
            List<ExcelColumn> structColumns = new List<ExcelColumn>();
            string structFieldName = curWorksheet.Cells[row, col].Value.ToString();
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
                            break;
                        default:
                            break;
                    }
                }

                structColumns.Add(new ExcelColumn(fieldName, aliasName, dataType, columnContent, col));
                col++;
                if (col > curWorksheet.Dimension.End.Column) break;
            } while (GetValue(curWorksheet, row, col).Length == 0 &&
                     GetValue(curWorksheet, row + 1, col).Length > 0);

            return new ExcelColumn(structFieldName, structColumns, startIndex, --col);
        }

        private Dictionary<string, ExcelStructSheet> GetAllStruct(ExcelWorksheet structSheet)
        {
            Dictionary<string, ExcelStructSheet> allStructDic = new Dictionary<string, ExcelStructSheet>();
            int maxColumnNum = structSheet.Dimension.End.Column;
            int maxRowNum = structSheet.Dimension.End.Row;

            int fullNameIndex = 0;
            int fieldNameIndex = 0;
            int typeIndex = 0;
            for (var col = 1; col <= maxColumnNum; col++)
            {
                switch (GetValue(structSheet, 1, col))
                {
                    case "full_name":
                        fullNameIndex = col;
                        break;
                    case "*fields":
                        fieldNameIndex = col;
                        typeIndex = col + 1;
                        break;
                    default:
                        break;
                }

                if (fullNameIndex != 0 && fullNameIndex != 0 && typeIndex != 0)
                    break;
            }

            for (var row = 1; row <= maxRowNum; row++)
            {
                string fullName = "";
                switch (GetValue(structSheet, row, 1))
                {
                    case "":
                        fullName = structSheet.Cells[row, fullNameIndex].Value.ToString();
                        Dictionary<string, string> structDic = new Dictionary<string, string>();
                        do
                        {
                            structDic.Add(structSheet.Cells[row, fieldNameIndex].Value.ToString(),
                                structSheet.Cells[row, typeIndex].Value.ToString());
                            row++;
                        } while (structSheet.Cells[row, fullNameIndex].Value.ToString().Length == 0 &&
                                 structSheet.Cells[row, fieldNameIndex].Value.ToString().Length > 0);

                        allStructDic.Add(fullName, new ExcelStructSheet(fullName, structDic));
                        break;
                    default:
                        break;
                }
            }

            return allStructDic;
        }

        private int GetTypeIndex(ExcelWorksheet curWorksheet)
        {
            for (int col = 1; col <= curWorksheet.Dimension.End.Column; ++col)
            {
                if (curWorksheet.Cells[1, col].Value.ToString() == "##type")
                {
                    return col;
                }
            }

            return -1;
        }

        public Dictionary<string, ExcelBook> GetExcelDic()
        {
            return m_ExcelBooksDic;
        }
    }
}