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
            DirectoryInfo folder = new DirectoryInfo(folderPath);
            ReadAllExcelData(folder);
        }

        private void ReadAllExcelData(DirectoryInfo folder)
        {
            if (!folder.Exists)
            {
                AddErrorMsg($"{folder.Name}文件夹不存在");
                return;
            }

            foreach (DirectoryInfo subFolder in folder.GetDirectories())
            {
                ReadAllExcelData(subFolder);
            }

            foreach (FileInfo file in folder.GetFiles("*.xlsx"))
            {
                if (file.Name.Contains("~$"))
                {
                    continue;
                }

                ReadOneExcel(file);
            }
        }

        public void ReadOneExcel(FileInfo file)
        {
            FileStream fileStream = null;
            ExcelPackage excel = null;
            Dictionary<string, ExcelStructSheet> structSheetDic = null;

            try
            {
                fileStream = new FileStream(file.Name, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                excel = new ExcelPackage(fileStream);
                List<ExcelSheet> sheets = new List<ExcelSheet>();
                // for (int i = 0; i < excel.Workbook.Worksheets.Count; ++i)
                // {
                if (excel.Workbook.Worksheets.Count == 0)
                {
                    AddErrorMsg("Excel表里的sheet数量为0");
                    return;
                }

                var curWorksheet = excel.Workbook.Worksheets[0];
                List<ExcelColumn> sheetColumns = new List<ExcelColumn>();
                int maxColumnNum = curWorksheet.Dimension.End.Column;
                int maxRowNum = curWorksheet.Dimension.End.Row;
                for (var col = 2; col <= maxColumnNum; col++)
                {
                    List<string> columnContent = new List<string>();
                    string fieldName = "";
                    string aliasName = "";
                    string dataType = "";
                    bool isStruct = false;
                    ExcelColumn excelStructColumn = null;
                    for (var row = 1; row <= maxRowNum; row++)
                    {
                        switch (curWorksheet.Cells[row, 1].Value)
                        {
                            case "##comment":
                                aliasName = curWorksheet.Cells[row, col].Value.ToString();
                                break;
                            case "##var":
                                fieldName = curWorksheet.Cells[row, col].Value.ToString();
                                if (fieldName.Contains("*"))
                                {
                                    isStruct = true;
                                    if (structSheetDic == null)
                                    {
                                        if (excel.Workbook.Worksheets.Count >= 2)
                                        {
                                            structSheetDic = new Dictionary<string, ExcelStructSheet>();
                                            var structSheet = excel.Workbook.Worksheets[1];
                                            structSheetDic = GetAllStruct(structSheet);
                                        }
                                        else
                                        {
                                            AddErrorMsg("找不到struct定义sheet");
                                            return;
                                        }
                                    }
                                    int typeIndex = GetTypeIndex(curWorksheet);
                                    if (typeIndex == -1)
                                    {
                                        AddErrorMsg($"{curWorksheet.Name} 找不到##type行");
                                        return;
                                    }
                                    var structNameString = curWorksheet.Cells[typeIndex, col].Value.ToString();
                                    var ss = structNameString.Split(',');
                                    var structName = ss[ss.Length - 1];
                                    if (structSheetDic.ContainsKey(structName))
                                    {
                                        excelStructColumn = GetStructColumn(curWorksheet, ref col, row, structSheetDic[structName].GetStructDic());
                                    }
                                    else
                                    {
                                        AddErrorMsg($"找不到{fieldName} struct");
                                        return;
                                    }
                                }

                                break;
                            case "##type":
                                dataType = curWorksheet.Cells[row, col].Value.ToString();
                                break;
                            case "##":
                                break;
                            case "":
                                columnContent.Add(curWorksheet.Cells[row, col].Value.ToString());
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

                sheets.Add(new ExcelSheet(curWorksheet.Name, sheetColumns));
                // }

                m_ExcelBooksDic.Add(file.Name, new ExcelBook(file.Name, sheets));
            }
            catch (Exception e)
            {
                AddErrorMsg($"读取{file.Name}出错   详细信息:{e}");
            }
            finally
            {
                if (excel != null)
                    excel.Dispose();
                if (fileStream != null)
                    fileStream.Dispose();
            }
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

        private ExcelColumn GetStructColumn(ExcelWorksheet curWorksheet, ref int col, int row,
             Dictionary<string, string> structSheetDic)
        {
            int startIndex = col;
            List<ExcelColumn> structColumns = new List<ExcelColumn>();
            string structFieldName = curWorksheet.Cells[row, col].Value.ToString();
            do
            {
                string aliasName = "";
                string fieldName = "";
                string dataType = "";
                List<string> columnContent = new List<string>();
                for (int i = 1; i <= curWorksheet.Dimension.End.Row; ++i)
                {
                    if (i == row) continue;
                    switch (curWorksheet.Cells[row, 1].Value)
                    {
                        case "##comment":
                            aliasName = curWorksheet.Cells[row, col].Value.ToString();
                            break;
                        case "##var":
                            fieldName = curWorksheet.Cells[row, col].Value.ToString();
                            break;
                        case "##type":
                            if (structSheetDic.ContainsKey(fieldName))
                            {
                                dataType = structSheetDic[fieldName];
                            }
                            else
                            {
                                AddErrorMsg($"{structFieldName}中{fieldName} 找不到");
                            }
                            break;
                        case "##":
                            break;
                        case "":
                            columnContent.Add(curWorksheet.Cells[row, col].Value.ToString());
                            break;
                        default:
                            break;
                    }
                }
                structColumns.Add(new ExcelColumn(fieldName, aliasName, dataType,columnContent, col));
                col++;
                if (col > curWorksheet.Dimension.End.Column)
                {
                    AddErrorMsg("在读取struct column的时候越界了");
                    break;
                }
            } while (curWorksheet.Cells[row, col].Value.ToString().Length == 0 &&
                     curWorksheet.Cells[row + 1, col].Value.ToString().Length > 0);


            return new ExcelColumn(structFieldName, structColumns, startIndex, col);
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
                switch (structSheet.Cells[1, col].Value)
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
                switch (structSheet.Cells[row, 1].Value)
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


        public void PrintErrorMsg(string msg)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(msg);
            Console.ResetColor();
        }

        
        public void PrintSuccessMsg()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("读取成功");
            Console.ResetColor();
        }
        private void AddErrorMsg(string msg)
        {
            if (!m_ErrorMsg.Contains(msg))
            {
                m_ErrorMsg.Add(msg);
            }
        }

        public Dictionary<string, ExcelBook> GetExcelDic()
        {
            return m_ExcelBooksDic;
        }
    }
}