using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace ConvertExcel
{
    public class ErrorMsgMgr : Singleton<ErrorMsgMgr>
    {
        // private static readonly ErrorMsgMgr instance = new ErrorMsgMgr();
        //
        // static ErrorMsgMgr()
        // {
        // }
        //
        // private ErrorMsgMgr()
        // {
        // }
        //
        // public static ErrorMsgMgr Instance
        // {
        //     get { return instance; }
        // }

        private List<string> m_ErrorMsg = new List<string>();

        public void AutoPrintErrorOrSucces()
        {
            if(m_ErrorMsg.Count == 0)
                PrintSuccessMsg("成功导出，结束");
            else
            {
                AddErrorMsg("如碰到问题请联系王慧东解决");
                PrintAllErrorMsg();
            }
        }
        
        public void PrintErrorMsg(string msg)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(msg);
            Console.ResetColor();
        }

        public void PrintAllErrorMsg()
        {
            Console.ForegroundColor = ConsoleColor.Red;
            foreach (var msg in m_ErrorMsg)
            {
                Console.WriteLine(msg);
            }
            Console.ResetColor();
        }

        public void PrintSuccessMsg(string msg)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(msg);
            Console.ResetColor();
        }

        public void AddErrorMsg(string msg)
        {
            if (!m_ErrorMsg.Contains(msg))
            {
                m_ErrorMsg.Add(msg);
            }
        }

        public void ClearErrorMsg()
        {
            m_ErrorMsg.Clear();
        }
    }
}