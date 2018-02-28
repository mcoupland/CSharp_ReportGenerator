using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Reflection;

namespace AppDevReportGenerator
{
    public static class Global
    {
        #region Properties
        public static Excel.Application ExcelApplication;
        public static List<Process> ExcelProcesses = new List<Process>();
        #endregion
        
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        public static void AddExcelProcess(Excel.Application app)
        {
            GetWindowThreadProcessId(app.Hwnd, out int pid);
            Global.ExcelProcesses.Add(Process.GetProcessById(pid));
        }

        public static void ReleaseExcelProcesses()
        {
            try
            {
                if (Global.ExcelApplication == null) { return; }
                GetWindowThreadProcessId(Global.ExcelApplication.Hwnd, out int pid);
                Global.ExcelProcesses.Add(Process.GetProcessById(pid));
                foreach (Process p in Global.ExcelProcesses)
                {
                    if (!string.IsNullOrEmpty(p.ProcessName))
                    {
                        p.Kill();
                    }
                }
            }
            catch{}  // This function is fragile and breaks frequently enough that this is needed
        }
    }
}
