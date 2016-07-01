using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace ExcelTOOL.Public
{
    public class EXCELINFO
    {
        public string A { get; set; }
        public Decimal B1 { get; set; }
        public Decimal C1 { get; set; }
        public Decimal D1 { get; set; }
        public Decimal E1 { get; set; }
        public Decimal F1 { get; set; }

        public Decimal B2 { get; set; }
        public Decimal C2 { get; set; }
        public Decimal D2 { get; set; }
        public Decimal E2 { get; set; }
        public Decimal F2 { get; set; }

        public Decimal B_Diff { get; set; }
        public Decimal C_Diff { get; set; }
        public Decimal D_Diff { get; set; }
        public Decimal E_Diff { get; set; }
        public Decimal F_Diff { get; set; }

        public EXCELINFO(string s, List<Decimal> list)
        {
            int i = 0;
            this.A = s;
            this.B1 = list[i++];
            this.C1 = list[i++];
            this.D1 = list[i++];
            this.E1 = list[i++];
            this.F1 = list[i++];
            this.B2 = list[i++];
            this.C2 = list[i++];
            this.D2 = list[i++];
            this.E2 = list[i++];
            this.F2 = list[i++];
            this.compare();
        }
        private void compare()
        {
            this.B_Diff = this.B1 - this.B2;
            this.C_Diff = this.C1 - this.C2;
            this.D_Diff = this.D1 - this.D2;
            this.E_Diff = this.E1 - this.E2;
            this.F_Diff = this.F1 - this.F2;
        }
    }
    public class PingAnExcelAnalysis
    {
        /// <summary>
        /// 读取Excel文件工作簿到DataTable，文件格式为平安信托综合信息查询导出Excel
        /// </summary>
        /// <param name="filename">文件路径</param>
        /// <param name="ws_index">工作簿，默认为第一个</param>
        /// <returns></returns>
        public Dictionary<string, List<Decimal>> ReadExcelFileToTable(string filename, int ws_index = 1)
        {
            System.Diagnostics.Process[] procs_old = System.Diagnostics.Process.GetProcessesByName("excel");
            object missing = System.Reflection.Missing.Value;
            Application excel = new Application();
            if (excel == null)
            {
                return null;
            }

            // 以只读的形式打开EXCEL文件
            excel.Visible = false; 
            excel.UserControl = true;
            Workbook wb = excel.Application.Workbooks.Open(filename, missing, true, missing, missing, missing, missing, missing, missing, true, missing, missing, missing, missing, missing);

            Worksheet ws = (Worksheet)wb.Worksheets.get_Item(ws_index);        //取得第一个工作薄
            int rowsint = ws.UsedRange.Cells.Rows.Count;                       //得到行数    (包括标题列)

            Range rngA = ws.Cells.get_Range("A2", "A" + (rowsint - 1));        //取得数据范围区域 (不包括标题列) 最后一行为汇总行，剔除
            Range rngB = ws.Cells.get_Range("B2", "B" + (rowsint - 1));        
            Range rngC = ws.Cells.get_Range("C2", "C" + (rowsint - 1));       
            Range rngD = ws.Cells.get_Range("D2", "D" + (rowsint - 1));        
            Range rngE = ws.Cells.get_Range("E2", "E" + (rowsint - 1));        
            Range rngF = ws.Cells.get_Range("F2", "F" + (rowsint - 1));

            object[,] arryA = (object[,])rngA.Value2;
            object[,] arryB = (object[,])rngB.Value2;
            object[,] arryC = (object[,])rngC.Value2;
            object[,] arryD = (object[,])rngD.Value2;
            object[,] arryE = (object[,])rngE.Value2;
            object[,] arryF = (object[,])rngF.Value2;

            Dictionary<string, List<Decimal>> dict = new Dictionary<string, List<Decimal>>();

            for (int i = 1; i <= rowsint - 1 - 1; i++)
            {
                dict.Add(arryA[i, 1].ToString(), new List<Decimal>{ Convert.ToDecimal(arryB[i, 1].ToString()),
                                                                    Convert.ToDecimal(arryC[i, 1].ToString()),
                                                                    Convert.ToDecimal(arryD[i, 1].ToString()),
                                                                    Convert.ToDecimal(arryE[i, 1].ToString()),
                                                                    Convert.ToDecimal(arryF[i, 1].ToString())});
            }
            excel.Quit(); 
            excel = null;

            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("excel");
            foreach (System.Diagnostics.Process pro in procs)
            {
                bool isexist = false;
                foreach (System.Diagnostics.Process pro_old in procs_old)
                {
                    if (pro.Id == pro_old.Id)
                    {
                        isexist = true;
                        break;
                    }
                }
                if (!isexist)
                    pro.Kill();//没有更好的方法,只有杀掉进程
            }

            return dict;
        }

        public List<EXCELINFO> CompExcelFileUsingTable(Dictionary<string, List<Decimal>> d1, Dictionary<string, List<Decimal>> d2)
        {
            List<EXCELINFO> list = new List<EXCELINFO>();
            Dictionary<string, List<Decimal>> dict_diff = new Dictionary<string, List<Decimal>>(d1);
            foreach (string s in d2.Keys)
            {
                if (dict_diff.ContainsKey(s))
                {
                    dict_diff[s].AddRange(d2[s] as List<Decimal>);
                }
                else
                {
                    for (int i = 0; i < 5; i++)  //前五个元素置空
                    {
                        dict_diff[s].Add(0);
                    }
                    dict_diff[s].AddRange(d2[s] as List<Decimal>);
                }
            }
            foreach (string s in dict_diff.Keys)
            {
                list.Add(new EXCELINFO(s, dict_diff[s]));
            }
            dict_diff = null;
            return list;
        }
    }
}
