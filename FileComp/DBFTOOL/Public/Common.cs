using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace dbfcomp
{
    public partial class DBFAnalysis
    {
        //public DBFCompare()
        //{
        //    InitializeComponent();
        //}

        private void button1_Click()
        {
            DataTable dt = GetCompTable(@"C:\Users\Administrator.WINDOWS-TCNHOG0\Desktop\tmp\dbf\20160615.dbf",
                                        @"C:\Users\Administrator.WINDOWS-TCNHOG0\Desktop\tmp\dbf\201606151.dbf");
        }

        /// <summary>
        /// 输入DBF文件，比较对应字段，返回结果集
        /// 结果集字段如下：
        /// RQ                Int32                 日期
        /// JJBH              String                基金编号
        /// JJMC              String                基金名称
        /// ZCDYDM            String                资产单元代码
        /// ZCDYMC            String                资产单元名称
        /// JJDM              String                基金代码
        /// ZCDYBH            String                资产单元编号
        /// DQXJYE1           Decimal               当前现金余额1
        /// DQXJYE2           Decimal               当前现金余额2
        /// DQXJYE_DIFF       Decimal               当前现金余额差值
        /// QCXJYE1           Decimal               期初现金余额1
        /// QCXJYE2           Decimal               期初现金余额2
        /// QCXJYE_DIFF       Decimal               期初现金余额差值
        /// T0JYKY_ZCD1       Decimal               T+0交易可用金额1
        /// T0JYKY_ZCD2       Decimal               T+0交易可用金额2
        /// T0JYKY_ZCD_DIFF   Decimal               T+0交易可用金额差值
        /// T1JYKY_ZCD1       Decimal               T+1交易可用金额1
        /// T1JYKY_ZCD2       Decimal               T+1交易可用金额2
        /// T1JYKY_ZCD_DIFF   Decimal               T+1交易可用金额差值
        /// T0ZLKY_ZCD1       Decimal               T+0指令可用金额1
        /// T0ZLKY_ZCD2       Decimal               T+0指令可用金额2
        /// T0ZLKY_ZCD_DIFF   Decimal               T+0指令可用金额差值
        /// T0ZLKY_NCD1       Decimal               T+0指令可用(不含T+1变化)1
        /// T0ZLKY_NCD2       Decimal               T+0指令可用(不含T+1变化)2
        /// T0ZLKY_NCD_DIFF   Decimal               T+0指令可用(不含T+1变化)差值
        /// T1ZLKY_ZCD1       Decimal               T+1指令可用金额1
        /// T1ZLKY_ZCD2       Decimal               T+1指令可用金额2
        /// T1ZLKY_ZCD_DIFF   Decimal               T+1指令可用金额差值
        /// </summary>
        /// <param name="filename1">DBF文件1路径</param>
        /// <param name="filename2">DBF文件2路径</param>
        /// <returns>dbf文件比较结果集</returns>
        public DataTable GetCompTable(string filename1, string filename2)
        {
            TDbfTable dbf1 = new TDbfTable(filename1);
            TDbfTable dbf2 = new TDbfTable(filename2);

            DataTable dt1 = dbf1.Table;
            foreach (DataRow r in dbf2.Table.Rows)
            {
                dt1.ImportRow(r);
            }
            dt1.DefaultView.Sort = "RQ, ZCDYDM";
            dt1 = dt1.DefaultView.ToTable();

            DataTable dt = new DataTable();
            dt.Columns.Add("RQ",              System.Type.GetType("System.Int32"));             //日期
            dt.Columns.Add("JJBH",            System.Type.GetType("System.String"));            //基金编号
            dt.Columns.Add("JJMC",            System.Type.GetType("System.String"));            //基金名称
            dt.Columns.Add("ZCDYDM",          System.Type.GetType("System.String"));            //资产单元代码
            dt.Columns.Add("ZCDYMC",          System.Type.GetType("System.String"));            //资产单元名称
            dt.Columns.Add("JJDM",            System.Type.GetType("System.String"));            //基金代码
            dt.Columns.Add("ZCDYBH",          System.Type.GetType("System.String"));            //资产单元编号
            dt.Columns.Add("DQXJYE1",         System.Type.GetType("System.Decimal"));           //当前现金余额1
            dt.Columns.Add("DQXJYE2",         System.Type.GetType("System.Decimal"));           //当前现金余额2
            dt.Columns.Add("DQXJYE_DIFF",     System.Type.GetType("System.Decimal"));           //当前现金余额 差值
            dt.Columns.Add("QCXJYE1",         System.Type.GetType("System.Decimal"));           //期初现金余额
            dt.Columns.Add("QCXJYE2",         System.Type.GetType("System.Decimal"));           //
            dt.Columns.Add("QCXJYE_DIFF",     System.Type.GetType("System.Decimal"));           //
            dt.Columns.Add("T0JYKY_ZCD1",     System.Type.GetType("System.Decimal"));           //T+0交易可用金额
            dt.Columns.Add("T0JYKY_ZCD2",     System.Type.GetType("System.Decimal"));           //
            dt.Columns.Add("T0JYKY_ZCD_DIFF", System.Type.GetType("System.Decimal"));           //
            dt.Columns.Add("T1JYKY_ZCD1",     System.Type.GetType("System.Decimal"));           //T+1交易可用金额
            dt.Columns.Add("T1JYKY_ZCD2",     System.Type.GetType("System.Decimal"));           //
            dt.Columns.Add("T1JYKY_ZCD_DIFF", System.Type.GetType("System.Decimal"));           //
            dt.Columns.Add("T0ZLKY_ZCD1",     System.Type.GetType("System.Decimal"));           //T+0指令可用金额
            dt.Columns.Add("T0ZLKY_ZCD2",     System.Type.GetType("System.Decimal"));           //
            dt.Columns.Add("T0ZLKY_ZCD_DIFF", System.Type.GetType("System.Decimal"));           //
            dt.Columns.Add("T0ZLKY_NCD1",     System.Type.GetType("System.Decimal"));           //T+0指令可用(不含T+1变化)
            dt.Columns.Add("T0ZLKY_NCD2",     System.Type.GetType("System.Decimal"));           //
            dt.Columns.Add("T0ZLKY_NCD_DIFF", System.Type.GetType("System.Decimal"));           //
            dt.Columns.Add("T1ZLKY_ZCD1",     System.Type.GetType("System.Decimal"));           //T+1指令可用金额
            dt.Columns.Add("T1ZLKY_ZCD2",     System.Type.GetType("System.Decimal"));           //
            dt.Columns.Add("T1ZLKY_ZCD_DIFF", System.Type.GetType("System.Decimal"));           //
            dt.PrimaryKey = new DataColumn[] { dt.Columns["RQ"], dt.Columns["ZCDYDM"] };

            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                DataRow dr = dt.NewRow();

                dr["RQ"] = dt1.Rows[i]["RQ"];
                dr["JJBH"] = dt1.Rows[i]["JJBH"];
                dr["JJMC"] = dt1.Rows[i]["JJMC"];
                dr["JJDM"] = dt1.Rows[i]["JJDM"];
                dr["ZCDYDM"] = dt1.Rows[i]["ZCDYDM"];
                dr["ZCDYMC"] = dt1.Rows[i]["ZCDYMC"];
                dr["ZCDYBH"] = dt1.Rows[i]["ZCDYBH"];
                dr["DQXJYE1"] = dt1.Rows[i]["DQXJYE"];
                dr["QCXJYE1"] = dt1.Rows[i]["QCXJYE"];
                dr["T0JYKY_ZCD1"] = dt1.Rows[i]["T0JYKY_ZCD"];
                dr["T1JYKY_ZCD1"] = dt1.Rows[i]["T1JYKY_ZCD"];
                dr["T0ZLKY_ZCD1"] = dt1.Rows[i]["T0ZLKY_ZCD"];
                dr["T0ZLKY_NCD1"] = dt1.Rows[i]["T0ZLKY_NCD"];
                dr["T1ZLKY_ZCD1"] = dt1.Rows[i]["T1ZLKY_ZCD"];
                if (i < dt1.Rows.Count - 1 && Convert.ToInt32(dr["RQ"]) == Convert.ToInt32(dt1.Rows[i + 1]["RQ"]) && Convert.ToString(dr["ZCDYDM"]).Equals(Convert.ToString(dt1.Rows[i + 1]["ZCDYDM"])))
                {
                    i++;
                    dr["DQXJYE2"] = dt1.Rows[i]["DQXJYE"];
                    dr["QCXJYE2"] = dt1.Rows[i]["QCXJYE"];
                    dr["T0JYKY_ZCD2"] = dt1.Rows[i]["T0JYKY_ZCD"];
                    dr["T1JYKY_ZCD2"] = dt1.Rows[i]["T1JYKY_ZCD"];
                    dr["T0ZLKY_ZCD2"] = dt1.Rows[i]["T0ZLKY_ZCD"];
                    dr["T0ZLKY_NCD2"] = dt1.Rows[i]["T0ZLKY_NCD"];
                    dr["T1ZLKY_ZCD2"] = dt1.Rows[i]["T1ZLKY_ZCD"];
                }
                else
                {
                    dr["DQXJYE2"] = 0;
                    dr["QCXJYE2"] = 0;
                    dr["T0JYKY_ZCD2"] = 0;
                    dr["T1JYKY_ZCD2"] = 0;
                    dr["T0ZLKY_ZCD2"] = 0;
                    dr["T0ZLKY_NCD2"] = 0;
                    dr["T1ZLKY_ZCD2"] = 0;
                }
                dr["DQXJYE_DIFF"] = Convert.ToDecimal(dr["DQXJYE1"]) - Convert.ToDecimal(dr["DQXJYE2"]);
                dr["QCXJYE_DIFF"] = Convert.ToDecimal(dr["QCXJYE1"]) - Convert.ToDecimal(dr["QCXJYE2"]);
                dr["T0JYKY_ZCD_DIFF"] = Convert.ToDecimal(dr["T0JYKY_ZCD1"]) - Convert.ToDecimal(dr["T0JYKY_ZCD2"]);
                dr["T1JYKY_ZCD_DIFF"] = Convert.ToDecimal(dr["T1JYKY_ZCD1"]) - Convert.ToDecimal(dr["T1JYKY_ZCD2"]);
                dr["T0ZLKY_ZCD_DIFF"] = Convert.ToDecimal(dr["T0ZLKY_ZCD1"]) - Convert.ToDecimal(dr["T0ZLKY_ZCD2"]);
                dr["T0ZLKY_NCD_DIFF"] = Convert.ToDecimal(dr["T0ZLKY_NCD1"]) - Convert.ToDecimal(dr["T0ZLKY_NCD2"]);
                dr["T1ZLKY_ZCD_DIFF"] = Convert.ToDecimal(dr["T1ZLKY_ZCD1"]) - Convert.ToDecimal(dr["T1ZLKY_ZCD2"]);

                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}
