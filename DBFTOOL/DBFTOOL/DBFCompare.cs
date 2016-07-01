using dbfcomp;
using HSUCF.Controls;
using Hundsun.Framework.BizControls.BaseControl.Export;
using Hundsun.Framework.BizControls.BaseControl.Print;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DBFTOOL
{
    public partial class DBFCompare : Form
    {
        public DBFCompare()
        {
            InitializeComponent();
            btnImport.SetValidateGroup(this.txtFileName1, "group");
            btnImport.SetValidateGroup(this.txtFileName2, "group");
            //btnExport.Exporter.Source = gctlDBFTool.DataSource;
            btnExport.Exporter = new GridControlExporter(gctlDBFTool);
            btnPrint.Printer = new GirdControlPrinter(gctlDBFTool);
        }

        private const decimal StandDiff = 0.005M;

        DBFAnalysis dbfanalysis = new DBFAnalysis();
        
        //DataTable tableSource = null;

        List<DBFINFO> sourceList = new List<DBFINFO>();

        private void Method()
        {
            if (File.Exists(txtFileName1.Value.Trim()) == false)
            {
                MsgBoxUtility.ShowTips("文件1不存在");
                return;
            }
            if (File.Exists(txtFileName2.Value.Trim()) == false)
            {
                MsgBoxUtility.ShowTips("文件2不存在");
                return;
            }
            string file1 = txtFileName1.Value.Trim();
            string file2 = txtFileName2.Value.Trim();
            //string file1 = @"C:\Users\yangzq13148\Desktop\导入\20160615.dbf";
            //string file2 = @"C:\Users\yangzq13148\Desktop\导入\201606151.dbf";
            DataTable tableSource = dbfanalysis.GetCompTable(file1, file2);
            sourceList = ConvertTo(tableSource);
            Fill();
            //sourceList = list.ToList();
            //gctlDBFTool.DataSource = sourceList;
        }

        private void Fill()
        {
            List<DBFINFO> list =
            sourceList.FindAll(p =>

                (chbSame.Checked && (
                Math.Abs(p.DQXJYE_DIFF) <= StandDiff
             && Math.Abs(p.QCXJYE_DIFF) <= StandDiff
             && Math.Abs(p.T0JYKY_ZCD_DIFF) <= StandDiff
             && Math.Abs(p.T0ZLKY_NCD_DIFF) <= StandDiff
             && Math.Abs(p.T0ZLKY_ZCD_DIFF) <= StandDiff
             && Math.Abs(p.T1JYKY_ZCD_DIFF) <= StandDiff
             && Math.Abs(p.T1ZLKY_ZCD_DIFF) <= StandDiff
             )) ||
                (chbDiff.Checked && (
                Math.Abs(p.DQXJYE_DIFF) > StandDiff
             || Math.Abs(p.QCXJYE_DIFF) > StandDiff
             || Math.Abs(p.T0JYKY_ZCD_DIFF) > StandDiff
             || Math.Abs(p.T0ZLKY_NCD_DIFF) > StandDiff
             || Math.Abs(p.T0ZLKY_ZCD_DIFF) > StandDiff
             || Math.Abs(p.T1JYKY_ZCD_DIFF) > StandDiff
             || Math.Abs(p.T1ZLKY_ZCD_DIFF) > StandDiff
                ))
            );

            gctlDBFTool.DataSource = list;
        }

        private void chbSame_CheckedChanged(object sender, EventArgs e)
        {
            //DataTable table = null;
            Fill();
        }

        private void chbDiff_CheckedChanged(object sender, EventArgs e)
        {
            Fill();
        }

        #region DataTable转换为List

        private List<DBFINFO> ConvertTo(DataTable table)
        {
            if (table == null)
            {
                return null;
            }

            List<DBFINFO> returnList = new List<DBFINFO>();
            

            foreach (DataRow row in table.Rows)
            {
                DBFINFO obj = new DBFINFO();
                obj.RQ = Convert.ToInt32(row["RQ"]);
                obj.JJBH = Convert.ToString(row["JJBH"]);
                obj.JJMC = Convert.ToString(row["JJMC"]);
                obj.ZCDYDM = Convert.ToString(row["ZCDYDM"]);
                obj.ZCDYMC = Convert.ToString(row["ZCDYMC"]);
                obj.JJDM = Convert.ToString(row["JJDM"]);
                obj.ZCDYBH = Convert.ToString(row["ZCDYBH"]);
                obj.DQXJYE1 = Convert.ToDecimal(row["DQXJYE1"]);
                obj.DQXJYE2 = Convert.ToDecimal(row["DQXJYE2"]);
                obj.DQXJYE_DIFF = Convert.ToDecimal(row["DQXJYE_DIFF"]);
                obj.QCXJYE1 = Convert.ToDecimal(row["QCXJYE1"]);
                obj.QCXJYE2 = Convert.ToDecimal(row["QCXJYE2"]);
                obj.QCXJYE_DIFF = Convert.ToDecimal(row["QCXJYE_DIFF"]);
                obj.T0JYKY_ZCD1 = Convert.ToDecimal(row["T0JYKY_ZCD1"]);
                obj.T0JYKY_ZCD2 = Convert.ToDecimal(row["T0JYKY_ZCD2"]);
                obj.T0JYKY_ZCD_DIFF = Convert.ToDecimal(row["T0JYKY_ZCD_DIFF"]);
                obj.T1JYKY_ZCD1 = Convert.ToDecimal(row["T1JYKY_ZCD1"]);
                obj.T1JYKY_ZCD2 = Convert.ToDecimal(row["T1JYKY_ZCD2"]);
                obj.T1JYKY_ZCD_DIFF = Convert.ToDecimal(row["T1JYKY_ZCD_DIFF"]);
                obj.T0ZLKY_ZCD1 = Convert.ToDecimal(row["T0ZLKY_ZCD1"]);
                obj.T0ZLKY_ZCD2 = Convert.ToDecimal(row["T0ZLKY_ZCD2"]);
                obj.T0ZLKY_ZCD_DIFF = Convert.ToDecimal(row["T0ZLKY_ZCD_DIFF"]);
                obj.T0ZLKY_NCD1 = Convert.ToDecimal(row["T0ZLKY_NCD1"]);
                obj.T0ZLKY_NCD2 = Convert.ToDecimal(row["T0ZLKY_NCD2"]);
                obj.T0ZLKY_NCD_DIFF = Convert.ToDecimal(row["T0ZLKY_NCD_DIFF"]);
                obj.T1ZLKY_ZCD1 = Convert.ToDecimal(row["T1ZLKY_ZCD1"]);
                obj.T1ZLKY_ZCD2 = Convert.ToDecimal(row["T1ZLKY_ZCD2"]);
                obj.T1ZLKY_ZCD_DIFF = Convert.ToDecimal(row["T1ZLKY_ZCD_DIFF"]);
                returnList.Add(obj);
            }

            return returnList;
        }

        #endregion

        
        #region DBF class

        public class DBFINFO
        {
            public Int32 RQ { get; set; }
            public String JJBH { get; set; }
            public String JJMC { get; set; }
            public String ZCDYDM { get; set; }
            public String ZCDYMC { get; set; }
            public String JJDM { get; set; }
            public String ZCDYBH { get; set; }
            public Decimal DQXJYE1 { get; set; }
            public Decimal DQXJYE2 { get; set; }
            public Decimal DQXJYE_DIFF { get; set; }
            public Decimal QCXJYE1 { get; set; }
            public Decimal QCXJYE2 { get; set; }
            public Decimal QCXJYE_DIFF { get; set; }
            public Decimal T0JYKY_ZCD1 { get; set; }
            public Decimal T0JYKY_ZCD2 { get; set; }
            public Decimal T0JYKY_ZCD_DIFF { get; set; }
            public Decimal T1JYKY_ZCD1 { get; set; }
            public Decimal T1JYKY_ZCD2 { get; set; }
            public Decimal T1JYKY_ZCD_DIFF { get; set; }
            public Decimal T0ZLKY_ZCD1 { get; set; }
            public Decimal T0ZLKY_ZCD2 { get; set; }
            public Decimal T0ZLKY_ZCD_DIFF { get; set; }
            public Decimal T0ZLKY_NCD1 { get; set; }
            public Decimal T0ZLKY_NCD2 { get; set; }
            public Decimal T0ZLKY_NCD_DIFF { get; set; }
            public Decimal T1ZLKY_ZCD1 { get; set; }
            public Decimal T1ZLKY_ZCD2 { get; set; }
            public Decimal T1ZLKY_ZCD_DIFF { get; set; }

        }

        #endregion

        private void hsButton3_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog(this) == DialogResult.OK)
            {
                txtFileName1.Value = file.FileName;
            }
        }

        private void hsButton4_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog(this) == DialogResult.OK)
            {
                txtFileName2.Value = file.FileName;
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            //Method();
            ExcelTOOL.Public.PingAnExcelAnalysis s = new ExcelTOOL.Public.PingAnExcelAnalysis();
            List<ExcelTOOL.Public.EXCELINFO> list = s.CompExcelFileUsingTable(
                s.ReadExcelFileToTable(@"C:\Users\Administrator.WINDOWS-TCNHOG0\Desktop\tmp\dbf\综合信息查询_账户资产30.xls"),
                s.ReadExcelFileToTable(@"C:\Users\Administrator.WINDOWS-TCNHOG0\Desktop\tmp\dbf\综合信息查询_账户资产30t.xls"));

            list.Add(list[1]);
            Method();
        }
        
        private void bgvDBFTool_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            if (e.CellValue == null)
                return;
            string[] containList = { "DQXJYE_DIFF", "QCXJYE_DIFF","T0JYKY_ZCD_DIFF","T1JYKY_ZCD_DIFF","T0ZLKY_ZCD_DIFF","T0ZLKY_NCD_DIFF","T1ZLKY_ZCD_DIFF" };
            if (containList.Contains<string>(e.Column.Name) == true && e.CellValue.ToString() != "")
            {
                decimal data = Convert.ToDecimal(e.CellValue);

                if (Math.Abs(data) > StandDiff)
                {
                    e.Appearance.ForeColor = Color.Red;
                }
            }
        }

    }
}
