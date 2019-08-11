using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ReportCreater.Entitys;
using ReportCreater.FileHandler;

namespace ReportCreater
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //string fileName;
        PayDetailFileHandler payDetailFile;
        private void BtnOpenFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.SelectedPath = Properties.Settings1.Default.curpath;

            if(fbd.ShowDialog() ==DialogResult.OK)
            {
               
                labFilePath.Text = fbd.SelectedPath;
                Properties.Settings1.Default.curpath = fbd.SelectedPath;
                Properties.Settings1.Default.Save();
                try
                {
                    payDetailFile = new PayDetailFileHandler(fbd.SelectedPath);
                    log("文件选择完成" + payDetailFile.fileName);
                    btnCalc.Enabled = true;
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    btnCalc.Enabled = false;
                    log(ex.Message);
                }

                //MessageBox.Show("读取完成");
            }
        }



        public void log(string msg)
        {
            txtLog.Text += msg + "\r\n";
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
        }

        private void BtnCalc_Click(object sender, EventArgs e)
        {
            DateTime selecteddate = dateTimePicker1.Value;
            string result = "";
            //债务融资工具缴款明细
            payDetailFile.loadData(selecteddate, int.Parse(txtQishu.Text));
            decimal curmonth = payDetailFile.getCurMonthPaySum();
            decimal lastyearmonth = payDetailFile.getLastYearMonthPaySum();
            decimal year = payDetailFile.getYearPaySum();
            decimal lastyear = payDetailFile.getLastYearPaySum();
            int avgcount;
            decimal avgamt;
            payDetailFile.getThisYearDayAvg(out avgcount, out avgamt);
            log("2数据加载完成");
            result += string.Format("截至当日，{0}月债务融资工具合计缴款{1}亿元，发行金额同比{2}%。\r\n" +
                "从年度情况看，债务融资工具合计发行{3}亿元，较去年同期({4}亿元){5}%；" +
                "2019年日均缴款{6}只，日均缴款规模{7}亿元。",
                selecteddate.Month,
                decimal.Round(curmonth,2),
                LYJUtil.getupdown(decimal.Round(decimal.Divide(curmonth, lastyearmonth) * 100 - 100, 0)),
                LYJUtil.changewan(decimal.Round(year,2)),
                LYJUtil.changewan(decimal.Round(lastyear,2)),
                LYJUtil.getupdown(decimal.Round(decimal.Divide(year, lastyear) * 100 - 100, 0)),
                avgcount,
                decimal.Round(avgamt,2)
                );
            textBox1.Text = result;
        }
    }
}
