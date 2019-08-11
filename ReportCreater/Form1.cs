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
        CompDebitHandler compDebitHandler;
        DailySendInfoHandler dailySendInfoHandler;
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
                    log("文件选择：" + payDetailFile.fileName);
                    compDebitHandler = new CompDebitHandler(fbd.SelectedPath);
                    log("文件选择：" + compDebitHandler.thisYearFileName);
                    log("文件选择：" + compDebitHandler.lastYearFileName);
                    dailySendInfoHandler = new DailySendInfoHandler(fbd.SelectedPath, dateTimePicker1.Value);
                    log("文件选择：" + dailySendInfoHandler.fileName);
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
            //发送登记材料初始化
            dailySendInfoHandler.loadData();
            log(dailySendInfoHandler.fileName + "数据加载完成");
            decimal dailyTotalSum = dailySendInfoHandler.getTotal();
            result += string.Format("{0}债务融资工具缴款规模{1}亿元。",
                selecteddate.DayOfWeek,
                LYJUtil.changewan(dailyTotalSum));
            //债务融资工具缴款明细
            payDetailFile.loadData(selecteddate, int.Parse(txtQishu.Text));
            log(payDetailFile.fileName + "数据加载完成");
            decimal curmonth = payDetailFile.getCurMonthPaySum();
            decimal lastyearmonth = payDetailFile.getLastYearMonthPaySum();
            decimal year = payDetailFile.getYearPaySum();
            decimal lastyear = payDetailFile.getLastYearPaySum();
            int avgcount;
            decimal avgamt;
            payDetailFile.getThisYearDayAvg(out avgcount, out avgamt);
            
            result += string.Format("截至当日，{0}月债务融资工具合计缴款{1}亿元，发行金额同比{2}%。\r\n" +
                "从年度情况看，债务融资工具合计发行{3}亿元，较去年同期({4}亿元){5}%；" +
                "2019年日均缴款{6}只，日均缴款规模{7}亿元。",
                selecteddate.Month,
                decimal.Round(curmonth,0),
                LYJUtil.getupdown(decimal.Round(decimal.Divide(curmonth, lastyearmonth) * 100 - 100, 0)),
                LYJUtil.changewan(decimal.Round(year,2)),
                LYJUtil.changewan(decimal.Round(lastyear,2)),
                LYJUtil.getupdown(decimal.Round(decimal.Divide(year, lastyear) * 100 - 100, 0)),
                avgcount,
                decimal.Round(avgamt,0)
                );
            //公司债
            compDebitHandler.loadData();
            log(compDebitHandler.thisYearFileName + compDebitHandler.lastYearFileName + "数据加载完成");
            decimal thisYearCompDebit = compDebitHandler.getThisYearSum();
            decimal lastYearCompDebit = compDebitHandler.getLastYearSum();
            result += string.Format("一般公司债(含ABS)本年累计发行{0}亿元，" +
                "较去年同期({1}亿元){2}%。",
                LYJUtil.changewan(thisYearCompDebit),
                LYJUtil.changewan(lastYearCompDebit),
                LYJUtil.getupdown(decimal.Round(decimal.Divide(thisYearCompDebit, lastYearCompDebit) * 100 - 100, 0))
                );


            textBox1.Text = result;
        }
    }
}
