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
        PayDetailHandler payDetailFile;
        CompDebitHandler compDebitHandler;
        DailySendInfoHandler dailySendInfoHandler;
        PublishAndCancelFileHandler publishAndCancelFileHandler;
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
                    if(radioButton2.Checked)
                    {
                        payDetailFile = new PayDetail总表数值Handler(fbd.SelectedPath);
                    }
                    else if(radioButton1.Checked)
                    {
                        payDetailFile = new PayDetailFileHandler(fbd.SelectedPath);
                        
                    }
                    else
                    {
                        MessageBox.Show("请选择取哪个版本");
                        return;
                    }
                    log("文件选择：" + payDetailFile.fileName);
                    compDebitHandler = new CompDebitHandler(fbd.SelectedPath);
                    log("文件选择：" + compDebitHandler.thisYearFileName);
                    log("文件选择：" + compDebitHandler.lastYearFileName);
                    dailySendInfoHandler = new DailySendInfoHandler(fbd.SelectedPath, dateTimePicker1.Value);
                    log("文件选择：" + dailySendInfoHandler.fileName);
                    publishAndCancelFileHandler = new PublishAndCancelFileHandler(fbd.SelectedPath, dateTimePicker1.Value);
                    log("文件选择：" + publishAndCancelFileHandler.fileName);
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
            try
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
                if(radioButton1.Checked)
                {
                    lastyear = lastyear + 130;
                }
                int avgcount;
                decimal avgamt;
                
                payDetailFile.getThisYearDayAvg(out avgcount, out avgamt);

                result += string.Format("截至当日，{0}月债务融资工具合计缴款{1}亿元，发行金额同比{2}%。\r\n" +
                    "从年度情况看，债务融资工具合计发行{3}亿元，较去年同期({4}亿元){5}%；" +
                    "2019年日均缴款{6}只，日均缴款规模{7}亿元。",
                    selecteddate.Month,
                    decimal.Round(curmonth, 0,MidpointRounding.AwayFromZero),
                    LYJUtil.getupdown(decimal.Round(decimal.Divide(curmonth, lastyearmonth == 0 ? curmonth : lastyearmonth) * 100 - 100, 0, MidpointRounding.AwayFromZero)),
                    LYJUtil.changewan(decimal.Round(year, 2, MidpointRounding.AwayFromZero)),
                    LYJUtil.changewan(decimal.Round(lastyear, 2, MidpointRounding.AwayFromZero)),
                    LYJUtil.getupdown(decimal.Round(decimal.Divide(year, lastyear == 0 ? year : lastyear) * 100 - 100, 0, MidpointRounding.AwayFromZero)),
                    avgcount,
                    decimal.Round(avgamt, 0, MidpointRounding.AwayFromZero)
                    );
                //公司债
                compDebitHandler.loadData();
                log(compDebitHandler.thisYearFileName + compDebitHandler.lastYearFileName + "数据加载完成");
                decimal thisYearCompDebit = compDebitHandler.getThisYearSum();
                decimal lastYearCompDebit = compDebitHandler.getLastYearSum();
                result += string.Format("公司债(含ABS)本年累计发行{0}亿元，" +
                    "较去年同期({1}亿元){2}%。",
                    LYJUtil.changewan(thisYearCompDebit),
                    LYJUtil.changewan(lastYearCompDebit),
                    LYJUtil.getupdown(decimal.Round(decimal.Divide(thisYearCompDebit, lastYearCompDebit) * 100 - 100, 0, MidpointRounding.AwayFromZero))
                    );
                //簿记建档情况
                publishAndCancelFileHandler.loadData();

                int todayPubCount = 0;
                decimal todayPubAmt = 0;
                int hisDoingCount = 0;
                decimal hisDoingAmt = 0;
                publishAndCancelFileHandler.getTodayData(out todayPubCount, out todayPubAmt);
                publishAndCancelFileHandler.getHisDoing(out hisDoingCount, out hisDoingAmt);

                result += string.Format("\r\n一、整体发行情况\r\n簿记建档情况，{4}挂网{0}只，金额{1}亿元；正在簿记{2}只，金额{3}亿元。",
                    todayPubCount, decimal.Round(todayPubAmt,0,MidpointRounding.AwayFromZero),
                    hisDoingCount, decimal.Round(hisDoingAmt,0, MidpointRounding.AwayFromZero),
                    selecteddate.DayOfWeek);

                result += string.Format("\r\n缴款规模方面，{2}债务融资工具缴款{0}只，金额{1}亿元。",
                    dailySendInfoHandler.dataList.Count,
                    decimal.Round(dailyTotalSum,0, MidpointRounding.AwayFromZero),
                    selecteddate.DayOfWeek);

                result += "\r\n品种分布方面，";
                var pingzhongList = dailySendInfoHandler.getPingZhongFenBu();
                foreach(var pingzhong in pingzhongList)
                {
                    result += string.Format("{0}占{1}%，",
                        pingzhong.Key, pingzhong.Value);
                }
                result = result.Remove(result.Length - 1, 1) + "。";

                result += string.Format("行业分布方面，涵盖{0}等。",
                    payDetailFile.getHangYeFenBu());

                result += string.Format("评级方面，超过{0}%为中高评级。",
                    dailySendInfoHandler.getPingJiPercent());

                int chengJianCount = 0;
                decimal chengJianAmt = 0;
                payDetailFile.getChengjianToday(out chengJianCount, out chengJianAmt);
                result += string.Format("城建类企业发行{0}只，金额共计{1}亿元。",
                    chengJianCount,chengJianAmt);

                int minYingCount = 0;
                decimal minYingAmt = 0;
                payDetailFile.getMingyingToday(out minYingCount, out minYingAmt);
                result += string.Format("民营企业发行{0}只，金额共计{1}亿元。",
                    minYingCount, minYingAmt);

                textBox1.Text = result;
            }
            catch(MyException me)
            {
                MessageBox.Show(me.Message);
            }
            
        }
    }
}
