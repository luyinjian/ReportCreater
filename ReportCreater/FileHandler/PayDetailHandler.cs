using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ReportCreater.Entitys;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportCreater.FileHandler
{
    public abstract class PayDetailHandler
    {
        //public event EventHandler<MyEventArgs> ReportProcess;
        public DateTime dateNow;
        public int qishu;
        public string fileName { get; set; }
        public List<RZGJPayDtlEntity> dataList { get; set; }
        public abstract void loadData(DateTime date, int _qishu);
        public decimal getCurMonthPaySum()
        {
            if (dataList != null)
            {
                List<RZGJPayDtlEntity> curMonthData = dataList
                    .Where(n => n.payDate.Month == dateNow.Month
                            && n.payDate.Year == dateNow.Year
                            && n.payDate.Day <= dateNow.Day).ToList();
                decimal sum = 0;
                foreach (RZGJPayDtlEntity entity in curMonthData)
                {
                    sum += entity.pubAmount;
                }
                return sum;
            }
            else
            {
                throw new MyException("未加载文件");
            }
        }
        public decimal getLastYearMonthPaySum()
        {
            if (dataList != null)
            {
                List<RZGJPayDtlEntity> curMonthData = dataList
                    .Where(n => n.payDate.Month == dateNow.Month
                    && n.payDate.Year == dateNow.AddYears(-1).Year
                    && n.payDate.Day <= dateNow.Day).ToList();
                decimal sum = 0;
                foreach (RZGJPayDtlEntity entity in curMonthData)
                {
                    sum += entity.pubAmount;
                }
                return sum;
            }
            else
            {
                throw new MyException("未加载文件");
            }
        }
        public decimal getYearPaySum()
        {
            if (dataList != null)
            {
                List<RZGJPayDtlEntity> curMonthData = dataList
                    .Where(n => n.payDate.Year == dateNow.Year
                    && (n.payDate.Month < dateNow.Month || (n.payDate.Month == dateNow.Month && n.payDate.Day <= dateNow.Day))).ToList();
                decimal sum = 0;
                foreach (RZGJPayDtlEntity entity in curMonthData)
                {
                    sum += entity.pubAmount;
                }
                return sum;
            }
            else
            {
                throw new MyException("未加载文件");
            }
        }
        public decimal getLastYearPaySum()
        {
            if (dataList != null)
            {
                List<RZGJPayDtlEntity> curMonthData = dataList
                    .Where(n => n.payDate.Year == dateNow.AddYears(-1).Year
                    && (n.payDate.Month < dateNow.Month || (n.payDate.Month == dateNow.Month && n.payDate.Day <= dateNow.Day))).ToList();
                decimal sum = 0;
                foreach (RZGJPayDtlEntity entity in curMonthData)
                {
                    sum += entity.pubAmount;
                }
                return sum;
            }
            else
            {
                throw new MyException("未加载文件");
            }
        }
        public void getThisYearDayAvg(out int avgCount, out decimal avgAmt)
        {
            if (dataList != null)
            {
                List<RZGJPayDtlEntity> curMonthData = dataList
                    .Where(n => n.payDate.Year == dateNow.Year
                    && (n.payDate.Month < dateNow.Month || (n.payDate.Month == dateNow.Month && n.payDate.Day <= dateNow.Day))).ToList();
                decimal sum = 0;
                foreach (RZGJPayDtlEntity entity in curMonthData)
                {
                    sum += entity.pubAmount;
                }
                avgCount = Convert.ToInt32((decimal.Round(decimal.Divide(decimal.Parse(curMonthData.Count.ToString()), decimal.Parse(qishu.ToString())), 2)));
                avgAmt = decimal.Round(decimal.Divide(sum, qishu), 0);
            }
            else
            {
                throw new MyException("未加载文件");
            }
        }

        public string getHangYeFenBu()
        {
            var tmpList = dataList.Where(n => n.payDate.Year == dateNow.Year
                                        && n.payDate.Month == dateNow.Month
                                        && n.payDate.Day == dateNow.Day)
                                        .GroupBy(m => m.hangye_1st)
                                        .Select(p => new
                                        {
                                            hangye_1st = p.Key,
                                            Count = p.Count(),
                                            amount = p.Sum(m => m.pubAmount)
                                        }).OrderByDescending(q => q.Count).ThenByDescending(f => f.amount);

            StringBuilder sb = new StringBuilder();
            foreach (var s in tmpList)
            {
                sb.Append(s.hangye_1st + "、");
            }
            sb.Remove(sb.Length - 1, 1);
            return sb.ToString();
        }
        public void getChengjianToday(out int count, out decimal amount)
        {
            count = 0;
            amount = 0;
            var todayList = dataList.Where(n => n.payDate.Year == dateNow.Year
                                        && n.payDate.Month == dateNow.Month
                                        && n.payDate.Day == dateNow.Day);

            foreach (var t in todayList)
            {
                if (t.hangye_1st.Contains("城市基础"))
                {
                    count++;
                    amount = decimal.Add(amount, t.pubAmount);
                }
            }
            amount = decimal.Round(amount, 2);
        }

        public void getMingyingToday(out int count, out decimal amount)
        {
            count = 0;
            amount = 0;
            var todayList = dataList.Where(n => n.payDate.Year == dateNow.Year
                                        && n.payDate.Month == dateNow.Month
                                        && n.payDate.Day == dateNow.Day);

            foreach (var t in todayList)
            {
                if (t.ownnerType.Contains("民营"))
                {
                    count++;
                    amount = decimal.Add(amount, t.pubAmount);
                }
            }

            amount = decimal.Round(amount, 2);
        }
    }
}
