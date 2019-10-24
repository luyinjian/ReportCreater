using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportCreater.Entitys;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace ReportCreater.FileHandler
{
    public class PayDetailFileHandler
    {
        public event EventHandler<MyEventArgs> ReportProcess;

        private DateTime dateNow;
        private int qishu;

        public string fileName { get; set; }
        public List<RZGJPayDtlEntity> dataList { get; set; }
        public PayDetailFileHandler(string filePath)
        {
            string payDtlFileName = System.Configuration.ConfigurationManager.AppSettings["payDetailFileName"];
            if (!File.Exists(filePath + "\\" + payDtlFileName))
            {
                throw new MyException("文件不存在" + payDtlFileName);
            }

            fileName = filePath + "\\" + payDtlFileName;
        }
        public void loadData(DateTime date, int _qishu)
        {
            if(fileName==null)
            {
                throw new MyException("文件读取失败");
            }
            dateNow = date;
            qishu = _qishu;
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName,false))
            {
                WorkbookPart workbook = doc.WorkbookPart;
                WorkbookPart wbPart = doc.WorkbookPart;
                List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                WorksheetPart worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheets[0].Id);
                Worksheet sheet = worksheetPart.Worksheet;
                List<Row> rows = sheet.Descendants<Row>().ToList();
                if(rows.Count<2)
                {
                    throw new MyException("表格数据不对");
                }
                List<Cell> firstRow = rows.FirstOrDefault().Descendants<Cell>().ToList();
                if(firstRow.Count!=23)
                {
                    throw new MyException("表格列数不对");
                }
                string kTitle = LYJUtil.GetValue(firstRow[10], workbook.SharedStringTablePart);
                if(kTitle.Trim() != "发行额（亿元）")
                {
                    throw new MyException("表格列数不对");
                }
                string nTitle = LYJUtil.GetValue(firstRow[13], workbook.SharedStringTablePart);
                if (nTitle.Trim() != "缴款日")
                {
                    throw new MyException("表格列数不对");
                }

                dataList = new List<RZGJPayDtlEntity>();
                for(int i=1;i<rows.Count;i++)
                {
                    dataList.Add(RZGJPayDtlEntity.getFromCell(rows[i], workbook.SharedStringTablePart));
                }
                if(ReportProcess!=null)
                {
                    ReportProcess(this, new MyEventArgs() { code = "0000", msg = "数据加载成功" });
                }
            }
        }
        public decimal getCurMonthPaySum()
        {
            if(dataList!=null)
            {
                List<RZGJPayDtlEntity> curMonthData = dataList
                    .Where(n => n.payDate.Month == dateNow.Month 
                            && n.payDate.Year == dateNow.Year
                            && n.payDate.Day <= dateNow.Day).ToList();
                decimal sum = 0;
                foreach(RZGJPayDtlEntity entity in curMonthData)
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
                    && (n.payDate.Month < dateNow.Month || (n.payDate.Month==dateNow.Month && n.payDate.Day<=dateNow.Day))).ToList();
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
        public void getThisYearDayAvg(out int avgCount,out decimal avgAmt)
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
                avgCount = Convert.ToInt32((decimal.Round(decimal.Divide(decimal.Parse(curMonthData.Count.ToString()),decimal.Parse(qishu.ToString())),2)));
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
                                            Count = p.Count()
                                        }).OrderByDescending(q=>q.Count);

            StringBuilder sb = new StringBuilder();
            foreach(var s in tmpList)
            {
                sb.Append(s.hangye_1st + "、");
            }
            sb.Remove(sb.Length - 1, 1);
            return sb.ToString();
        }
        public void getChengjianToday(out int count,out decimal amount)
        {
            count = 0;
            amount = 0;
            var todayList = dataList.Where(n => n.payDate.Year == dateNow.Year
                                        && n.payDate.Month == dateNow.Month
                                        && n.payDate.Day == dateNow.Day);

            foreach(var t in todayList)
            {
                if(t.hangye_1st.Contains("城市基础"))
                {
                    count++;
                    amount = decimal.Add(amount, t.pubAmount);
                }
            }
        }

        public void getMingyingToday(out int count,out decimal amount)
        {
            count = 0;
            amount = 0;
            var todayList = dataList.Where(n => n.payDate.Year == dateNow.Year
                                        && n.payDate.Month == dateNow.Month
                                        && n.payDate.Day == dateNow.Day);

            foreach (var t in todayList)
            {
                if(t.ownnerType.Contains("民营"))
                {
                    count++;
                    amount = decimal.Add(amount, t.pubAmount);
                }
            }
        }
    }
}
