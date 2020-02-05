using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ReportCreater.Entitys;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportCreater.FileHandler
{
    class PublishAndCancelFileHandler
    {
        public event EventHandler<MyEventArgs> ReportProcess;

        private DateTime dateNow;

        public string fileName { get; set; }

        public List<PublishAndCancelFileEntity> todayList { get; set; }
        public List<PublishAndCancelFileEntity> hisList { get; set; }

        public PublishAndCancelFileHandler(string filePath, DateTime datetime)
        {
            string payDtlFileName = System.Configuration.ConfigurationManager.AppSettings["publishAndCancelFile"];
            dateNow = datetime;
            payDtlFileName = string.Format(payDtlFileName, dateNow.ToString("yyyyMMdd"));
            if (!File.Exists(filePath + "\\" + payDtlFileName))
            {
                throw new MyException("文件不存在" + payDtlFileName);
            }

            fileName = filePath + "\\" + payDtlFileName;
        }

        public void loadData()
        {
            if (fileName == null)
            {
                throw new MyException("文件读取失败");
            }

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbook = doc.WorkbookPart;
                WorkbookPart wbPart = doc.WorkbookPart;
                List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                if (sheets.Count != 3)
                {
                    throw new MyException("当日披露文件sheet数不对");
                }
                //当日读取
                todayList = readEntity(doc, workbook, sheets, 1);
                //累计数据
                hisList = readEntity(doc, workbook, sheets, 2);

                if (ReportProcess != null)
                {
                    ReportProcess(this, new MyEventArgs() { code = "0000", msg = "数据加载成功" });
                }
            }
        }

        public void getTodayData(out int count,out decimal amount)
        {
            count = 0;
            amount = 0;
            foreach(PublishAndCancelFileEntity pf in todayList)
            {
                if(pf.pubOrCancel.Trim().Equals("发行"))
                {
                    count++;
                    amount = decimal.Add(amount, pf.amount);
                }
            }
            //转为亿元
            amount = decimal.Divide(amount, 10000);
        }

        public void getHisDoing(out int count,out decimal amount)
        {
            count = 0;
            amount = 0;

            DateTime tmpDateStart = DateTime.Parse(dateNow.ToString("yyyy-MM-dd") + " 23:59:59");
            DateTime tmpDateEnd = DateTime.Parse(dateNow.ToString("yyyy-MM-dd") + " 00:00:00");

            

            //List<PublishAndCancelFileEntity> resultList = new List<PublishAndCancelFileEntity>();

            foreach (PublishAndCancelFileEntity pf in hisList)
            {
                int days = (pf.endDate - pf.startDate).Days;


                if (pf.startDate.CompareTo(tmpDateStart) <=0
                    &&
                    pf.endDate.CompareTo(tmpDateEnd) >=0
                    &&
                    pf.pubOrCancel.Trim().Equals("发行")
                    &&
                    (days<=31))
                {
                    count++;
                    amount = decimal.Add(amount, pf.amount);
                    //resultList.Add(pf);
                }
            }

            //using(StreamWriter sw = new StreamWriter("d:\\qyclog.txt"))
            //{
            //    foreach(var en in resultList)
            //    {
            //        sw.WriteLine("{0},{1},{2},{3},{4},{5},{6}",
            //            en.seqNo,
            //            en.publishDate.ToString("yyyy/MM/dd"),
            //            en.pubOrCancel,
            //            en.fullName,
            //            en.amount,
            //            en.startDate.ToString("yyyy/MM/dd"),
            //            en.endDate.ToString("yyyy/MM/dd"));
            //    }
            //    sw.Flush();
            //    sw.Close();
            //}
            //转为亿元
            amount = decimal.Divide(amount, 10000);
        }

        private List<PublishAndCancelFileEntity> readEntity(SpreadsheetDocument doc,WorkbookPart workbook, List<Sheet> sheets,int sheet_index)
        {
            List<PublishAndCancelFileEntity> entity = new List<PublishAndCancelFileEntity>();
            WorksheetPart worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheets[sheet_index].Id);
            Worksheet sheet = worksheetPart.Worksheet;
            List<Row> rows = sheet.Descendants<Row>().ToList();
            if (rows.Count < 2)
            {
                throw new MyException("当日sheet数据不对");
            }
            List<Cell> firstRow = rows.FirstOrDefault().Descendants<Cell>().ToList();
            if (firstRow.Count < 20)
            {
                throw new MyException("当日sheet列数不对");
            }
            string bTitle = LYJUtil.GetValue(firstRow[1], workbook.SharedStringTablePart);
            if (bTitle.Trim() != "披露日期")
            {
                throw new MyException("表格列数不对");
            }
            string cTitle = LYJUtil.GetValue(firstRow[2], workbook.SharedStringTablePart);
            if (!(cTitle.Trim().Contains("发行") && cTitle.Trim().Contains("取消发行")))
            {
                throw new MyException("表格列数不对");
            }
            string jTitle = LYJUtil.GetValue(firstRow[9], workbook.SharedStringTablePart);
            if (jTitle.Trim() != "计划发行金额（万元）")
            {
                throw new MyException("表格列数不对");
            }
            for (int i = 1; i < rows.Count; i++)
            {
                var rowdat = PublishAndCancelFileEntity.getFromRow(rows[i], workbook.SharedStringTablePart);
                if(rowdat != null)
                {
                    entity.Add(rowdat);
                }
                
            }
            return entity;
        }

        
    }
}
