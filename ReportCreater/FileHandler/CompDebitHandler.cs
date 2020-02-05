using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ReportCreater.Entitys;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace ReportCreater.FileHandler
{
    public class CompDebitHandler
    {
        public string thisYearFileName { get; set; }
        public string lastYearFileName { get; set; }
        public string thisYearFileNameQ { get; set; }
        public string lastYearFileNameQ { get; set; }
        public DateTime dateNow { get; set; }

        public List<CompDebitEntity> thisYearData { get; set; }
        public List<CompDebitEntity> lastYearData { get; set; }
        public List<CompDebitEntity> thisYearDataQ { get; set; }
        public List<CompDebitEntity> lastYearDataQ { get; set; }
        public CompDebitHandler(string filePath,DateTime selectedDate)
        {
            string thisyearname = System.Configuration.ConfigurationManager.AppSettings["compDebitThisYear"];
            string lastyearname = System.Configuration.ConfigurationManager.AppSettings["compDebitLastYear"];
            string thisyearnameQ = System.Configuration.ConfigurationManager.AppSettings["compDebitThisYearQ"];
            string lastyearnameQ = System.Configuration.ConfigurationManager.AppSettings["compDebitLastYearQ"];
            thisYearFileName = filePath + "\\" + thisyearname;
            lastYearFileName = filePath + "\\" + lastyearname;
            thisYearFileNameQ = filePath + "\\" + thisyearnameQ;
            lastYearFileNameQ = filePath + "\\" + lastyearnameQ;

            if (!File.Exists(thisYearFileName))
            {
                thisYearFileName = null;
                throw new MyException("找不到文件" + thisYearFileName);
            }
            if (!File.Exists(lastYearFileName))
            {
                lastYearFileName = null;
                throw new MyException("找不到文件" + lastYearFileName);
            }
            if (!File.Exists(thisYearFileNameQ))
            {
                thisYearFileNameQ = null;
                throw new MyException("找不到文件" + thisYearFileNameQ);
            }
            if (!File.Exists(lastYearFileNameQ))
            {
                lastYearFileNameQ = null;
                throw new MyException("找不到文件" + lastYearFileNameQ);
            }
            dateNow = selectedDate;
        }

        public void loadData()
        {
            if (thisYearFileName == null || lastYearFileName == null ||
                thisYearFileNameQ == null || lastYearFileNameQ == null)
            {
                throw new MyException("文件未选择");
            }

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(thisYearFileName, false))
            {
                WorkbookPart workbook = doc.WorkbookPart;
                WorkbookPart wbPart = doc.WorkbookPart;
                List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                WorksheetPart worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheets[0].Id);
                Worksheet sheet = worksheetPart.Worksheet;
                List<Row> rows = sheet.Descendants<Row>().ToList();
                if (rows.Count < 2)
                {
                    throw new MyException("表格数据不对");
                }
                List<Cell> firstRow = rows.FirstOrDefault().Descendants<Cell>().ToList();
                string aTitle = LYJUtil.GetValue(firstRow[0], workbook.SharedStringTablePart);
                if (aTitle.Trim() != "交易代码")
                {
                    throw new MyException("表格列不对");
                }
                string fTitle = LYJUtil.GetValue(firstRow[5], workbook.SharedStringTablePart);
                if (fTitle.Trim() != "计划发行规模(亿)")
                {
                    throw new MyException("表格列不对");
                }
                string hTitle = LYJUtil.GetValue(firstRow[7], workbook.SharedStringTablePart);
                if (hTitle.Trim() != "发行规模(亿)")
                {
                    throw new MyException("表格列不对");
                }
                thisYearData = new List<CompDebitEntity>();
                for (int i = 1; i < rows.Count; i++)
                {
                    var data = CompDebitEntity.getFromCell(rows[i], workbook.SharedStringTablePart,false, "新发行债券今年");
                    if (data != null)
                    {
                        thisYearData.Add(data);
                    }
                }
            }

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(lastYearFileName, false))
            {
                WorkbookPart workbook = doc.WorkbookPart;
                WorkbookPart wbPart = doc.WorkbookPart;
                List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                WorksheetPart worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheets[0].Id);
                Worksheet sheet = worksheetPart.Worksheet;
                List<Row> rows = sheet.Descendants<Row>().ToList();
                if (rows.Count < 2)
                {
                    throw new MyException("表格数据不对");
                }
                List<Cell> firstRow = rows.FirstOrDefault().Descendants<Cell>().ToList();
                string aTitle = LYJUtil.GetValue(firstRow[0], workbook.SharedStringTablePart);
                if (aTitle.Trim() != "交易代码")
                {
                    throw new MyException("表格列不对");
                }
                string fTitle = LYJUtil.GetValue(firstRow[5], workbook.SharedStringTablePart);
                if (fTitle.Trim() != "计划发行规模(亿)")
                {
                    throw new MyException("表格列不对");
                }
                string hTitle = LYJUtil.GetValue(firstRow[7], workbook.SharedStringTablePart);
                if (hTitle.Trim() != "发行规模(亿)")
                {
                    throw new MyException("表格列不对");
                }
                lastYearData = new List<CompDebitEntity>();
                for (int i = 1; i < rows.Count; i++)
                {
                    var data = CompDebitEntity.getFromCell(rows[i], workbook.SharedStringTablePart,false, "新发行债券去年");
                    if (data != null)
                    {
                        lastYearData.Add(data);
                    }
                }
            }

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(thisYearFileNameQ, false))
            {
                WorkbookPart workbook = doc.WorkbookPart;
                WorkbookPart wbPart = doc.WorkbookPart;
                List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                WorksheetPart worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheets[0].Id);
                Worksheet sheet = worksheetPart.Worksheet;
                List<Row> rows = sheet.Descendants<Row>().ToList();
                if (rows.Count < 2)
                {
                    throw new MyException("表格数据不对");
                }
                List<Cell> firstRow = rows.FirstOrDefault().Descendants<Cell>().ToList();
                string aTitle = LYJUtil.GetValue(firstRow[0], workbook.SharedStringTablePart);
                if (aTitle.Trim() != "交易代码")
                {
                    throw new MyException("表格列不对");
                }
                string fTitle = LYJUtil.GetValue(firstRow[5], workbook.SharedStringTablePart);
                if (fTitle.Trim() != "计划发行规模(亿)")
                {
                    throw new MyException("表格列不对");
                }
                string hTitle = LYJUtil.GetValue(firstRow[7], workbook.SharedStringTablePart);
                if (hTitle.Trim() != "发行规模(亿)")
                {
                    throw new MyException("表格列不对");
                }
                thisYearDataQ = new List<CompDebitEntity>();
                for (int i = 1; i < rows.Count; i++)
                {
                    var data = CompDebitEntity.getFromCell(rows[i], workbook.SharedStringTablePart, true, "新发行债券今年企");
                    if (data != null)
                    {
                        thisYearDataQ.Add(data);
                    }
                }
            }

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(lastYearFileNameQ, false))
            {
                WorkbookPart workbook = doc.WorkbookPart;
                WorkbookPart wbPart = doc.WorkbookPart;
                List<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>().ToList();
                WorksheetPart worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheets[0].Id);
                Worksheet sheet = worksheetPart.Worksheet;
                List<Row> rows = sheet.Descendants<Row>().ToList();
                if (rows.Count < 2)
                {
                    throw new MyException("表格数据不对");
                }
                List<Cell> firstRow = rows.FirstOrDefault().Descendants<Cell>().ToList();
                string aTitle = LYJUtil.GetValue(firstRow[0], workbook.SharedStringTablePart);
                if (aTitle.Trim() != "交易代码")
                {
                    throw new MyException("表格列不对");
                }
                string fTitle = LYJUtil.GetValue(firstRow[5], workbook.SharedStringTablePart);
                if (fTitle.Trim() != "计划发行规模(亿)")
                {
                    throw new MyException("表格列不对");
                }
                string hTitle = LYJUtil.GetValue(firstRow[7], workbook.SharedStringTablePart);
                if (hTitle.Trim() != "发行规模(亿)")
                {
                    throw new MyException("表格列不对");
                }
                lastYearDataQ = new List<CompDebitEntity>();
                for (int i = 1; i < rows.Count; i++)
                {
                    var data = CompDebitEntity.getFromCell(rows[i], workbook.SharedStringTablePart, true, "新发行债券去年企");
                    if (data != null)
                    {
                        lastYearDataQ.Add(data);
                    }
                }
            }
        }

        public decimal getThisYearSum()
        {
            decimal sum = 0;
            foreach (CompDebitEntity entity in thisYearData)
            {
                sum += entity.amt;
            }
            return sum;
        }

        public decimal getLastYearSum()
        {
            decimal sum = 0;
            foreach (CompDebitEntity entity in lastYearData)
            {
                sum += entity.amt;
            }
            return sum;
        }

        public int getTodayCount()
        {
            int count = 0;
            foreach(var entity in thisYearData)
            {
                if(entity.calcDate.Year == dateNow.Year
                    && entity.calcDate.Month == dateNow.Month
                    && entity.calcDate.Day == dateNow.Day)
                {
                    count++;
                }
            }
            return count;

        }

        public decimal getTodaySum()
        {
            decimal sum = 0;
            foreach (var entity in thisYearData)
            {
                if (entity.calcDate.Year == dateNow.Year
                    && entity.calcDate.Month == dateNow.Month
                    && entity.calcDate.Day == dateNow.Day)
                {
                    sum += entity.amt;
                }
            }
            return sum;

        }

        public decimal getThisYearSumQ()
        {
            decimal sum = 0;
            foreach (CompDebitEntity entity in thisYearDataQ)
            {
                sum += entity.amt;
            }
            return sum;
        }

        public decimal getLastYearSumQ()
        {
            decimal sum = 0;
            foreach (CompDebitEntity entity in lastYearDataQ)
            {
                sum += entity.amt;
            }
            return sum;
        }

        public int getTodayCountQ()
        {
            int count = 0;
            foreach (var entity in thisYearDataQ)
            {
                if (entity.calcDate.Year == dateNow.Year
                    && entity.calcDate.Month == dateNow.Month
                    && entity.calcDate.Day == dateNow.Day)
                {
                    count++;
                }
            }
            return count;

        }

        public decimal getTodaySumQ()
        {
            decimal sum = 0;
            foreach (var entity in thisYearDataQ)
            {
                if (entity.calcDate.Year == dateNow.Year
                    && entity.calcDate.Month == dateNow.Month
                    && entity.calcDate.Day == dateNow.Day)
                {
                    sum += entity.amt;
                }
            }
            return sum;

        }
    }
}
