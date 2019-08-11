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

        public List<CompDebitEntity> thisYearData { get; set; }
        public List<CompDebitEntity> lastYearData { get; set; }
        public CompDebitHandler(string filePath)
        {
            string thisyearname = System.Configuration.ConfigurationManager.AppSettings["compDebitThisYear"];
            string lastyearname = System.Configuration.ConfigurationManager.AppSettings["compDebitLastYear"];
            thisYearFileName = filePath + "\\" + thisyearname;
            lastYearFileName = filePath + "\\" + lastyearname;

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
        }

        public void loadData()
        {
            if (thisYearFileName == null || lastYearFileName == null)
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
                    var data = CompDebitEntity.getFromCell(rows[i], workbook.SharedStringTablePart);
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
                    var data = CompDebitEntity.getFromCell(rows[i], workbook.SharedStringTablePart);
                    if (data != null)
                    {
                        lastYearData.Add(data);
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
    }
}
