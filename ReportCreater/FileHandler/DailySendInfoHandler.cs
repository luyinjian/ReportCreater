using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ReportCreater.Entitys;

namespace ReportCreater.FileHandler
{
    public class DailySendInfoHandler
    {
        public string fileName { get; set; }

        List<SendSuccessEntity> dataList { get; set; }
        public DailySendInfoHandler(string filePath,DateTime date)
        {
            string tmp = System.Configuration.ConfigurationManager.AppSettings["dailySendFileName"];
            string tmpName = filePath + "\\" + string.Format(tmp, date.ToString("yyyyMMdd"));
            if(!File.Exists(tmpName))
            {
                throw new MyException("文件不存在" + tmpName);
            }
            fileName = tmpName;
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
                WorksheetPart worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheets[0].Id);
                Worksheet sheet = worksheetPart.Worksheet;
                List<Row> rows = sheet.Descendants<Row>().ToList();
                if (rows.Count < 3)
                {
                    throw new MyException("表格数据不对");
                }
                List<Cell> firstRow = rows[1].Descendants<Cell>().ToList();
                string aTitle = LYJUtil.GetValue(LYJUtil.GetCell("B", rows[1].RowIndex, firstRow), workbook.SharedStringTablePart);
                if (aTitle.Trim() != "债券简称")
                {
                    throw new MyException("表格列数不对");
                }
                string cTitle = LYJUtil.GetValue(LYJUtil.GetCell("E", rows[1].RowIndex, firstRow), workbook.SharedStringTablePart);
                if (cTitle.Trim() != "债券品种")
                {
                    throw new MyException("表格列数不对");
                }
                string kTitle = LYJUtil.GetValue(LYJUtil.GetCell("K", rows[1].RowIndex, firstRow), workbook.SharedStringTablePart);
                if (kTitle.Trim() != "发行额（亿元）")
                {
                    throw new MyException("表格列数不对");
                }

                dataList = new List<SendSuccessEntity>();
                for(int i=2;i<rows.Count;i++)
                {
                    var data = SendSuccessEntity.getFromCell(rows[i], workbook.SharedStringTablePart);
                    if(data != null)
                    {
                        dataList.Add(data);
                    }
                }
            }
        }

        public decimal getTotal()
        {
            decimal sum = 0;
            foreach(var data in dataList)
            {
                sum = decimal.Add(sum, data.pubAmout);
            }
            return sum;
        }
    }
}
