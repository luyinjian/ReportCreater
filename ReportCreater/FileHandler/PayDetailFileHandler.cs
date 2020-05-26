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
    public class PayDetailFileHandler : PayDetailHandler
    {
        public PayDetailFileHandler(string filePath)
        {
            string payDtlFileName = System.Configuration.ConfigurationManager.AppSettings["payDetailFileName"];
            if (!File.Exists(filePath + "\\" + payDtlFileName))
            {
                throw new MyException("文件不存在" + payDtlFileName);
            }

            fileName = filePath + "\\" + payDtlFileName;
        }
        public override void loadData(DateTime date, int _qishu)
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
                    RZGJPayDtlEntity data = RZGJPayDtlEntity.getFromCell(rows[i], workbook.SharedStringTablePart);
                    if(data != null)
                    {
                        dataList.Add(data);
                    }
                    
                }
            }
        }
    }
}
