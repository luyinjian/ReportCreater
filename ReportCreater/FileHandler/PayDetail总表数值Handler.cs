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
    public class PayDetail总表数值Handler : PayDetailHandler
    {
        public PayDetail总表数值Handler(string filePath)
        {
            string payDtlFileName = System.Configuration.ConfigurationManager.AppSettings["总表数值"];
            if (!File.Exists(filePath + "\\" + payDtlFileName))
            {
                throw new MyException("文件不存在" + payDtlFileName);
            }

            fileName = filePath + "\\" + payDtlFileName;
        }
        public override void loadData(DateTime date, int _qishu)
        {
            if (fileName == null)
            {
                throw new MyException("文件读取失败");
            }
            dateNow = date;
            qishu = _qishu;
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))
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
                string kTitle = LYJUtil.GetValue(firstRow[10], workbook.SharedStringTablePart);
                if (kTitle.Trim() != "发行额(亿元)")
                {
                    throw new MyException("表格列数不对");
                }
                string mTitle = LYJUtil.GetValue(firstRow[12], workbook.SharedStringTablePart);
                if (mTitle.Trim() != "缴款日")
                {
                    throw new MyException("表格列数不对");
                }

                dataList = new List<RZGJPayDtlEntity>();
                for (int i = 1; i < rows.Count; i++)
                {
                    RZGJPayDtlEntity data = RZGJPayDtlEntity.getFrom总表数值Cell(rows[i], workbook.SharedStringTablePart);
                    if(data != null)
                    {
                        dataList.Add(data);
                    }
                    
                }
            }
        }
    }
}
