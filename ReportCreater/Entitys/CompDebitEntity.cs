using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace ReportCreater.Entitys
{
    public class CompDebitEntity
    {
        public string code { get; set; }
        public string bondName { get; set; }
        public decimal planAmt { get; set; }
        public decimal pubAmt { get; set; }
        public decimal amt { get; set; }
        public DateTime calcDate { get; set; }
        public static CompDebitEntity getFromCell(Row row, SharedStringTablePart t,bool is企业债,string fName)
        {
            string curCol = "";
            try
            {
                if (row != null)
                {
                    CompDebitEntity entity = new CompDebitEntity();
                    List<Cell> cells = row.Descendants<Cell>().ToList();
                    if (cells[0] == null)
                    {
                        return null;
                    }
                    curCol = "A";
                    entity.code = LYJUtil.GetValue(LYJUtil.GetCell("A", row.RowIndex, cells), t);
                    if (string.IsNullOrWhiteSpace(entity.code))
                    {
                        return null;
                    }
                    curCol = "B";
                    Cell cellB = LYJUtil.GetCell("B", row.RowIndex, cells);
                    if (cellB != null)
                    {
                        entity.bondName = LYJUtil.GetValue(cellB, t);
                    }
                    else
                    {
                        return null;
                    }

                    curCol = "F";
                    string planAmtStr = LYJUtil.GetValue(LYJUtil.GetCell("F", row.RowIndex, cells), t);
                    curCol = "H";
                    string pubAmtStr = LYJUtil.GetValue(LYJUtil.GetCell("H", row.RowIndex, cells), t);
                    entity.planAmt = decimal.Parse(planAmtStr, System.Globalization.NumberStyles.Float);
                    if (string.IsNullOrWhiteSpace(pubAmtStr))
                    {
                        entity.pubAmt = 0;
                        entity.amt = entity.planAmt;
                    }
                    else
                    {
                        entity.pubAmt = decimal.Parse(pubAmtStr, System.Globalization.NumberStyles.Float);
                        entity.amt = entity.pubAmt;
                    }

                    if (is企业债)
                    {
                        curCol = "E";
                    }
                    else
                    {
                        curCol = "X";
                    }
                    string dateValue = LYJUtil.GetValue(LYJUtil.GetCell(curCol, row.RowIndex, cells), t);
                    entity.calcDate = DateTime.FromOADate(double.Parse(dateValue));
                    return entity;

                }
                else
                {
                    throw new MyException("存在空行");
                }
            }
            catch(Exception ex)
            {
                string msg = ex.Message;
                if(row !=null)
                {
                    msg = "新发行债券(" + fName +")第" + row.RowIndex + "行" + curCol +"列存在问题";
                }
                throw new MyException(msg + "\r\n" + ex.Message + ex.StackTrace);
            }
            
        }
    }
}
