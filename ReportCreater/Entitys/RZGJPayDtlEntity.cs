using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace ReportCreater.Entitys
{
    public class RZGJPayDtlEntity
    {
        public string seqNo { get; set; }
        public string pubCompName { get; set; }
        public string bondName { get; set; }
        public DateTime payDate { get; set; }
        public decimal pubAmount { get; set; }

        public static RZGJPayDtlEntity getFromCell(Row row,SharedStringTablePart t)
        {
            if(row != null)
            {
                RZGJPayDtlEntity entity = new RZGJPayDtlEntity();
                List<Cell> cells = row.Descendants<Cell>().ToList();
                entity.seqNo = LYJUtil.GetValue(LYJUtil.GetCell("A",row.RowIndex,cells), t);
                entity.pubCompName = LYJUtil.GetValue(LYJUtil.GetCell("B", row.RowIndex, cells), t);
                entity.bondName = LYJUtil.GetValue(LYJUtil.GetCell("C", row.RowIndex, cells), t);
                string amtValue = LYJUtil.GetValue(LYJUtil.GetCell("K", row.RowIndex, cells), t);
               // if(amtValue.Contains("E"))
              //  {
                    entity.pubAmount = decimal.Parse(amtValue, System.Globalization.NumberStyles.Float);
                //  }
                //   else
                //    {
                //        entity.pubAmount = decimal.Parse(amtValue);
                //    }
                string dateValue = LYJUtil.GetValue(LYJUtil.GetCell("N", row.RowIndex, cells), t);
                entity.payDate = DateTime.FromOADate(double.Parse(dateValue));
                return entity;
            }
            else
            {
                throw new MyException("存在空行");
            }
        }
    }
}
