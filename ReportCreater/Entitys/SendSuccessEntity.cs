using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportCreater.Entitys
{
    public class SendSuccessEntity
    {
        public string bondName { get; set; }
        public string bondManager { get; set; }
        public string bondType { get; set; }
        public string bondLevel { get; set; }
        public decimal pubAmout { get; set; }

        public static SendSuccessEntity getFromCell(Row row, SharedStringTablePart t)
        {
            if (row != null)
            {
                SendSuccessEntity entity = new SendSuccessEntity();
                List<Cell> cells = row.Descendants<Cell>().ToList();
                Cell cellB = LYJUtil.GetCell("B", row.RowIndex, cells);
                if (cellB != null)
                {
                    entity.bondName = LYJUtil.GetValue(cellB, t);
                    if(string.IsNullOrWhiteSpace(entity.bondName))
                    {
                        return null;
                    }
                }
                else
                {
                    return null;
                }

                entity.bondManager = LYJUtil.GetValue(LYJUtil.GetCell("C", row.RowIndex, cells), t);
                entity.bondType = LYJUtil.GetValue(LYJUtil.GetCell("E", row.RowIndex, cells), t);
                entity.bondLevel = LYJUtil.GetValue(LYJUtil.GetCell("G", row.RowIndex, cells), t);

                string pubAmtStr = LYJUtil.GetValue(LYJUtil.GetCell("K", row.RowIndex, cells), t);
                entity.pubAmout = decimal.Parse(pubAmtStr, System.Globalization.NumberStyles.Float);
               
                return entity;

            }
            else
            {
                throw new MyException("存在空行");
            }
        }
    }
}
