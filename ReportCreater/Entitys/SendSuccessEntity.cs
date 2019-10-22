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
            string curCol = "";
            try
            {
                if (row != null)
                {
                    SendSuccessEntity entity = new SendSuccessEntity();
                    List<Cell> cells = row.Descendants<Cell>().ToList();
                    curCol = "B";
                    Cell cellB = LYJUtil.GetCell("B", row.RowIndex, cells);
                    if (cellB != null)
                    {
                        entity.bondName = LYJUtil.GetValue(cellB, t);
                        if (string.IsNullOrWhiteSpace(entity.bondName))
                        {
                            return null;
                        }
                    }
                    else
                    {
                        return null;
                    }
                    curCol = "C";
                    entity.bondManager = LYJUtil.GetValue(LYJUtil.GetCell("C", row.RowIndex, cells), t);
                    curCol = "E";
                    entity.bondType = LYJUtil.GetValue(LYJUtil.GetCell("E", row.RowIndex, cells), t);
                    curCol = "G";
                    entity.bondLevel = LYJUtil.GetValue(LYJUtil.GetCell("G", row.RowIndex, cells), t);
                    curCol = "K";
                    string pubAmtStr = LYJUtil.GetValue(LYJUtil.GetCell("K", row.RowIndex, cells), t);
                    entity.pubAmout = decimal.Parse(pubAmtStr, System.Globalization.NumberStyles.Float);

                    return entity;

                }
                else
                {
                    throw new MyException("存在空行");
                }
            }
            catch(Exception ex)
            {
                string msg = "向清算所发送登记材料第" + row.RowIndex + "行" + curCol + "列存在问题";
                throw new MyException(msg + ex.Message + ex.StackTrace);
            }
        }
    }
}
