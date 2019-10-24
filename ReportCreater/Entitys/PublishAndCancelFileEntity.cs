using ReportCreater.FileHandler;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace ReportCreater.Entitys
{

    public class PublishAndCancelFileEntity
    {
        public string seqNo { get; set; } 
        public DateTime publishDate { get; set; }
        public string pubOrCancel { get; set; }
        public string fullName { get; set; }
        public decimal amount { get; set; }
        public DateTime startDate { get; set; }
        public DateTime endDate { get; set; }
        
        public static PublishAndCancelFileEntity getFromRow(Row row,SharedStringTablePart t)
        {
            string curCol = "";
            string dateValue = "";

            try
            {
                if (row != null)
                {
                    PublishAndCancelFileEntity entity = new PublishAndCancelFileEntity();
                    List<Cell> cells = row.Descendants<Cell>().ToList();
                    curCol = "A";
                    entity.seqNo = LYJUtil.GetValue(LYJUtil.GetCell(curCol, row.RowIndex, cells), t);

                    curCol = "B";
                    dateValue = LYJUtil.GetValue(LYJUtil.GetCell(curCol, row.RowIndex, cells), t);
                    entity.publishDate = LYJUtil.GetDateTime(dateValue);

                    curCol = "C";
                    entity.pubOrCancel = LYJUtil.GetValue(LYJUtil.GetCell(curCol, row.RowIndex, cells), t);

                    curCol = "E";
                    entity.fullName = LYJUtil.GetValue(LYJUtil.GetCell(curCol, row.RowIndex, cells), t);
                    
                    curCol = "J";
                    string amtValue = LYJUtil.GetValue(LYJUtil.GetCell(curCol, row.RowIndex, cells), t);
                    entity.amount = decimal.Parse(amtValue, System.Globalization.NumberStyles.Float);

                    curCol = "K";
                    dateValue = LYJUtil.GetValue(LYJUtil.GetCell(curCol, row.RowIndex, cells), t);
                    entity.startDate = LYJUtil.GetDateTime(dateValue);

                    curCol = "L";
                    dateValue = LYJUtil.GetValue(LYJUtil.GetCell(curCol, row.RowIndex, cells), t);
                    entity.endDate = LYJUtil.GetDateTime(dateValue);
                    return entity;
                }
                else
                {
                    throw new MyException("存在空行");
                }
            }
            catch (Exception ex)
            {
                string msg = "当日披露发行文件及申请取消发行债券基本信息列表第" + row.RowIndex + "行" + curCol + "列存在问题";
                throw new MyException(msg + ex.Message + ex.StackTrace);
            }
        }
    }

}
