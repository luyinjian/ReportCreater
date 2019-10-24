using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportCreater
{
    class LYJUtil
    {
        /// <summary>
        /// 获取单元格信息  这也是官方获取值的方法
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="stringTablePart">stringTablePart就是WorkbookPart.SharedStringTablePart，它存储了所有以SharedStringTable方式存储数据的子元素。</param>
        /// <returns></returns>
        public static string GetValue(Cell cell, SharedStringTablePart stringTablePart)
        {
            if (cell.ChildElements.Count == 0)
                return null;
            //get cell value
            String value = cell.CellValue.InnerText;
            //Look up real value from shared string table
            if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                value = stringTablePart.SharedStringTable
                    .ChildElements[Int32.Parse(value)]
                    .InnerText;
            return value;
        }

        public static Cell GetCell(string col,string row,List<Cell> cells)
        {
            return cells.Where(n => n.CellReference.Value.Equals(col + row)).FirstOrDefault();
        }

        public static string changewan(decimal input)
        {
            if(input>9999)
            {
                return decimal.Round(decimal.Divide(input, 10000), 2) + "万";
            }
            return decimal.Round(input,0).ToString();
        }

        public static DateTime GetDateTime(string value)
        {
            double d = 0;
            if(double.TryParse(value,out d))
            {
                return DateTime.FromOADate(d);
            }
            else
            {
                return DateTime.Parse(value);
            }
        }

        public static string getupdown(decimal input)
        {
            if(input < 0)
            {
                return "下降" + input*-1;
            }
            else
            {
                return "上升" + input;
            }
        }
    }
}
