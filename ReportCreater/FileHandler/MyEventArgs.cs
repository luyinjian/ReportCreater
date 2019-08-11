using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportCreater.FileHandler
{
    public class MyEventArgs : EventArgs
    {
        public string code { get; set; }
        public string msg { get; set; }
        public int value { get; set; }
    }
}
