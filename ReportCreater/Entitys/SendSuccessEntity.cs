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
        public string bondLmt { get; set; }
        public string bondLevel { get; set; }
        public double pubPercent { get; set; }
        public double buyMulti { get; set; }
        public decimal pubAmout { get; set; }
    }
}
