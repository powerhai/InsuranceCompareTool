using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceCompareTool.Domain
{
    public class ColumnDefine
    {
        public string FormalName
        {
            get { return Name[0]; }
        }
        public string[] Name { get; set; }
        public string ExportName { get; set; }
        public Type Type { get; set; }
        public bool IsSystemColumn { get; set; }
        public bool IsRequired { get; set; }
        public int Sort { get; set; }
        public bool IsCopyAble { get; set; }
    }
}
