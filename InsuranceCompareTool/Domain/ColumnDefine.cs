using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceCompareTool.Domain
{
    public class ColumnDefine
    {
        public string Name { get; set; }
        public Type Type { get; set; }
        public bool IsSystemColumn { get; set; }
    }
}
