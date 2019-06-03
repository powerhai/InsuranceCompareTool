using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceCompareTool.Dispatch
{
    public class BillOperator
    {
        public string Title { get; set; }
        public DispatcherBase Dispatcher { get; set; }
        public BillFilter BillFilter { get; set; }


    }
}
