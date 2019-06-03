using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceCompareTool.Dispatch
{
    public abstract class DispatcherBase
    {
        public abstract void DispatchBills();
    }

    public class DispatchToManager : DispatcherBase
    {
        public override void DispatchBills()
        {

        }
    }
     
}
