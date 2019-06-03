using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceCompareTool.Models
{
    public class Policy
    { 
        public List<Condition> Conditions { get;  } = new List<Condition>();
        public bool RefuseNextAction { get; set; }
    }
    public class Allocation
    {

    }

    public class Condition
    {

    }
}
