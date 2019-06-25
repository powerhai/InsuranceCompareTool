using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace InsuranceCompareTool.ShareCommon
{
    public class EnumPackager
    {
        public string Title
        {
            get { return Enum.GetDescription(); }
        }
        public Enum Enum { get; set; }

    }
}
