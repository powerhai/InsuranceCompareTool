using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Prism.Mvvm;
namespace InsuranceCompareTool.Models
{
    public abstract class ViewModelBase : BindableBase
    {
        public abstract string Title { get; set; }

    }
}
