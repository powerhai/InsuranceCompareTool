using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Linq;
using InsuranceCompareTool.Models;
namespace InsuranceCompareTool.ViewModels
{
    public class PolicyViewViewModel : ViewModelBase
    {
        public PolicyViewViewModel()
        {

        }
        public override string Title { get; set; } = "Policy";
    }
}
