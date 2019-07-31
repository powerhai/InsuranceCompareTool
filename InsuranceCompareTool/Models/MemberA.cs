using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Prism.Mvvm;
namespace InsuranceCompareTool.Models
{
    public class MemberA : BindableBase
    {
        public string Name
        {
            get => mName;
            set => SetProperty(ref mName, value);
        }
        private int mCount;
        private string mName;
        private string mArea;
        private double mPrice;
        public int Count
        {
            get => mCount;
            set
            {
                SetProperty(ref mCount, value);
                RaisePropertyChanged(nameof(Title));
                RaisePropertyChanged(nameof(Title2));
            }
        }
        public double Price
        {
            get => mPrice;
            set
            {
                SetProperty(ref mPrice, value);
                RaisePropertyChanged(nameof(Title));
                RaisePropertyChanged(nameof(Title2));
            }
        }
        public string Area
        {
            get => mArea;
            set => SetProperty(ref mArea, value);
        }
        public string Title
        {
            get
            {
                var name = Name.PadRight(6, ' ');
                var count = Count.ToString().PadLeft(5, ' ');
                var price = Price.ToString("C").PadLeft(12, ' ');
                return $"{name}\t{count}\t{price}\t{ID}";
            }
        }

        public string Title2
        {
            get
            {
                var name = Name.PadRight(6, ' ');
                var count = Count.ToString().PadLeft(5, ' ');
                var price = Price.ToString("C").PadLeft(12, ' ');
                return $"{name}\t{count}\t{price}\t{ID}\t{Area}";
            }
        }

        public string ID { get; set; }
        public string Position { get; set; }
    }
}
