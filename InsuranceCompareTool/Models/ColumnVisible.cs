using System;
namespace InsuranceCompareTool.Models {
    [Serializable]
    public class ColumnVisible
    { 
        public string Title { get; set; }
        public Guid Guid { get; set; }
        public string[] Columns { get; set; } = new string[0];
    }
}