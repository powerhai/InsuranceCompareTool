using System.ComponentModel;
namespace InsuranceCompareTool.Domain {
    public enum BoolType
    {
        [DescriptionAttribute("且")]
        And,
        [DescriptionAttribute("或")]
        Or
    }
}