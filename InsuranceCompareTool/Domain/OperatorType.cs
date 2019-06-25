using System.ComponentModel;
namespace InsuranceCompareTool.Domain {
    public enum OperatorType
    {
        [DescriptionAttribute("=")]
        Equal,
        [DescriptionAttribute(">")]
        Greater,
        [DescriptionAttribute("<")]
        Less,
        [DescriptionAttribute("包含")]
        Contain,
        Other
    }
}