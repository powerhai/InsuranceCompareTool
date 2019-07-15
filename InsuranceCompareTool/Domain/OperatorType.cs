using System.ComponentModel;
namespace InsuranceCompareTool.Domain {
    public enum OperatorType
    {
        [DescriptionAttribute("!=")]
        NoEqual,
        [DescriptionAttribute("=")]
        Equal,
        [DescriptionAttribute(">")]
        Greater,
        [DescriptionAttribute(">=")]
        GreaterAndEqual,
        [DescriptionAttribute("<")]
        Less,
        [DescriptionAttribute("<=")]
        LessAndEqual,
        [DescriptionAttribute("包含")]
        Contain,
        [DescriptionAttribute("")]
        Other
    }
}