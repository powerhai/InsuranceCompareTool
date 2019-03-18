using System.ComponentModel;
namespace InsuranceCompareTool.Models {
    public enum BillStatus
    {
        [Description("")]
        NoNeedUpdate,
        [Description("客服非在职")]
        ServiceNotAvailable,
        [Description("营销员非在职")]
        SellerNotAvailable,
        [Description("义乌单")]
        Yiwu,
        [Description("客服空缺")]
        NoService,
        [Description("非上期客服")]
        NoLastService,
        [Description("客服是营销")]
        ServiceSameToSeller,
        [Description("绑定客服")]
        BindService,
        [Description("跨区-234期")]
        AcrossArea234,
        [Description("跨区-5期")]
        AcrossArea5,
        [Description("区拓")]
        AreaSell,
        [Description("分辨错误")]
        Error
    }
}