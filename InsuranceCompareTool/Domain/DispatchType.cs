using System.ComponentModel;
namespace InsuranceCompareTool.Domain {
    public enum DispatchType
    {
 
        /// <summary>
        /// 分配给指定人
        /// </summary>
        [Description("指定人员")]
        DispatchToDesignated,
         [Description("客服专员所在地主管")]
        DispatchToManagerOfService,
        /// <summary>
        /// 分配给营销员所在地主管
        /// </summary>
        [Description("营销员所在地主管")]
        DispatchToManagerOfSeller,
 
        [Description("投保人所在地主管")]
        DispatchToManagerOfCustomer,

        [Description("上期客服专员")]
        DispatchToPreviousService,

        [Description("不分配")]
        DoNot
             
    }
}