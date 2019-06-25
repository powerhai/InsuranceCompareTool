using System.ComponentModel;
namespace InsuranceCompareTool.Domain {
    public enum DispatchType
    {
        /// <summary>
        /// 分配给地区主管
        /// </summary>
        [Description("地区主管")]
        DispatchToAreaManager,
        /// <summary>
        /// 分配给指定人
        /// </summary>
        [Description("指定人员")]
        DispatchToDesignated,
        /// <summary>
        /// 分配给上期客服
        /// </summary>
        [Description("上期客服")]
        DispatchToLastService,
        /// <summary>
        /// 分配给营销员所在地主管
        /// </summary>
        [Description("营销员所在地主管")]
        DispatchToManagerOfSeller,
        /// <summary>
        /// 分配给顾客所在地主管
        /// </summary>
        [Description("顾客所在地主管")]
        DispatchToManagerOfCustomer

    }
}