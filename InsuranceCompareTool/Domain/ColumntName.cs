using System.ComponentModel;
using InsuranceCompareTool.Models;
namespace InsuranceCompareTool.Domain {
    public enum LogicColumnName
    {
        [DescriptionAttribute(BillSheetColumns.CURRENT_SERVICE_ID)]
        ServiceID,

        [DescriptionAttribute("客服在职状态")]
        ServiceStatus,
        [DescriptionAttribute(BillSheetColumns.SELLER_STATE)]

        SellerStatus,
        [DescriptionAttribute(BillSheetColumns.BILL_PRICE)]
        CurrentPrice,

        [Description(BillSheetColumns.CUSTOMER_PASSPORT_ID)]
        CustomerPassportId,

        [Description(BillSheetColumns.BILL_ID)] 
        BillID,

        [Description("缴费地址")]
        PayAddress,
        [Description(BillSheetColumns.PREVIOUS_SERVICE_ID)]
        PreviousServiceID,

        [Description(BillSheetColumns.SELLER_ID)]
        SellerID,

        [Description("同一投保人，客服专员不一致")]
        DifferentService,

        [Description(BillSheetColumns.PAY_NO)]
        PayNo,
        [Description("投保人与营销员所在地不同")]
        DifferentArea 

    }
}