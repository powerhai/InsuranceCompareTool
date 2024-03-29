﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceCompareTool.Models
{
     

    public static class AreaNames
    {
        public const string JH = "金华";
        public const string YK = "永康";
        public const string WY = "武义";
        public const string LX = "兰溪";
        public const string YW = "义乌";
        public const string PJ = "浦江";
        public const string PA = "磐安";
        public const string DY = "东阳";
        public const string QT = "区拓";
        public static readonly string[] Areas = new string[8]{JH, YK, WY, LX, YW, PJ, PA, DY };
    }

    public static class BillSheetColumns
    {
        public const string ORGANIZATION4 = "四级机构"; 
        //public const string LAST_SERVICE_ID = "客服专员工号";
        //public const string LAST_SERVICE_NAME = "客服专员姓名";
        public const string CURRENT_SERVICE_ID = "客服专员工号";
        public const string CURRENT_SERVICE_NAME = "客服专员姓名"; 
        public const string BILL_ID = "保单号";
        public const string BILL_ID2 = "保险单号";
        public const string CUSTOMER_NAME = "投保人";
        public const string PAY_DATE = "应缴日期";
        public const string PAY_DATE2 = "缴费起始期";

        public const string SELLER_ID = "营销员工号";
        public const string SELLER_NAME = "营销员姓名";
        public const string PAY_NO = "缴费期次";
        public const string SELL_MODEL = "销售渠道";
        public const string SELLER_STATE = "营销员状态"; 
        public const string CUSTOMER_PASSPORT_ID = "投保人证件"; 
        public const string CUSTOMER_PASSPORT_ID2 = "投保人身份证";

        public const string CUSTOMER_ADDRESS = "投保人地址";
        public const string PAY_ADDRESS = "缴费地址";
        public const string PAY_ADDRESS2 = "缴费地址（一）";
        public const string PAY_ADDRESS4 = "缴费地址（二）";
        public const string PAY_ADDRESS3 = "缴费地址1";
        public const string PAY_ADDRESS5 = "缴费地址2";

        public const string BILL_PRICE = "本期应缴保费";
        public const string BILL_PRICE2 = "保费";
        public const string PRODUCT_NAME = "险种名称";
        public const string PRODUCT_ID = "险种代码";
        public const string PER_PAY = "是否垫缴";
        public const string CREDITOR = "被保险人";
        public const string CUSTOMER_ACCOUNT = "帐号";
        public const string MOBILE_PHONE = "手机";
        public const string IS_OURS = "是否自保件";
        public const string IS_OURS3 = "是否为自保件";
        public const string IS_OURS2 = "自保件";
        public const string STATUS_OF_DJ = "督缴状态";

        public const string ASSIGNED_SERVICE_ID = "改派后专员工号";
        public const string ASSIGNED_SERVICE_NAME = "改派后专员姓名";
        public const string SYS_SERVICE_STATUS = "客服状态";
        public const string PREVIOUS_SERVICE_ID = "上期服务人员工号";
        public const string PREVIOUS_SERVICE = "上期服务人员";
        public const string PREVIOUS_PAY_DATE = "上期缴费日期";
        public const string CLIENT_AREA = "投保人地区";
        public const string SYS_FILTER = "System_Filter";
        public const string SYS_FINISHED = "System_Finished";
        public const string SYS_HISTORY = "System_History";
        public const string SYS_SERVICE = "System_SERVICE";
        public const string SYS_SERVICE_ID = "System_SERVICEID"; 
        public const string SYS_ERROR = "System_Error";
        public const string SYS_GUID = "System_GUID";
        public const string SYS_SERVICE_AREA = "客服地区";
        public const string SYS_SELLER_AREA = "营销员地区";
    }

    public static class MemberSheetColumns
    {
        public const string ID = "工号";
        public const string NAME = "姓名";
        public const string POSITION = "职务";
        public const string STATUS = "状态";
        public const string AREA = "地区";
        public const string VIRTUAL = "虚拟";
        public const string REPORTABLE = "列出清单";
    }

    public static class DepartmentSheetColumns
    {
        public const string ID = "部门代码";
        public const string NAME = "部门名称";
        public const string AREA = "地区";
    }

    public static class RelationSheetColumns
    {
        public const string SERVICE_ID = "客服专员工号";
        public const string SERVICE_NAME = "客服专员姓名";
        public const string AREA = "地区";
        public const string BIND_TYPE = "绑定类型";
        public const string BIND_STRING = "绑定字符";
    }

    public static class StatisticsColumns
    {
        public const string COUNT = "总单数";
        public const string SUM = "总金额";
    }

    public static class StatusNames
    {
        public const string ZAI_ZI = "在职";
    }
    public static class PayModelNames
    {
        public const string AREA_SELL = "区域拓展";
    }
    public static class PositionNames
    {
        public const string MUSTER = "地区主管"; 
        public const string SELLER = "营销员";
        public const string SERVICE = "客服专员";
    }
}
