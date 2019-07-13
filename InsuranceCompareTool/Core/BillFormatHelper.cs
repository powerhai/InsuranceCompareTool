using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using InsuranceCompareTool.Models;
using log4net;
namespace InsuranceCompareTool.Core
{
    public class BillFormatHelper
    {
        private readonly ILog mLogger;
        private readonly List<Member> mMembers;
        public BillFormatHelper(ILog logger, List<Member> members)
        { 
            mLogger = logger;
            mMembers = members;
        }
        public void Format(DataTable dt)
        {
            FormatPreviousService(dt);
            FormatAddress(dt);
            FormatValidColumns(dt);
            FormatMobilePhone(dt);
        }
        private void FormatPreviousService(DataTable dt)
        {
            var reg = new Regex("^[A-Za-z\\d]*$");
           
            if(dt.Columns.Contains(BillSheetColumns.PREVIOUS_SERVICE))
            {
                var services = mMembers.Where(a => a.Position.Equals(PositionNames.SERVICE)).ToArray();

                foreach(DataRow dr in dt.Rows)
                { 
                    var value = dr[BillSheetColumns.PREVIOUS_SERVICE] as string;
                    if(!string.IsNullOrEmpty(value))
                    {
                        value = value.Trim();
                        if(reg.IsMatch(value))
                        {
                            var mem = services.FirstOrDefault(a =>
                                a.ID.Equals(value, StringComparison.CurrentCultureIgnoreCase));
                            if(mem != null)
                            {
                                dr[BillSheetColumns.PREVIOUS_SERVICE_ID] = value;
                                dr[BillSheetColumns.PREVIOUS_SERVICE] = mem.Name;
                            }
                            else
                            {
                                mLogger.Warn($"未匹配到客服专员：{value}");
                            }
                        }
                        else
                        {
                            var mems = services.Where(a =>
                                a.Name.Equals(value, StringComparison.CurrentCultureIgnoreCase)).ToArray();
                            if(mems.Length == 1)
                            {
                                var mem = mems.FirstOrDefault();
                                dr[BillSheetColumns.PREVIOUS_SERVICE_ID] = mem.ID; 
                            }

                            if(mems.Length == 0)
                            {
                                mLogger.Warn($"未匹配到客服专员：{value}");
                            }

                            if(mems.Length > 1)
                            {
                                mLogger.Warn($"存在两个以上相同的营销员：{value}， 未能补足客服专员信息");
                            }
                        }
                    } 
                }
            }
        }
        private void FormatMobilePhone(DataTable dt)
        {
            if(dt.Columns.Contains(BillSheetColumns.MOBILE_PHONE))
            {
                foreach (DataRow dr in dt.Rows)
                {
                    var phone = dr[BillSheetColumns.MOBILE_PHONE] as string;
                    if(!string.IsNullOrEmpty(phone))
                    {
                        phone = phone.Replace("M:", "");
                        phone = phone.Replace("m:", "" );
                        dr[BillSheetColumns.MOBILE_PHONE] = phone;
                    }
                }
            }

        }
        private void FormatAddress(DataTable dt)
        { 
            FormatAddressColumn(dt, BillSheetColumns.PAY_ADDRESS);
            FormatAddressColumn(dt, BillSheetColumns.PAY_ADDRESS2);
            FormatAddressColumn(dt, BillSheetColumns.PAY_ADDRESS3);
            FormatAddressColumn(dt, BillSheetColumns.PAY_ADDRESS4);
            FormatAddressColumn(dt, BillSheetColumns.CUSTOMER_ADDRESS);  
        }
        private void FormatAddressColumn(DataTable dt, string  colName)
        {
             if (dt.Columns.Contains(colName))
            {
                foreach (DataRow dr in dt.Rows)
                {
                    dr[colName] = FormatAddress(dr[colName]);
                }
            }
        }
        private string FormatAddress(object colValue)
        {
            var col = colValue as string;
            if(string.IsNullOrEmpty(col))
                return "";
            var v = col.Replace("浙江省金华市", "");
            v = v.Replace("金华市婺城区", "婺城区");
            v = v.Replace("金华市金东区", "金东区");
            v = v.Replace("浙江省永康市", "永康市");
            v = v.Replace("浙江省武义县", "武义县");
            v = v.Replace("浙江省浦江县", "浦江县");
            v = v.Replace("浙江省义乌市", "义乌市");
            v = v.Replace("浙江省兰溪市", "兰溪市");
            v = v.Replace("浙江省磐安县", "磐安县");
            v = v.Replace("浙江省东阳市", "东阳市");

            return v;
        }

        private void FormatValidColumns(DataTable dt)
        {
            foreach(DataRow dr in dt.Rows)
            {
                if(dt.Columns.Contains(BillSheetColumns.CUSTOMER_PASSPORT_ID))
                {
                    dr[BillSheetColumns.CUSTOMER_PASSPORT_ID] = ClearVaildChars(dr[BillSheetColumns.CUSTOMER_PASSPORT_ID]);
                }

                if (dt.Columns.Contains(BillSheetColumns.CUSTOMER_PASSPORT_ID2))
                {
                    dr[BillSheetColumns.CUSTOMER_PASSPORT_ID2] = ClearVaildChars(dr[BillSheetColumns.CUSTOMER_PASSPORT_ID2]);
                }

                if (dt.Columns.Contains(BillSheetColumns.BILL_ID))
                {
                    dr[BillSheetColumns.BILL_ID] = ClearVaildChars(dr[BillSheetColumns.BILL_ID]);
                }
                if (dt.Columns.Contains(BillSheetColumns.BILL_ID2))
                {
                    dr[BillSheetColumns.BILL_ID2] = ClearVaildChars(dr[BillSheetColumns.BILL_ID2]);
                }
            }
        }

        private string ClearVaildChars(object val)
        {
            var value = val as string;
            if(string.IsNullOrEmpty(value))
            {
                return value;
            }

            value = value.Replace("[", "");
            value = value.Replace("]", "");
            return value;
        }
    }
}
