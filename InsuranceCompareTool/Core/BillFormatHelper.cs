using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using InsuranceCompareTool.Domain;
using InsuranceCompareTool.Models;
using log4net;
namespace InsuranceCompareTool.Core
{
    public class BillFormatHelper
    {
        Regex mIDReg = new Regex("^[A-Za-z\\d]*$");
        private readonly ILog mLogger;
        private readonly List<Member> mMembers;
        private Dictionary<string,bool> mNonSellerId = new Dictionary<string, bool>();
        private Dictionary<string,bool> mNonServiceId = new Dictionary<string, bool>();

        public BillFormatHelper(ILog logger, List<Member> members)
        { 
            mLogger = logger;
            mMembers = members;
        }
        public void Format(DataTable dt)
        {
            AttachRequiredColumns(dt);
            FormatStringColumns(dt); 
            FormatAddress(dt);
            FormatValidColumns(dt);
            FormatMobilePhone(dt);
            AttachColumnServiceStatus(dt); 
            AttachColumnPreviousServiceID(dt);
            AttachColumnReAssign(dt);
            AttachColumnClientArea(dt);
            FormatPreviousService(dt);
            FormatClientArea(dt);
            FormatReAssign(dt);
            FormatSysService(dt); 
            FormatCurrentServiceID(dt); 
            FormatSellerArea(dt);
            FormatServiceArea(dt);
            FormatGuid(dt);

            foreach(KeyValuePair<string, bool>  item in mNonSellerId)
            {
                mLogger.Error($"未找到工号为 {item.Key} 的营销员");
            }


            foreach (KeyValuePair<string, bool> item in mNonServiceId)
            {
                mLogger.Error($"未找到工号为 {item.Key} 的客服专员");
            }
        }

        private void FormatGuid(DataTable dt)
        {
            foreach(DataRow dr in dt.Rows)
            {
                dr[BillSheetColumns.SYS_GUID] = Guid.NewGuid();
            }
        }
        private void FormatSellerArea(DataTable dt)
        {
            foreach (DataRow dr in dt.Rows)
            {
                if (dr.IsNull(BillSheetColumns.SELLER_ID))
                    continue;
                string serId = (string) dr[BillSheetColumns.SELLER_ID];
                var mem = mMembers.FirstOrDefault(  a => a.ID.Equals(serId));
                if (mem != null)
                {
                    dr[BillSheetColumns.SYS_SELLER_AREA] = mem.Area;
                }
                else
                {
                    mNonSellerId[serId] = true;
                }
            }
        }

        private void FormatServiceArea(DataTable dt)
        {
            foreach (DataRow dr in dt.Rows)
            {
                if(dr.IsNull(BillSheetColumns.CURRENT_SERVICE_ID))
                    continue;
                string sellerId = (string) dr[BillSheetColumns.CURRENT_SERVICE_ID];
                var mem = mMembers.FirstOrDefault(a => a.ID.Equals(sellerId));
                if (mem != null)
                {
                    dr[BillSheetColumns.SYS_SERVICE_AREA] = mem.Area;
                }
                else
                {
                    mNonServiceId[sellerId] = true; 
                }
            }
        }

 
        private void FormatStringColumns(DataTable dt)
        {
            foreach(DataRow dr in dt.Rows)
            {
                foreach(DataColumn col in dt.Columns)
                {
                    if(col.DataType == typeof(string))
                    {
                        if(!dr.IsNull(col))
                        {
                            dr[col] = ((string)dr[col]).Trim().Replace("\"", "").Replace("[","").Replace("]",""); 
                        }
                    }
                }
            }
        }

        private void FormatCurrentServiceID(DataTable dt)
        {
            foreach (DataRow dr in dt.Rows)
            {
                var serviceName = "";
                var serviceId = "";
                if (!dr.IsNull(BillSheetColumns.CURRENT_SERVICE_NAME))
                {
                    serviceName = Convert.ToString(dr[BillSheetColumns.CURRENT_SERVICE_NAME]);
                }

                if (!dr.IsNull(BillSheetColumns.CURRENT_SERVICE_ID))
                {
                    serviceId = Convert.ToString(dr[BillSheetColumns.CURRENT_SERVICE_ID]);
                }

                if (!string.IsNullOrEmpty(serviceName) && string.IsNullOrEmpty(serviceId))
                {
                    var service = mMembers.Where(a =>
                        a.Name.Equals(serviceName, StringComparison.CurrentCultureIgnoreCase) && a.Position.Equals(PositionNames.SERVICE )).ToArray();
                    if(service.Length == 0)
                    {
                        mLogger.Warn($"未找到服务员：{serviceName}");
                        continue;
                    }

                    if(service.Length > 1)
                    {
                        mLogger.Warn($"匹配到多名服务员：{serviceName}, 需手工选择");
                        continue;
                    }  
                    dr[BillSheetColumns.CURRENT_SERVICE_ID] = service[0].ID;
                }
            }
        }
 

   
        private void FormatSysService(DataTable dt)
        {
            foreach(DataRow dr in dt.Rows)
            {
                var serviceName = "";
                var serviceId = "";
                if(!dr.IsNull(BillSheetColumns.CURRENT_SERVICE_NAME))
                {
                    serviceName = Convert.ToString( dr[BillSheetColumns.CURRENT_SERVICE_NAME]);
                }

                if(!dr.IsNull(BillSheetColumns.CURRENT_SERVICE_ID))
                {
                    serviceId = Convert.ToString(dr[BillSheetColumns.CURRENT_SERVICE_ID]);
                }

                if(!string.IsNullOrEmpty(serviceName) || !string.IsNullOrEmpty(serviceId))
                {
                    var service = $"{serviceName} ({serviceId})";
                    dr[BillSheetColumns.SYS_SERVICE] = service;
                }

                dr[BillSheetColumns.SYS_SERVICE_ID] = serviceId;

            }
        }

        private void FormatReAssign(DataTable dt)
        {
            if(!dt.Columns.Contains( BillSheetColumns.ASSIGNED_SERVICE_NAME))
                return;
            foreach (DataRow dr in dt.Rows)
            { 
                if(dr.IsNull( BillSheetColumns.ASSIGNED_SERVICE_NAME))
                    continue;
                var value =Convert.ToString(  dr[BillSheetColumns.ASSIGNED_SERVICE_NAME]);

                if (!string.IsNullOrEmpty(value))
                {
                    value = value.Trim();
                    
                    {
                        var mems = mMembers.Where(a =>
                            a.Name.Equals(value, StringComparison.CurrentCultureIgnoreCase)).ToArray();
                        if (mems.Length == 1)
                        {
                            var mem = mems.FirstOrDefault();
                            dr[BillSheetColumns.ASSIGNED_SERVICE_ID] = mem.ID;
                        }

                        if (mems.Length == 0)
                        {
                            mLogger.Warn($"未匹配到改派后专员：{value}");
                        }

                        if (mems.Length > 1)
                        {
                            mLogger.Warn($"存在两个以上相同的改派后专员：{value}， 未能补足改派后专员工号");
                        }
                    }
                }
            }

        }
        private void AttachColumnReAssign(DataTable dt)
        {
            var col = BillTableColumns.COL_ASSIGNED_SERVICE_ID;

            if(dt.Columns.Contains(BillSheetColumns.ASSIGNED_SERVICE_NAME))
            {
                if(!dt.Columns.Contains(col.FormalName))
                {
                    dt.Columns.Add(col.FormalName, col.Type);
                }
            }
        }
 

        private void AttachRequiredColumns(DataTable dt)
        {
            var sysCols = BillTableColumns.Columns.Where(a => a.IsRequired).ToArray();

            foreach (var column in sysCols)
            {
                if (!dt.Columns.Contains(column.FormalName))
                {
                    dt.Columns.Add(column.FormalName, column.Type);
                }
            }
        }

        private void AttachColumnClientArea(DataTable dt)
        {
            var col = BillTableColumns.COL_CUSTOMER_AREA;
            if (dt.Columns.Contains(col.FormalName))
            {
                return;
            }
            dt.Columns.Add(col.FormalName, col.Type);
        }
        private void AttachColumnServiceStatus(DataTable dt)
        {
            var col = BillTableColumns.COL_SERVICE_STATUS;
            if(dt.Columns.Contains(col.FormalName))
            {
                return;
            }
             
            dt.Columns.Add(col.FormalName, col.Type);
            foreach(DataRow row in dt.Rows)
            {
                if(!row.IsNull(BillSheetColumns.CURRENT_SERVICE_ID))
                {
                    var sellerID = (string) row[BillSheetColumns.CURRENT_SERVICE_ID];
                    var mem = mMembers.Find(a => a.ID.Equals(sellerID));
                    row[col.Name.First()] = mem?.Status;
                }
            }
        }

        

        private void AttachColumnPreviousServiceID(DataTable dt)
        {
            if (dt.Columns.Contains(BillSheetColumns.PREVIOUS_SERVICE))
            {
                if (!dt.Columns.Contains(BillSheetColumns.PREVIOUS_SERVICE_ID))
                {
                    var col = BillTableColumns.COL_PREVIOUS_SERVICE_ID;
                    dt.Columns.Add(new DataColumn(col.Name.First(),col.Type));
                }
            }
        }

        private void FormatPreviousService(DataTable dt)
        {
            
           
            if(dt.Columns.Contains(BillSheetColumns.PREVIOUS_SERVICE))
            {
                var services = mMembers.Where(a => a.Position.Equals(PositionNames.SERVICE)).ToArray();

                foreach(DataRow dr in dt.Rows)
                { 
                    var value = dr[BillSheetColumns.PREVIOUS_SERVICE] as string;
                    if(!string.IsNullOrEmpty(value))
                    {
                        value = value.Trim();
                        if(mIDReg.IsMatch(value))
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

        private void FormatClientArea(DataTable dt)
        {
            foreach(DataRow dr in dt.Rows)
            {
                var area = GetClientArea(dr);
                dr[BillSheetColumns.CLIENT_AREA] = area;
            }
        }

        private string GetClientArea(DataRow bill)
        {
            var cols = new List<string>();
            foreach (DataColumn column in bill.Table.Columns)
            {
                if (column.ColumnName.Contains("地址"))
                {
                    cols.Add(column.ColumnName);
                }
            }
            if (cols.Count <= 0)
                return "";
            var address = "";
            foreach (var col in cols)
            {
                if (!bill.IsNull(col))
                {
                    address += (string)bill[col];
                    continue;
                }
            }

            if (string.IsNullOrEmpty(address))
            {
                return "";
            }

            var areas = GetAreas(address);
            if (areas.Count <= 0)
                return "";
            var area = areas.FirstOrDefault(a => a != AreaNames.JH);
            if (string.IsNullOrEmpty(area))
            {
                if(areas.Contains(AreaNames.JH))
                {
                    area = AreaNames.JH;
                }
            }
            return area;
        }
        private List<string> GetAreas(string address)
        {
            string[] mJhAddress = new string[] { "永康街", "义乌街", "浦江街", "兰溪街", "东阳街", "武义街", "磐安街", "婺城区", "金东区" };
            foreach(var jh in mJhAddress)
            {
                address = address.Replace(jh, AreaNames.JH );
            }
            var list = new List<string>();
            foreach (var area in AreaNames.Areas)
            {
                if (address.Contains(area))
                {
                    list.Add(area);
                }
            }
            return list;
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
