using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InsuranceCompareTool.Domain;
using InsuranceCompareTool.Models;
using InsuranceCompareTool.Models.Dispatch;
using InsuranceCompareTool.ShareCommon;
using log4net;
using NPOI.SS.Formula.Functions;
namespace InsuranceCompareTool.Core
{

    /// <summary>
    /// 指派订单类
    /// </summary>
    public class BillAssignHelper
    {
        private readonly ILog mLogger;
        private readonly List<Member> mMembers;
          
        private const string TARGET_SERVICE_ID = BillSheetColumns.ASSIGNED_SERVICE_ID;
        private const string TARGET_SERVICE_NAME = BillSheetColumns.ASSIGNED_SERVICE_NAME;

        public BillAssignHelper(ILog logger, List<Member> members )
        {
            mLogger = logger;
            mMembers = members;
        }
 
        private void UpdateDataTable(DataTable table)
        { 
            foreach (DataRow row in table.Rows)
            {
                row[BillSheetColumns.SYS_FINISHED] = false;
                row[BillSheetColumns.SYS_FILTER] = "";
                row[BillSheetColumns.SYS_FILTER] = "";
                row[BillSheetColumns.SYS_ERROR] = "";
            }
        }
        public void AssignBills(DataTable dt, Project project )
        {
            UpdateDataTable(dt);

            foreach (var step in project.Steps)
            {
                if(step.Filters.Any(a => a.Column == LogicColumnName.DifferentService))
                {
                    HandleDifferentServiceBills(dt, step);
                }
                else
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        var hit = GetHitValue(row, step.Filters);

                        if (hit)
                        {
                            if(step.StopNextStep)
                            {
                                row[BillSheetColumns.SYS_FINISHED] = true;
                            }
                            AssignTo(row, step);
                        }
                    }                    
                }
            }
        }
        /// <summary>
        /// 订单命中测试
        /// </summary>  
        private bool GetHitValue(DataRow bill, IEnumerable<Filter> filters)
        {
            var finished =  Convert.ToBoolean(bill[BillSheetColumns.SYS_FINISHED]); 

            if (finished == true)
                return false;
            var hit = new List<bool>();

            foreach (var filter in filters)
            {
                switch (filter.Column)
                { 
                    case LogicColumnName.ServiceStatus:
                    { 
                        hit.Add(GetHitStringValue(BillSheetColumns.SYS_SERVICE_STATUS, filter, bill));
                        break;
                    }
                    case LogicColumnName.ServiceID:
                    {
                        hit.Add(GetHitStringValue(BillSheetColumns.CURRENT_SERVICE_ID, filter, bill));
                        break;
                    }
                    case LogicColumnName.BillID:
                    {
                        hit.Add(GetHitStringValue(BillSheetColumns.BILL_ID, filter, bill));
                        break;
                    }
                    case LogicColumnName.PreviousServiceID:
                    {
                        hit.Add(GetHitStringValue(BillSheetColumns.PREVIOUS_SERVICE_ID, filter, bill));
                        break;
                    }
                    case LogicColumnName.CurrentPrice:
                    {

                        var b1 = GetHitNumberValue(BillSheetColumns.BILL_PRICE, filter, bill);
                        var b2 = GetHitNumberValue(BillSheetColumns.BILL_PRICE2, filter, bill);
                        hit.Add(b1 || b2);
                        break;
                    }
                    case LogicColumnName.SellerID:
                    {
                        hit.Add(GetHitStringValue(BillSheetColumns.SELLER_ID, filter, bill));
                        break;
                    }
                    case LogicColumnName.SellerStatus:
                    {
                        hit.Add(GetHitStringValue(BillSheetColumns.SELLER_STATE, filter, bill));
                        break;
                    }
                    case LogicColumnName.ProductID:
                    {
                        hit.Add(GetHitStringValue(BillSheetColumns.PRODUCT_ID, filter, bill));
                        break;
                    }
                    case LogicColumnName.ProductName:
                    {
                        hit.Add(GetHitStringValue(BillSheetColumns.PRODUCT_NAME, filter, bill)); 
                        break;
                    }
                    case LogicColumnName.PayAddress:
                    {
                        var a1 = GetHitStringValue(BillSheetColumns.PAY_ADDRESS, filter, bill);
                        var a2 = GetHitStringValue(BillSheetColumns.PAY_ADDRESS2, filter, bill);
                        var a3 = GetHitStringValue(BillSheetColumns.PAY_ADDRESS3, filter, bill);
                        var a4 = GetHitStringValue(BillSheetColumns.PAY_ADDRESS4, filter, bill);
                        var a5 = GetHitStringValue(BillSheetColumns.PAY_ADDRESS5, filter, bill);
                        hit.Add( a1 || a2 || a3 ||a4 || a5);
                        break;
                    }
                    case LogicColumnName.PayNo:
                    {
                        hit.Add(GetHitIntValue(BillSheetColumns.PAY_NO, filter,bill));
                        break;
                    }
                    case LogicColumnName.DifferentArea:
                    {
                        hit.Add(GetHitByDifferentArea(filter, bill));
                        break;
                    }
                    case LogicColumnName.SellerIsService:
                    {
                        hit.Add(GetHitBySellerIsService(filter,bill));
                        break;
                    }
                    case LogicColumnName.CustomerPassportId:
                    {
                        var b1 = GetHitStringValue(BillSheetColumns.CUSTOMER_PASSPORT_ID, filter, bill);
                        var b2 = GetHitStringValue(BillSheetColumns.CUSTOMER_PASSPORT_ID2, filter, bill);
                        hit.Add(b1 || b2);
                        break;
                    }
 
                    default:
                    {
                        hit.Add(true);
                        break;
                    }
                }
            }

            bool hited = false;
            var fs = filters.ToArray();
            for (var i = 0; i < fs.Length; i++)
            {
                Filter f = fs[i];
                if(i == 0)
                {
                    hited = hit[i];
                }
                else
                {
                    if (f.BoolType == BoolType.And)
                    {
                        hited = hited && hit[i]; 
                    }

                    if (f.BoolType == BoolType.Or)
                    { 
                        hited = hited || hit[i];
                    }                   
                }

            }
             

            return hited;
        }



        private bool GetHitIntValue(string col, Filter filter, DataRow bill)
        {
            bool b = false;
            if (!bill.Table.Columns.Contains(col))
                return b;
            switch (filter.OperatorType)
            {
                case OperatorType.Equal:
                {
                    b = GetHitByIntEqual(col, Convert.ToInt32(filter.Value), bill);
                    break;
                }
                case OperatorType.NoEqual:
                {
                    b = !GetHitByIntEqual(col, Convert.ToInt32(filter.Value), bill);
                    break;
                }
                case OperatorType.Greater:
                {
                    b = GetHitByIntGreater(col, Convert.ToInt32(filter.Value), bill);
                    break;
                }
                case OperatorType.Less:
                {
                    b = GetHitByIntLess(col, Convert.ToInt32(filter.Value), bill);
                    break;
                }
                case OperatorType.GreaterAndEqual:
                {
                    b = GetHitByIntGreaterAndEqual(col, Convert.ToInt32(filter.Value), bill);
                    break;
                }
                case OperatorType.LessAndEqual:
                {
                    b = GetHitByIntLessAndEqual(col, Convert.ToInt32(filter.Value), bill);
                    break;
                }
            }
            return b;
        }

        private bool GetHitNumberValue(string col , Filter filter , DataRow bill)
        {
            bool b = false;
            if(!bill.Table.Columns.Contains(col))
                return b;
            switch(filter.OperatorType)
            {
                case OperatorType.Equal:
                {
                    b = GetHitByNumberEqual(col, Convert.ToDouble(filter.Value), bill);
                    break;
                }
                case OperatorType.NoEqual:
                {
                    b = !GetHitByNumberEqual(col, Convert.ToDouble(filter.Value), bill);
                    break;
                }
                case OperatorType.Greater:
                {
                    b = GetHitByNumberGreater(col, Convert.ToDouble(filter.Value), bill);
                   break;
                }
                case OperatorType.Less:
                {
                    b = GetHitByNumberLess(col, Convert.ToDouble(filter.Value), bill);
                    break;
                }
                case OperatorType.GreaterAndEqual:
                {
                    b = GetHitByNumberGreaterAndEqual(col, Convert.ToDouble(filter.Value), bill);
                    break;
                }
                case OperatorType.LessAndEqual:
                {
                    b = GetHitByNumberLessAndEqual(col, Convert.ToDouble(filter.Value), bill);
                    break;
                }
            }
            return b;
        }
        private bool GetHitStringValue(string col, Filter filter, DataRow bill)
        {
            bool b = false;
            if (!bill.Table.Columns.Contains(col))
                return b;
            switch (filter.OperatorType)
            {
                case OperatorType.Equal:
                {
                   
                    b = GetHitByStringEqual(col, filter.Value, bill);
                    break;
                }
                case OperatorType.NoEqual:
                {
                    b = GetHitByStringNoEqual(col, filter.Value, bill);
                    break; 
                }
                case OperatorType.Contain:
                {
                    b = GetHitByStringContains(col, filter.Value, bill);
                    break;
                }

            }
            return b;
        }


        private bool GetHitByNumberLessAndEqual(string col, double val, DataRow bill)
        {
            bool b = false;
            var value = Convert.ToDouble( bill[col]);
            b = value <= val;
            return b;
        }

        private bool GetHitByNumberGreaterAndEqual(string col, double val, DataRow bill)
        {
            bool b = false;
            var value = Convert.ToDouble( bill[col]);
            b = value >= val;
            return b;
        }


        private bool GetHitByNumberLess(string col, double val, DataRow bill)
        {
            bool b = false;
            var value = Convert.ToDouble( bill[col]);
            b = value < val;
            return b;
        }

        private bool GetHitByNumberGreater(string col, double val, DataRow bill)
        {
            bool b = false;
            var value = Convert.ToDouble( bill[col]);
            b = value > val;
            return b;
        }

        private bool GetHitByIntLess(string col, int val, DataRow bill)
        {
            bool b = false;
            var value = (int)bill[col];
            b = value < val;
            return b;
        }
        private bool GetHitByIntGreater(string col, int val, DataRow bill)
        {
            bool b = false;
             
            var value = Convert.ToInt32( bill[col]);
            b = value > val;
            return b;
        }
        private bool GetHitByIntLessAndEqual(string col, int val, DataRow bill)
        {
            bool b = false;
            var value = (int)bill[col];
            b = value <= val;
            return b;
        }
        private bool GetHitByIntGreaterAndEqual(string col, int val, DataRow bill)
        {
            bool b = false;
            var value = (int)bill[col];
            b = value >= val;
            return b;
        }

        private bool GetHitByIntEqual(string col, int val, DataRow bill)
        {
            bool b = false;
            if (!bill.IsNull(col))
            {
                var value = (int)bill[col];
                b = value.Equals(val);
            }
            return b;
        }

        private bool GetHitByNumberEqual(string col, double val, DataRow bill)
        {
            bool b = false;
            if (!bill.IsNull(col))
            {
                var value = (double)bill[col];
                b = value.Equals(val);
            }
            return b;
        }

        private bool GetHitByStringEqual(string col,string val, DataRow bill)
        {
            bool b = false;
            if(!bill.IsNull(col))
            {
                var value = (string) bill[col];
                b = value.Equals(val, StringComparison.CurrentCultureIgnoreCase);
            }
            else
            {
                b = string.IsNullOrEmpty(val);
            }
            return b;
        }

        private bool GetHitByStringNoEqual(string col, string val, DataRow bill)
        {
            return !GetHitByStringEqual(col, val, bill);
        }

        private bool GetHitByStringContains(string col, string val, DataRow bill)
        {
            bool b = false;
            if(bill.Table.Columns.Contains(col))
            {
                if (!bill.IsNull(col))
                {
                    var value = (string)bill[col];
                    if(value.ToUpper().Contains(val.ToUpper()))
                    {
                        b = true;
                    }
                    else
                    {
                        b = false;
                    }
                } 
            }

            return b;
        }

        private void HandleDifferentServiceBills(DataTable dt , Step step)
        {
            string cprCol = "";
            if(dt.Columns.Contains(BillSheetColumns.CUSTOMER_PASSPORT_ID))
            {
                cprCol = BillSheetColumns.CUSTOMER_PASSPORT_ID;
            }

            if(dt.Columns.Contains(BillSheetColumns.CUSTOMER_PASSPORT_ID2))
            {
                cprCol = BillSheetColumns.CUSTOMER_PASSPORT_ID2;
            }
            Dictionary<string,List<MemIndex>> list = new Dictionary<string, List<MemIndex>>();
            
            for (var i = 0;i<dt.Rows.Count;i++)
            {
                var row = dt.Rows[i];
                if(!row.IsNull(cprCol))
                {
                    var passport = (string)row[cprCol];
                    var serviceID = (string)row[BillSheetColumns.CURRENT_SERVICE_ID];
                    var mem = new MemIndex(){ Index =  i, ServiceID = serviceID};
                    if(list.ContainsKey(passport))
                    {
                        var items = list[passport];
                        items.Add(mem);
                    }
                    else
                    {
                        var items = new List<MemIndex>();
                        items.Add(mem);
                        list.Add(passport,items);
                    }
                }
            }



            foreach(KeyValuePair<string, List<MemIndex>> item in list)
            {
                string lastService = "";
                bool hasDifferentService = false;
                if(item.Value.Count <= 1)
                    continue;
                foreach(var m in item.Value)
                {
                    if(!string.IsNullOrEmpty(lastService))
                    {
                        if(!lastService.Equals(m.ServiceID, StringComparison.CurrentCultureIgnoreCase))
                        {
                            hasDifferentService = true;
                            break;
                        }
                    } 
                    lastService = m.ServiceID; 
                }

                if(hasDifferentService)
                {
                    foreach(var m in item.Value)
                    {
                        var dr = dt.Rows[m.Index];
                        dr[BillSheetColumns.SYS_FILTER] = step.Title;
                    }

                    var targetServiceId = "";

                    if(step.DispatchType == DispatchType.DispatchToSmallService)
                    { 
                        var member = GetSmallServiceID(item.Value,dt );

                        foreach (var m2 in item.Value)
                        {
                            var dr = dt.Rows[m2.Index];
                            dr[BillSheetColumns.CURRENT_SERVICE_ID] = member.ID;
                            dr[BillSheetColumns.CURRENT_SERVICE_NAME] = member.Name;
                            dr[BillSheetColumns.SYS_SERVICE_AREA] = member.Area; 
                            LogHistory(dr, step, member);
                        }
                    }

                    if(step.DispatchType == DispatchType.DispatchToLargeService)
                    {
                        var member = GetLargeServiceID(item.Value, dt); 
                        foreach (var m2 in item.Value)
                        {
                            var dr = dt.Rows[m2.Index];
                            dr[BillSheetColumns.CURRENT_SERVICE_ID] = member.ID;
                            dr[BillSheetColumns.CURRENT_SERVICE_NAME] = member.Name;
                            dr[BillSheetColumns.SYS_SERVICE_AREA] = member.Area;
                            LogHistory(dr, step, member);
                        } 
                    } 
                } 
            }
        }

        private Member GetLargeServiceID(List<MemIndex> mems,DataTable dt)
        {
            var ms = new List<Member>();
            foreach (var mem in mems)
            {
                var m = mMembers.FirstOrDefault(a =>
                    a.ID.Equals(mem.ServiceID, StringComparison.CurrentCultureIgnoreCase));
                m.BillCount = 0;
                ms.Add(m);
                foreach(DataRow dr in dt.Rows)
                {
                    if(!dr.IsNull(BillSheetColumns.CURRENT_SERVICE_ID))
                    {
                        var rowSerId = (string)dr[BillSheetColumns.CURRENT_SERVICE_ID];
                        if(rowSerId.Equals(mem.ServiceID, StringComparison.CurrentCultureIgnoreCase))
                        {
                            m.BillCount += 1;
                        }
                    }
                }
            }

            return ms.OrderByDescending(t => t.BillCount).FirstOrDefault();

        }

        private Member GetSmallServiceID(List<MemIndex> mems,DataTable dt)
        {
            var ms = new List<Member>();
            foreach (var mem in mems)
            {
                var m = mMembers.FirstOrDefault(a =>
                    a.ID.Equals(mem.ServiceID, StringComparison.CurrentCultureIgnoreCase));
                m.BillCount = 0;
                ms.Add(m);
                foreach (DataRow dr in dt.Rows)
                {
                    if (!dr.IsNull(BillSheetColumns.CURRENT_SERVICE_ID))
                    {
                        var rowSerId = (string)dr[BillSheetColumns.CURRENT_SERVICE_ID];
                        if (rowSerId.Equals(mem.ServiceID, StringComparison.CurrentCultureIgnoreCase))
                        {
                            m.BillCount += 1;
                        }
                    }
                }

            }
            return ms.OrderBy(t => t.BillCount).FirstOrDefault();

        }
        
        private void AssignTo(DataRow bill, Step step)
        {
            switch(step.DispatchType)
            {
                case DispatchType.DispatchToDesignated:
                {
                    AssignToSomebody(bill,step);
                    break;
                }
                case DispatchType.DispatchToManagerOfSeller:
                {
                    AssignToManagerOfSeller(bill,step);
                    break;
                }
                case DispatchType.DispatchToManagerOfService:
                {
                    AssignToManagerOfService(bill,step);
                    break;
                }
                case DispatchType.DispatchToManagerOfCustomer:
                {
                    AssignToManagerOfCustomer(bill, step);
                    break;
                }
                case DispatchType.DispatchToPreviousService:
                {
                    AssignToPreviousService(bill, step);
                    break;
                }
                case DispatchType.DispatchToSeller:
                {
                    AssignToSeller(bill, step);
                    break;
                }
                case DispatchType.DoNot:
                {
                    AssignToNot(bill, step);
                    break;
                }
            }
        }
        private void AssignToNot(DataRow bill, Step step)
        {
            bill[BillSheetColumns.SYS_FILTER] = step.Title; 
        }

        private void AssignToManagerOfCustomer(DataRow bill, Step step)
        {
            var bid = GetBillID(bill);
            var area = (string) bill[BillSheetColumns.CLIENT_AREA];
            var mem = mMembers.FirstOrDefault(a => a.Area.Equals(area) && a.Position == PositionNames.MUSTER);
            if(mem == null)
            {
                LogError(bill, $"{step.Title}: 未能找到 {area} 的主管，无法分配给主管. (保单号： {bid})"); 
            }
            else
            { 
                AssignTo(bill,mem); 
                LogHistory(bill, step, mem);

            }
            bill[BillSheetColumns.SYS_FILTER] = step.Title;
        }
        /// <summary>
        /// 指派订单给某人
        /// </summary> 
        /// <param name="step"></param>
        private void AssignToSomebody(DataRow row, Step step)
        {
            var bid = GetBillID(row);
            if(string.IsNullOrEmpty(step.DispatchDesignated))
            {
                LogError(row, $"{step.Title}: 未能找到工号为 空 的员工，无法分配给指定人员. (保单号： {bid})");
            }
            else
            {
                var mem = mMembers.FirstOrDefault(a => a.ID.Equals(step.DispatchDesignated.Trim(), StringComparison.CurrentCultureIgnoreCase));
                if(mem == null)
                {
                    
                    LogError(row,$"{step.Title}: 未能找到工号为 {step.DispatchDesignated} 的员工，无法分配给指定人员. (保单号： {bid})");
                }
                else
                { 
                    AssignTo(row, mem);
                    LogHistory(row, step, mem);

                }
            }
            

            row[BillSheetColumns.SYS_FILTER] = step.Title;
        }
        /// <summary>
        /// 分配给客服-主管
        /// </summary>
        private void AssignToManagerOfService(DataRow row, Step step )
        {
            var bid = GetBillID(row);
            if(row.IsNull(BillSheetColumns.CURRENT_SERVICE_ID))
            {
                
                LogError(row,$"{step.Title}: 客服工号空缺，无法分配给主管. (保单号： {bid})");
            }
            else
            {
                var sid =(string) row[BillSheetColumns.CURRENT_SERVICE_ID];
                var mem = mMembers.FirstOrDefault(a =>
                    a.ID.Equals(sid.Trim(), StringComparison.CurrentCultureIgnoreCase));
                if(mem == null)
                {
                    LogError(row, $"{step.Title}: 未能找到工号为 {sid} 的客服，无法分配给主管. (保单号： {bid})");
                }
                else
                {
                    var manager = mMembers.FirstOrDefault(a => a.Area.Equals(mem.Area) && a.Position == PositionNames.MUSTER);
                    if(manager != null)
                    {
                        AssignTo(row, manager);
                        LogHistory(row, step, manager);

                    }
                    else
                    {
                        LogError(row, $"{step.Title}: 未能找到 {mem.Area} 的主管，无法分配给主管. (保单号： {bid})");
                    }
                }
            } 
            row[BillSheetColumns.SYS_FILTER] = step.Title; 
        }

        /// <summary>
        /// 分配给客服-主管
        /// </summary>
        private void AssignToManagerOfSeller(DataRow row, Step step)
        {
            var bid = GetBillID(row);
            if (row.IsNull(BillSheetColumns.SELLER_ID))
            { 
                LogError(row, $"{step.Title}: 营销员工号空缺，无法分配给主管. (保单号： {bid})");
            }
            else
            {
                var sid = (string)row[BillSheetColumns.SELLER_ID];
                var mem = mMembers.FirstOrDefault(a =>
                    a.ID.Equals(sid.Trim(), StringComparison.CurrentCultureIgnoreCase));
                if (mem == null)
                {
                    LogError(row, $"{step.Title}: 未能找到工号为 {sid} 的营销员，无法分配给主管. (保单号： {bid})");
                }
                else
                {
                    var manager = mMembers.FirstOrDefault(a => a.Area.Equals(mem.Area) && a.Position == PositionNames.MUSTER);
                    if (manager != null)
                    { 
                        AssignTo(row, manager);
                        LogHistory(row, step, manager); 
                    }
                    else
                    {
                        LogError(row, $"{step.Title}: 未能找到 {mem.Area} 的主管，无法分配给主管. (保单号： {bid})");
                    }
                }
            }
            row[BillSheetColumns.SYS_FILTER] = step.Title; 
        }
        /// <summary>
        /// 分配给上期客服专员
        /// </summary>
        /// <param name="bill"></param>
        /// <param name="step"></param>
        private void AssignToPreviousService(DataRow bill, Step step)
        {
            var bid = GetBillID(bill);
            if (bill.IsNull(BillSheetColumns.PREVIOUS_SERVICE_ID))
            {
                LogError(bill,$"{step.Title}: 上期客服专员工号为空，无法分配给上期客服专员.(保单号: {bid}) "); 
            }
            else
            {
                 var sid = Convert.ToString(bill[BillSheetColumns.PREVIOUS_SERVICE_ID]).Trim();
                var mem = mMembers.FirstOrDefault(a => a.ID.Equals(sid, StringComparison.CurrentCultureIgnoreCase));
                if (mem == null)
                { 
                    LogError(bill, $"{step.Title}: 未能找到工号为 {sid} 的客服专员，无法分配. (保单号： {bid})");
                }
                else
                {
                    AssignTo(bill, mem);
                    LogHistory(bill, step, mem);
                    
                }
            }

            bill[BillSheetColumns.SYS_FILTER] = step.Title;
        }

        private void AssignToSeller(DataRow bill, Step step)
        {
            var bid = GetBillID(bill);
            if(bill.IsNull(BillSheetColumns.SELLER_ID))
            {
                LogError(bill, $"{step.Title}: 营销员工号为空，无法分配.(保单号: {bid}) "); 
                return;
            }

            var sid = Convert.ToString(bill[BillSheetColumns.SELLER_ID]).Trim();
            if(string.IsNullOrEmpty(sid))
            {
                LogError(bill, $"{step.Title}: 营销员工号为空，无法分配.(保单号: {bid}) ");
                return;
            }
             
            var mem = mMembers.FirstOrDefault(a => a.ID.Equals(sid, StringComparison.CurrentCultureIgnoreCase) && a.Position.Equals( PositionNames.SERVICE, StringComparison.CurrentCultureIgnoreCase));
            if (mem == null)
            {
                LogError(bill, $"{step.Title}: 未能找到工号为 {sid} 的客服专员，无法分配. (保单号： {bid})");
            }
            else
            {
                AssignTo(bill,mem);
                LogHistory(bill, step, mem); 
            }
            bill[BillSheetColumns.SYS_FILTER] = step.Title;

        }

        private void AssignTo(DataRow bill, Member member)
        {
            bill[BillSheetColumns.CURRENT_SERVICE_NAME] = member.Name;
            bill[BillSheetColumns.CURRENT_SERVICE_ID] = member.ID;
            bill[BillSheetColumns.SYS_SERVICE_AREA] = member.Area;
        }
        private void LogError(DataRow bill, string error)
        {
            mLogger.Error(error);
            var log = Convert.ToString(bill[BillSheetColumns.SYS_ERROR]);
            if(string.IsNullOrEmpty(log))
            {
                log = error;
            }
            else
            {
                log = $"{log}\r\n{error}";
            }

            bill[BillSheetColumns.SYS_ERROR] = log;
        }
        private void LogHistory(DataRow bill, Step step, Member mem)
        { 
            var log = Convert.ToString(bill[BillSheetColumns.SYS_HISTORY]);
            var service = "";
            if(!bill.IsNull(BillSheetColumns.SYS_SERVICE))
            {
                service =  Convert.ToString( bill[BillSheetColumns.SYS_SERVICE]);
            }
            var newlog = $"{step.Title} -> {mem.Name} ({mem.ID} - {mem.Area}) ,  分给{step.DispatchType.GetDescription()}";
            var history = "";
            if(string.IsNullOrEmpty(log))
            {
                if(string.IsNullOrEmpty(service))
                {
                    history = newlog;
                }
                else
                {
                    history = $"原始客服专员： {service}\r\n{newlog}";
                }
                
            }
            else
            {
                history = $"{log}\r\n{newlog}";
            }

            bill[BillSheetColumns.SYS_HISTORY] = history;
        }


        private string GetBillID(DataRow row)
        {
            var bid = "";
            if(row.Table.Columns.Contains(BillSheetColumns.BILL_ID))
            {
                bid = (string) row[BillSheetColumns.BILL_ID];
            }
            if (row.Table.Columns.Contains(BillSheetColumns.BILL_ID2))
            {
                bid = (string)row[BillSheetColumns.BILL_ID2];
            }
            return bid;
        }

        class MemIndex
        {
            public int Index { get; set; }
            public string ServiceID { get; set; }
        }

        private bool GetHitBySellerIsService(Filter filter, DataRow bill)
        {
            bool b = false;
            var bid = GetBillID(bill);
            if (bill.IsNull(BillSheetColumns.SELLER_ID))
            { 
                return b;
            }

            var sid = Convert.ToString(bill[BillSheetColumns.SELLER_ID]).Trim();
            if (string.IsNullOrEmpty(sid))
            { 
                return b;
            }

            var mem = mMembers.FirstOrDefault(a => a.ID.Equals(sid, StringComparison.CurrentCultureIgnoreCase) && a.Position.Equals(PositionNames.SERVICE, StringComparison.CurrentCultureIgnoreCase));
            b = mem != null;
            return b;
        }

        private bool GetHitByDifferentArea(Filter filter, DataRow bill)
        {
            bool b = false;
            var clientArea = (string) bill[BillSheetColumns.CLIENT_AREA];

            if(string.IsNullOrEmpty(clientArea))
                return b;

            var sellerId = (string) bill[BillSheetColumns.SELLER_ID];
            var seller = mMembers.FirstOrDefault(a => a.ID.Equals(sellerId.Trim(), StringComparison.CurrentCultureIgnoreCase));
            if(seller == null)
            {
                var bid = GetBillID(bill);
                LogError(bill, $"未能找到工号为 {sellerId} 的营销员，无法识别跨区单. (保单号： {bid})");
                return b;
            }
            b = clientArea != seller.Area; 

            return b;
        }


    }
}
