using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InsuranceCompareTool.Domain;
using InsuranceCompareTool.Models;
using InsuranceCompareTool.Models.Dispatch;
namespace InsuranceCompareTool.Core
{

    /// <summary>
    /// 指派订单类
    /// </summary>
    public class BillAssignHelper
    {
        class ColumnDefine
        {
            public string Name { get; set; }
            public Type Type { get; set; }
        }
        private const string TARGET_SERVICE_ID = BillSheetColumns.ASSIGNED_SERVICE_ID;
        private const string TARGET_SERVICE_NAME = BillSheetColumns.ASSIGNED_SERVICE_NAME;
          
        private readonly List<ColumnDefine> mSystemColumns = new List<ColumnDefine>()
        {
            new ColumnDefine(){Name = BillSheetColumns.SYS_FILTER, Type = typeof(string)},
            new ColumnDefine(){Name = BillSheetColumns.SYS_HISTORY, Type = typeof(string)},
            new ColumnDefine(){Name = BillSheetColumns.SYS_FINISHED, Type = typeof(bool)},
            //BillSheetColumns.SYS_FILTER, "System_History", BillSheetColumns.SYS_FINISHED
        };
        private void UpdateDataTable(DataTable table)
        {
            foreach(var column in mSystemColumns)
            {
                if(!table.Columns.Contains(column.Name))
                {
                    table.Columns.Add(column.Name, column.Type );
                }
            }

            foreach(DataRow row in table.Rows)
            {
                row[BillSheetColumns.SYS_FINISHED] = false ;
            }
        }
        public void AssignBills(DataTable dt, Project project)
        { 
            UpdateDataTable(dt);

            foreach (var step in project.Steps)
            {
                foreach(DataRow row in dt.Rows)
                {
                    var hit = GetHitValue(row,step.Filters); 

                    if(hit )
                    {
                        AssignToSomebody(row, step);
                    }
                }
            } 
        }
        /// <summary>
        /// 订单命中测试
        /// </summary>  
        private bool GetHitValue(DataRow bill, IEnumerable<Filter> filters)
        {
            if((bool)bill[BillSheetColumns.SYS_FINISHED] == true)
                return false;
            var hit = new List<bool>();

            foreach (var filter in filters)
            {
                switch (filter.Column)
                {
                    case LogicColumnName.ServiceStatus:
                    {
                        hit.Add(GetHitByServiceStatusFilter(filter, bill));
                        break;
                    }
                    case LogicColumnName.ServiceID:
                    {
                        hit.Add(GetHitByServiceID(filter, bill));
                        break;
                    }
                    case LogicColumnName.BillID:
                    {
                        hit.Add(GetHitByBillID(filter, bill));
                        break;
                    }
                    case LogicColumnName.CurrentPrice:
                    {
                        hit.Add(GetHitByCurrentPrice(filter,bill));
                        break;
                    } 
                    default:
                    {
                        break;
                    }
                }
            }
            var hited =  hit.All(a => a == true);
            if(hited)
            {
                bill[BillSheetColumns.SYS_FINISHED] = true;
            }
            return hited;
        }

        private bool GetHitByServiceStatusFilter(Filter filter, DataRow bill)
        {
            bool b = false;
            
            if(filter.OperatorType == OperatorType.Equal)
            {
                b =  (string) bill[BillSheetColumns.SYS_SERVICE_STATUS] == filter.Value;
            }
            else if(filter.OperatorType == OperatorType.Contain)
            {
                b = ((string)bill[BillSheetColumns.SYS_SERVICE_STATUS]).Contains(filter.Value);
            }
            return b;
        }
        private bool GetHitByServiceID(Filter filter, DataRow bill)
        {
            bool b = false;

            if (filter.OperatorType == OperatorType.Equal)
            {
                b = (string)bill[BillSheetColumns.CURRENT_SERVICE_ID] == filter.Value;
            }
            else if (filter.OperatorType == OperatorType.Contain)
            {
                b = ((string)bill[BillSheetColumns.CURRENT_SERVICE_ID]).Contains(filter.Value);
            }
            return b;
        }
        private bool GetHitByBillID(Filter filter, DataRow bill)
        {
            return true; 
        }
        private bool GetHitByCurrentPrice(Filter filter, DataRow bill)
        {
            return true;
        }
        

        /// <summary>
        /// 指派订单给某人
        /// </summary> 
        /// <param name="step"></param>
        private void AssignToSomebody(DataRow row, Step step)
        {
            row[TARGET_SERVICE_ID] = step.DispatchDesignated;
            row[TARGET_SERVICE_NAME] = step.DispatchDesignated;
            row[BillSheetColumns.SYS_FILTER] = step.Title;
        }
    }
}
