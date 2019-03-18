using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using InsuranceCompareTool.Models;
namespace InsuranceCompareTool.Services
{
    public static class DataValidateService
    {
        public static void ValidateDate(List<Relation> relations, List<Bill> bills, List<Member> members )
        {
             
            List<string> pl = new List<string>(); 

            foreach (var relation in relations)
            {
                if (!string.IsNullOrEmpty(relation.ServiceID))
                {
                    if (!members.Any(a => a.ID.Equals(relation.ServiceID)))
                    {
                        if(!pl.Contains(relation.ServiceID))
                        {
                            pl.Add(relation.ServiceID);
                        } 
                    }
                }
            }
 
            foreach(var bill in bills)
            {
                if(!string.IsNullOrEmpty(bill.LastServiceID))
                { 
                    if(!members.Any(a => a.ID.Equals(bill.LastServiceID)))
                    {
                        if (!pl.Contains(bill.LastServiceID))
                        {
                            pl.Add(bill.LastServiceID);
                        } 
                    } 
                } 
            }

            if(pl.Count > 0  )
            {
                string pls = string.Join(", ", pl); 
                string str = (pl.Count > 0 ? $"人员表缺少人员： {pls}" : "")   + "\r\n\r\n(相关内容已复制到剪贴板)";

                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    System.Windows.Clipboard.SetText(str);
                });
                
                throw new Exception(str);
            }
 
        }
    }
}
