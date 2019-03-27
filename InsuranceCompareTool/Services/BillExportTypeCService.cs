using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using InsuranceCompareTool.Core;
using InsuranceCompareTool.Models;
using NPOI.HSSF.UserModel;
namespace InsuranceCompareTool.Services
{
    public class BillExportTypeCService
    {
        #region Fields
        private static BillExportTypeCService Singleton;
        #endregion
        #region Public Methods
        public static BillExportTypeCService CreateInstance()
        {
            if(Singleton == null)
                Singleton = new BillExportTypeCService();
            return Singleton;
        }
        public void Export(string targetPath, List<Bill> bills, List<Member> members)
        {
            
            InitDirAndClearFiles(targetPath);

            var areas = bills.Select(a => a.SellArea).Distinct().ToArray();
            foreach(string area in areas)
            {
                var areaBills = bills.Where(a => a.SellArea.Equals(area)).ToList();
                var serIds = areaBills.Select(a => a.CurrentServiceID).Distinct().ToArray();
                var writer = new ServiceBillsTableWriterB();
                var statistics = new ServiceBillsStatisticsWriter();
                var virServices = new List<Member>();
                var virSellers = new List<string>();
                //导出正常的
                foreach (var serId in serIds)
                { 
                    if (string.IsNullOrEmpty(serId))
                        continue;
                    var member = members.FirstOrDefault(a => a.ID.Equals(serId));
                    if(member == null)
                    {
                        throw new Exception($"缺少工号为{serId}的客服专员");
                    }

                    if(member.VirtualMember == true)
                    {
                        var virBills = areaBills.Where(a => a.CurrentServiceID.Equals(member.ID)).Select(a=>a.SellerID).Distinct().ToArray();
                        virSellers.AddRange(virBills);
                        virServices.Add(member);
                        continue;
                    }
                       
                    var memberName = member.Name;
                    var serBills = bills.Where(a => a.CurrentServiceID.Equals(serId)).OrderBy(a => a.PayDate).ThenBy(a => a.CustomerName).ThenBy(a => a.ID).ToList();
                    var count = serBills.Count;
                    var sum = serBills.Sum(a => a.Price);
                    var title =  $"{area} - {serId} - {memberName} \t 合计： {count} 单 , 合计保费：{sum} 元";
                    writer.WriteBills(member?.Name, title,area, serBills, WriteType.Service);
                    statistics.WriteLine(serId,memberName, count.ToString(),sum.ToString("N") );
                }

                //导出营销员的 
                var sellers = members.Where(a => a.Area.Equals(area) && a.Reportable == true).Select(a => a.ID).Distinct().ToList();
                foreach (var sellerId in sellers)
                {
                    if (string.IsNullOrEmpty(sellerId))
                        continue;
                    var member = members.FirstOrDefault(a => a.ID.Equals(sellerId));
                    if (member == null)
                    {
                        throw new Exception($"缺少工号为{sellerId}的营销员");
                    }
                    var sellerBills = bills.Where(a => a.SellerID.Equals(sellerId)).OrderBy(a => a.PayDate).ThenBy(a => a.CustomerName).ThenBy(a => a.ID).ToList();
                    if(sellerBills.Count <= 0)
                    {
                        continue;
                    }
                    var count = sellerBills.Count;
                    var sum = sellerBills.Sum(a => a.Price);
                    var title = $"{area} - {sellerId} - {member.Name} \t 合计： {count} 单 , 合计保费：{sum} 元";
                    writer.WriteBills(member?.Name, title,area, sellerBills, WriteType.Seller);
                    statistics.WriteLine(sellerId, "@" + member.Name,count.ToString(), sum.ToString("N"));
                }

                //导出虚拟工号的
                foreach (var sellerId in virSellers)
                {
                    if (string.IsNullOrEmpty(sellerId))
                        continue;
                    var member = members.FirstOrDefault(a => a.ID.Equals(sellerId));
                    if (member == null)
                    {
                        throw new Exception($"缺少工号为{sellerId}的营销员");
                    }
                    var sellerBills = bills.Where(a => a.SellerID.Equals(sellerId)).OrderBy(a => a.PayDate).ThenBy(a => a.CustomerName).ThenBy(a => a.ID).ToList();
                    if (sellerBills.Count <= 0)
                    {
                        continue;
                    }
                    var count = sellerBills.Count;
                    var sum = sellerBills.Sum(a => a.Price);
                    var title = $"{area} - {sellerId} - {member.Name} \t 合计： {sellerBills.Count} 单 , 合计保费：{sellerBills.Sum(a => a.Price)} 元";
                    writer.WriteBills(member?.Name, title,area, sellerBills, WriteType.Virtual);
                    statistics.WriteLine(sellerId, "@" + member.Name, count.ToString(), sum.ToString("N"));
                }
                //统计虚拟工号的
                foreach(var ser in virServices)
                {
                    var serBills = bills.Where(a => a.CurrentServiceID.Equals(ser.ID));
                    var count = serBills.Count();
                    var sum = serBills.Sum(a => a.Price); 
                    statistics.WriteLine(ser.ID, "$ " + ser.Name, count.ToString(), sum.ToString("N"));
                }

                areaBills = areaBills.OrderBy(a=>a.CurrentServiceName).ThenBy(a => a.PayDate).ThenBy(a => a.CustomerName).ThenBy(a => a.ID).ToList();
                var fileName = $"{targetPath}\\{area}-{DateTime.Now.AddMonths(1).ToString("yyyy-MM")}.xlsx";
                var alltitle = $"{area} \t 合计： {areaBills.Count} 单 , 合计保费：{areaBills.Sum(a => a.Price)} 元";
                writer.WriteBills("总清单", alltitle, area,areaBills, WriteType.All);
                writer.Save(fileName);
                statistics.WriteLine("", "合计", areaBills.Count.ToString(), areaBills.Sum(a => a.Price).ToString("N"));
                statistics.Save($"{targetPath}\\{area}-{DateTime.Now.AddMonths(1).ToString("yyyy-MM")}-统计.xlsx");
            } 
        }

        #endregion
        #region Private or Protect Methods
        private void InitDirAndClearFiles(string targetPath)
        {
            var dir = new DirectoryInfo(targetPath);
            if(!dir.Exists)
            {
                dir.Create();
            }
            else
            {
                //var files = dir.GetFiles();
                //foreach(var f in files)
                //    f.Delete();
            }
        }
        #endregion
    }
}